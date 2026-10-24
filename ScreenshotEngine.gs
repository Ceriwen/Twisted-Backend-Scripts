/**
 * ScreenshotEngine.gs - Automated Gemini OCR Batch Ingestion Pipeline
 * TWST Roster Management System
 */

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
const GEMINI_MODEL = "gemini-3.8-flash"; // Matches your active API model configuration
const BATCH_SIZE = 5;

/**
 * Strips alliance tags like [TWST] and trailing squad tags like " D1", " A2", " C4b"
 * so "Soulcollector D1" cleanly matches "Soulcollector".
 */
function stripSquadTag(str) {
  if (!str) return "";
  return str.toString()
    .replace(/\[.*?\]\s*/g, "")
    .replace(/\(.*?\)\s*/g, "")
    .replace(/\s+[A-Da-d][1-4][a-bA-B]?$/i, "")
    .trim();
}

/**
 * Resolves a raw OCR name against CONFIG.MEMBER_ALIASES and strips squad tags.
 */
function resolveCandidateName_(rawName) {
  if (!rawName) return "";
  let trimmed = rawName.toString().trim();

  if (typeof CONFIG !== "undefined" && CONFIG.MEMBER_ALIASES) {
    if (CONFIG.MEMBER_ALIASES[trimmed]) return CONFIG.MEMBER_ALIASES[trimmed];
    let stripped = stripSquadTag(trimmed);
    if (CONFIG.MEMBER_ALIASES[stripped]) return CONFIG.MEMBER_ALIASES[stripped];

    let lowerTrim = trimmed.toLowerCase();
    let lowerStrip = stripped.toLowerCase();
    let aliasKey = Object.keys(CONFIG.MEMBER_ALIASES).find(
      k => k.trim().toLowerCase() === lowerTrim || k.trim().toLowerCase() === lowerStrip
    );
    if (aliasKey) return CONFIG.MEMBER_ALIASES[aliasKey];
  }

  return stripSquadTag(trimmed);
}

/**
 * Main ingestion handler for folder processing (Fast Option A - Zero Cell Highlighting).
 */
function processFolderDirectly(headingLabelsArray, returnStats = false, ignoreRanks = false) {
  const START_TIME = Date.now();
  const MAX_RUNTIME_MS = 300000; // 5-minute threshold
  let timeoutTriggered = false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  // --- PERSISTENT IMPORT BACKUP CREATION (FOR ROLLBACK LAST IMPORT) ---
  try {
    let tempSnap = ss.getSheetByName("_ImportBackup");
    if (!tempSnap) {
      tempSnap = ss.insertSheet("_ImportBackup");
    } else {
      tempSnap.clear();
    }
    let srcRange = sheet.getDataRange();
    if (srcRange.getNumRows() > 0 && srcRange.getNumColumns() > 0) {
      srcRange.copyTo(tempSnap.getRange(1, 1));
    }
    SpreadsheetApp.flush();
    tempSnap.hideSheet(); // Hidden so it doesn't clutter tabs, but preserved for Rollback!
  } catch (e) {
    debugSheet.appendRow([new Date(), "⚠️ Backup Error: " + e.message]);
  }

  // Ensure clean table structure before parsing
  if (typeof cleanAndSortRoster === "function") {
    cleanAndSortRoster(sheet);
  }

  let totalProcessedFiles = 0;
  let cumulativeMatchedRows = new Set();
  let auditMatched = [];
  let auditNewMembers = [];
  let auditMissing = [];
  let rankDiscrepancies = [];

  for (let h = 0; h < headingLabelsArray.length; h++) {
    let currentLabel = headingLabelsArray[h];
    let folderId = CONFIG.FOLDER_MAPPINGS[currentLabel];
    if (!folderId || folderId.startsWith("YOUR_")) continue;

    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      debugSheet.appendRow([new Date(), `⚠️ Folder lookup error for ${currentLabel}: ${e.message}`]);
      continue;
    }

    let filesIter = targetFolder.getFiles();
    let allFiles = [];
    while (filesIter.hasNext()) {
      let f = filesIter.next();
      let mime = f.getMimeType();
      if (mime.includes("image") || mime.includes("octet-stream") || f.getName().toLowerCase().endsWith(".jpg") || f.getName().toLowerCase().endsWith(".png")) {
        allFiles.push(f);
      }
    }

    if (allFiles.length === 0) continue;

    let targetProcessedSubfolder = null;
    try {
      let master = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID);
      let subName = CONFIG.PROCESSED_SUBFOLDER_NAMES[currentLabel] || currentLabel;
      let subIter = master.getFoldersByName(subName);
      targetProcessedSubfolder = subIter.hasNext() ? subIter.next() : master.createFolder(subName);
    } catch (e) {
      debugSheet.appendRow([new Date(), `⚠️ Processed folder lookup error: ${e.message}`]);
    }

    let isStrengthRanking = (currentLabel === "Strength Ranking - Power");
    let rosterRange = findRosterRange(sheet);
    let startRow = rosterRange.startRow;
    let totalCurrentRows = rosterRange.rowCount;
    let sheetData = sheet.getRange(startRow, 1, totalCurrentRows, rosterRange.colCount).getValues();
    let headers = sheetData[0].map(item => item.toString().trim().toLowerCase());

    let nameColIdx = headers.findIndex(item => item.includes("member"));
    let metricColIdx = isStrengthRanking 
      ? headers.findIndex(item => item.includes("power")) 
      : (currentLabel === "Strength Ranking - Donations"
          ? headers.findIndex(item => item.includes("donation"))
          : headers.findIndex(item => item === currentLabel.toLowerCase() || item.includes(currentLabel.toLowerCase())));
    let rankColIdx = headers.findIndex(item => item === "rank");

    if (nameColIdx === -1) {
      debugSheet.appendRow([new Date(), `❌ 'Member' column not found in sheet.`]);
      continue;
    }

    for (let bIdx = 0; bIdx < allFiles.length; bIdx += BATCH_SIZE) {
      if (Date.now() - START_TIME > MAX_RUNTIME_MS) {
        timeoutTriggered = true;
        updateQueueVisualStatus(`⏱️ 5m limit reached at file ${bIdx + 1}. Auto-resuming in 1m...`, "#fff2cc");
        debugSheet.appendRow([new Date(), `⚠️ Safety limit reached at file ${bIdx + 1}. Auto-scheduling self-resumption...`]);
        try {
          let triggers = ScriptApp.getProjectTriggers();
          triggers.forEach(t => {
            if (t.getHandlerFunction() === "processNextAvailableFolder") {
              ScriptApp.deleteTrigger(t);
            }
          });
          ScriptApp.newTrigger("processNextAvailableFolder")
            .timeBased()
            .after(60 * 1000)
            .create();
        } catch (trigErr) {
          debugSheet.appendRow([new Date(), `⚠️ Trigger Error: ${trigErr.message}`]);
        }
        break;
      }

      let batchFiles = allFiles.slice(bIdx, bIdx + BATCH_SIZE);
      let startNum = bIdx + 1;
      let endNum = bIdx + batchFiles.length;
      let batchLabelStr = `[${currentLabel}] Batch (${startNum}-${endNum} of ${allFiles.length})`;

      updateQueueVisualStatus(`⚡ Scanning ${batchLabelStr}`, "#fff2cc");
      debugSheet.appendRow([new Date(), `Scanning ${batchLabelStr}`]);

      let ocrResults = [];
      try {
        ocrResults = parseLeaderboardBatchWithGemini(batchFiles, currentLabel);
      } catch (geminiErr) {
        debugSheet.appendRow([new Date(), `❌ Gemini Parse Error on ${batchLabelStr}: ${geminiErr.message}`]);
        continue;
      }

      ocrResults.forEach(item => {
        if (!item || !item.name) return;
        let cleanCandidateName = resolveCandidateName_(item.name);
        let scoreVal = cleanNumericValues(item.score);
        let rankVal = item.rank ? item.rank.toString().trim().toUpperCase() : "";

        let candBase = normalizeString(stripSquadTag(cleanCandidateName));
        let matchedRowIdx = -1;

        // 1. Exact & Squad-Tag-Stripped Match (e.g. "Soulcollector D1" === "Soulcollector")
        for (let r = 1; r < sheetData.length; r++) {
          let rowMemberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
          if (!rowMemberName || rowMemberName.startsWith("---")) continue;
          let rowBase = normalizeString(stripSquadTag(rowMemberName));
          if (rowBase === candBase && candBase.length > 0) {
            matchedRowIdx = r;
            break;
          }
        }

        // 2. Prefix / Substring Stem Match for minor OCR cutoffs or trailing numbers
        if (matchedRowIdx === -1 && candBase.length >= 5) {
          for (let r = 1; r < sheetData.length; r++) {
            let rowMemberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
            if (!rowMemberName || rowMemberName.startsWith("---")) continue;
            let rowBase = normalizeString(stripSquadTag(rowMemberName));
            if (rowBase.length >= 5 && (rowBase.startsWith(candBase) || candBase.startsWith(rowBase))) {
              matchedRowIdx = r;
              break;
            }
          }
        }

        if (matchedRowIdx > -1) {
          let actualRow = startRow + matchedRowIdx;
          cumulativeMatchedRows.add(actualRow);

          let matchedRosterName = sheetData[matchedRowIdx][nameColIdx].toString().trim();
          if (metricColIdx > -1) {
            sheetData[matchedRowIdx][metricColIdx] = scoreVal;
          }

          auditMatched.push({ name: matchedRosterName, score: scoreVal, raw: item.name, category: currentLabel });

          // Power Ranking Discrepancy Protection (flags differences in sidebar instead of overwriting)
          if (rankColIdx > -1 && rankVal && isStrengthRanking) {
            let existingRank = sheetData[matchedRowIdx][rankColIdx] ? sheetData[matchedRowIdx][rankColIdx].toString().trim().toUpperCase() : "";
            if (existingRank && existingRank !== rankVal && existingRank !== "R5" && existingRank !== "R4") {
              if (!rankDiscrepancies.some(d => d.row === actualRow)) {
                rankDiscrepancies.push({ row: actualRow, name: matchedRosterName, sheetRank: existingRank, ocrRank: rankVal, category: currentLabel });
              }
            } else if (!existingRank && !ignoreRanks) {
              sheetData[matchedRowIdx][rankColIdx] = rankVal;
            }
          }
        } else {
          // New / Unmatched Member -> Append to bottom of sheetData
          let newRow = new Array(rosterRange.colCount).fill("");
          if (rankColIdx > -1) newRow[rankColIdx] = rankVal || "R3";
          newRow[nameColIdx] = cleanCandidateName;
          if (metricColIdx > -1) newRow[metricColIdx] = scoreVal;

          sheetData.push(newRow);
          totalCurrentRows++;
          let newActualRow = startRow + sheetData.length - 1;
          cumulativeMatchedRows.add(newActualRow);

          auditNewMembers.push({ name: cleanCandidateName, score: scoreVal, raw: item.name, row: newActualRow, category: currentLabel });
        }
      });

// Fast Raw Value Write Only (NO paintBatchFormatting!)
      sheet.getRange(startRow, 1, sheetData.length, rosterRange.colCount).setValues(sheetData);

      // Attach '⚠️ New Member' hover note to any newly added members in this batch
      auditNewMembers.forEach(nm => {
        if (nm.row) {
          let nameCell = sheet.getRange(nm.row, nameColIdx + 1);
          let curNote = nameCell.getNote();
          if (!curNote || !curNote.includes("⚠️ New Member")) {
            nameCell.setNote(curNote ? `${curNote}\n⚠️ New Member` : "⚠️ New Member");
          }
        }
      });

      SpreadsheetApp.flush();

      batchFiles.forEach(file => {
        try {
          if (targetProcessedSubfolder) file.moveTo(targetProcessedSubfolder);
          else file.setTrashed(true);
        } catch (e) {}
      });

      totalProcessedFiles += batchFiles.length;
    }

    // Identify existing roster members who were missing from screenshots (0 score counts as valid!)
    if (!timeoutTriggered) {
      let memberNotes = sheet.getRange(startRow + 1, nameColIdx + 1, sheetData.length - 1, 1).getNotes();
      for (let r = 1; r < sheetData.length; r++) {
        let memberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
        if (!memberName || memberName.startsWith("---") || memberName.startsWith("⚠️") || memberName.toLowerCase() === "member") continue;

        let actualRow = startRow + r;
        let scoreCell = metricColIdx > -1 ? sheetData[r][metricColIdx] : "";
        let isMissingScore = (scoreCell === "" || scoreCell === null || scoreCell === undefined);
        let isMatchedInRun = cumulativeMatchedRows.has(actualRow);
        let noteStr = (memberNotes[r - 1] && memberNotes[r - 1][0]) ? memberNotes[r - 1][0] : "";
        let exempt = (typeof isExemptNote === "function") ? isExemptNote(noteStr) : false;

        if (!isMatchedInRun && isMissingScore && !exempt) {
          auditMissing.push({ name: memberName, score: "Missing", category: currentLabel, row: actualRow });
        }
      }
    }
  }

  let primaryLabel = headingLabelsArray[0] || "Import";

  // Only write to 'Import Audit' if files were actually processed in this run!
  if (totalProcessedFiles > 0) {
    try {
      writeToImportAuditTab(primaryLabel, auditMatched, auditNewMembers, auditMissing, rankDiscrepancies, timeoutTriggered);
    } catch (e) {
      debugSheet.appendRow([new Date(), `⚠️ Audit Tab Write Error: ${e.message}`]);
    }
  }

  // Update F1 status immediately upon completion
  if (!timeoutTriggered) {
    try {
      let triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(t => {
        if (t.getHandlerFunction() === "processNextAvailableFolder") {
          ScriptApp.deleteTrigger(t);
        }
      });
    } catch (e) {}

    updateQueueVisualStatus(`✅ Import Complete: ${primaryLabel} (${auditMatched.length} matched, ${auditMissing.length} missing)`, "#d9ead3");
    debugSheet.appendRow([new Date(), `✅ Completed "${primaryLabel}": ${auditMatched.length} matched, ${auditNewMembers.length} new, ${auditMissing.length} missing.`]);
  }

  // Pop open the right-hand Audit Sidebar (works when triggered manually from desktop UI)
  try {
    showImportAuditSidebar(primaryLabel, auditMatched, auditNewMembers, auditMissing, rankDiscrepancies, timeoutTriggered);
  } catch (uiErr) {}

  if (returnStats) {
    return { timedOut: timeoutTriggered, processedCount: totalProcessedFiles };
  }
}

/**
 * CUMULATIVE MULTI-DAY AUDIT TAB WRITER:
 * Preserves audit history from earlier folders/days processed by the Auto-Watcher!
 * Only replaces rows if the exact same category is re-imported (unless resuming from a timeout).
 */
function writeToImportAuditTab(currentLabel, matched, newMembers, missing, rankChanges, isContinuation) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName("Import Audit");
  let existingRows = [];

  if (!auditSheet) {
    auditSheet = ss.insertSheet("Import Audit");
  } else {
    let lastRow = auditSheet.getLastRow();
    if (lastRow > 1) {
      let data = auditSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      // Keep all rows from OTHER days/categories (or keep everything if this is a timeout continuation)
      existingRows = data.filter(row => {
        let rowCat = String(row[1] || "").trim();
        if (isContinuation) {
          // If resuming same folder, keep matched/new from Batch 1, but refresh missing list
          let rowStatus = String(row[2] || "");
          return rowCat !== currentLabel || !rowStatus.includes("Missing");
        }
        return rowCat !== currentLabel;
      });
    }
    auditSheet.clearContents();
  }

  auditSheet.getRange(1, 1, 1, 6)
    .setValues([["Timestamp", "Category", "Status", "Player Name", "OCR Score / Value", "Details / Notes"]])
    .setFontWeight("bold")
    .setBackground("#d9ead3");
  auditSheet.setFrozenRows(1);

  let timestamp = new Date();
  let newRows = [];

  rankChanges.forEach(r => newRows.push([timestamp, currentLabel, "⚠️ Rank Discrepancy", r.name, "", `Sheet: ${r.sheetRank} vs Screenshot: ${r.ocrRank} (Row ${r.row})`]));
  missing.forEach(m => newRows.push([timestamp, currentLabel, "❌ Missing from Screenshots", m.name, "", `Not detected in ${currentLabel} screenshots (Row ${m.row || "-"})`]));
  newMembers.forEach(m => newRows.push([timestamp, currentLabel, "🟡 New / Unmatched Added", m.name, m.score, `Raw OCR: "${m.raw}" (Appended at Row ${m.row})`]));
  matched.forEach(m => newRows.push([timestamp, currentLabel, "🟢 Matched", m.name, m.score, `Raw OCR: "${m.raw}"`]));

  // Prepend newest category rows at the top, followed by earlier days' audit logs
  let combinedRows = newRows.concat(existingRows);
  if (combinedRows.length > 0) {
    auditSheet.getRange(2, 1, combinedRows.length, 6).setValues(combinedRows);
  }
}

/**
 * Clears the 'Import Audit' tab on demand or during Sunday Reset.
 */
function clearImportAuditLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName("Import Audit");
  if (auditSheet) {
    auditSheet.clearContents();
    auditSheet.getRange(1, 1, 1, 6)
      .setValues([["Timestamp", "Category", "Status", "Player Name", "OCR Score / Value", "Details / Notes"]])
      .setFontWeight("bold")
      .setBackground("#d9ead3");
  }
  ss.toast("Import Audit history cleared.", "🧹 Audit Cleared", 4);
}

/**
 * Builds a complete multi-day audit dataset by reading the 'Import Audit' tab
 * AND scanning 'Data Import' vs 'Last Week'.
 * - Uses full alias + stem matching so existing members (McNastyyyyy) aren't flagged as New.
 * - Excludes members with existing strike notes (DelmiraSera - Pending Removal) from being flagged as New.
 * - Automatically ensures true New Members (Onyx A, NeuroToxic) have the '⚠️ New Member' hover note on their name cell!
 */
function getFullWeekAuditData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Import Audit");
  const dataSheet = ss.getSheetByName("Data Import");
  const lastWeekSheet = ss.getSheetByName("Last Week");

  let byCategory = {};
  let categoryOrder = [];
  let knownNewMembersSet = new Set();

  const ensureCat = (cat, ts) => {
    if (!byCategory[cat]) {
      byCategory[cat] = { matched: [], missing: [], newMembers: [], rankChanges: [], timestamp: ts || "" };
      categoryOrder.push(cat);
    }
  };

  // Helper: canonical normalized base name (applies MEMBER_ALIASES + strips squad tags like "A", "D1", etc.)
  const toCanonicalNorm = (raw) => {
    if (!raw) return "";
    let resolved = resolveCandidateName_(raw);
    // Also strip single trailing squad letter like "Onyx A" -> "Onyx" only for comparison if needed,
    // or standard squad tag
    return normalizeString(stripSquadTag(resolved));
  };

  // 0. Read 'Last Week' roster names using findRosterRange + full alias resolution
  let lastWeekRosterNorms = [];
  if (lastWeekSheet && lastWeekSheet.getLastRow() >= 2) {
    let lwRange = findRosterRange(lastWeekSheet);
    let lwData = lastWeekSheet.getRange(lwRange.startRow, 1, lwRange.rowCount, lwRange.colCount).getValues();
    let lwHeaders = lwData[0].map(h => h.toString().trim().toLowerCase());
    let lwNameIdx = lwHeaders.findIndex(h => h.includes("member"));
    if (lwNameIdx === -1) lwNameIdx = 1;

    for (let r = 1; r < lwData.length; r++) {
      let n = lwData[r][lwNameIdx] ? lwData[r][lwNameIdx].toString().trim() : "";
      if (!n || n.startsWith("---") || n.startsWith("⚠️")) continue;
      let cNorm = toCanonicalNorm(n);
      if (cNorm) lastWeekRosterNorms.push(cNorm);
    }
  }

  // Helper: check if a member existed on Last Week (exact canonical match OR 5+ char prefix stem match like McNasty / McNastyyyyy)
  const wasInLastWeek = (memberName) => {
    if (lastWeekRosterNorms.length === 0) return true;
    let cand = toCanonicalNorm(memberName);
    if (!cand) return true;
    for (let i = 0; i < lastWeekRosterNorms.length; i++) {
      let lw = lastWeekRosterNorms[i];
      if (lw === cand) return true;
      if (lw.length >= 5 && cand.length >= 5 && (lw.startsWith(cand) || cand.startsWith(lw))) return true;
    }
    return false;
  };

  // 1. Read logged entries from 'Import Audit' tab
  if (auditSheet && auditSheet.getLastRow() > 1) {
    let rows = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 6).getValues();
    rows.forEach(r => {
      let ts = r[0] ? Utilities.formatDate(new Date(r[0]), ss.getSpreadsheetTimeZone(), "MM/dd h:mm a") : "";
      let cat = String(r[1] || "").trim();
      let status = String(r[2] || "").trim();
      let name = String(r[3] || "").trim();
      let score = r[4];
      let details = String(r[5] || "").trim();
      if (!cat || !name) return;

      ensureCat(cat, ts);
      let normName = toCanonicalNorm(name);

      if (status.includes("New")) {
        knownNewMembersSet.add(normName);
        let m = details.match(/Row\s*(\d+)/i);
        byCategory[cat].newMembers.push({ category: cat, name: name, score: score, row: m ? parseInt(m[1], 10) : -1, raw: details });
      } else if (status.includes("Rank Discrepancy")) {
        let m = details.match(/Sheet:\s*(\S+)\s*vs\s*Screenshot:\s*(\S+)\s*\(Row\s*(\d+)\)/i);
        byCategory[cat].rankChanges.push({
          category: cat,
          name: name,
          sheetRank: m ? m[1] : "",
          ocrRank: m ? m[2] : "",
          row: m ? parseInt(m[3], 10) : -1
        });
      } else if (status.includes("Missing")) {
        let m = details.match(/Row\s*(\d+)/i);
        byCategory[cat].missing.push({ category: cat, name: name, row: m ? parseInt(m[1], 10) : -1 });
      } else if (status.includes("Matched")) {
        byCategory[cat].matched.push({ category: cat, name: name, score: score, raw: details });
      }
    });
  }

  // 2. Live Sheet Scan on 'Data Import'
  if (dataSheet) {
    let rosterRange = findRosterRange(dataSheet);
    let sheetData = dataSheet.getRange(rosterRange.startRow, 1, rosterRange.rowCount, rosterRange.colCount).getValues();
    let headers = sheetData[0].map(h => h.toString().trim().toLowerCase());
    let nameColIdx = headers.findIndex(h => h.includes("member"));
    let rankColIdx = headers.findIndex(h => h === "rank");
    let reasonColIdx = headers.findIndex(h => h.includes("reason") || h.includes("notes"));

    if (nameColIdx > -1 && sheetData.length > 1) {
      let nameRange = dataSheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, sheetData.length - 1, 1);
      let memberNotes = nameRange.getNotes();
      let notesModified = false;

      // Identify true New Members on Data Import & ensure their '⚠️ New Member' hover note is set!
      for (let r = 1; r < sheetData.length; r++) {
        let memberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
        if (!memberName || memberName.startsWith("---") || memberName.startsWith("⚠️") || memberName.toLowerCase() === "member") continue;

        let norm = toCanonicalNorm(memberName);
        let curRank = (rankColIdx > -1 && sheetData[r][rankColIdx]) ? sheetData[r][rankColIdx].toString().trim().toUpperCase() : "R3";
        let reasonStr = (reasonColIdx > -1 && sheetData[r][reasonColIdx]) ? sheetData[r][reasonColIdx].toString() : "";
        let noteStr = (memberNotes[r - 1] && memberNotes[r - 1][0]) ? memberNotes[r - 1][0] : "";

        // Anyone with a Pending Removal or Demoted note, or R1/R2 rank, is an existing veteran—NEVER a New Member!
        let hasStrikeOrVeteranStatus =
          noteStr.includes("Pending Removal") ||
          noteStr.includes("Demoted Last Week") ||
          reasonStr.includes("[Removal]") ||
          reasonStr.includes("[Demotion]") ||
          curRank === "R1" ||
          curRank === "R2";

        if (hasStrikeOrVeteranStatus) {
          knownNewMembersSet.delete(norm);
          continue;
        }

        let isNewFromLastWeek = !wasInLastWeek(memberName);
        let isTaggedNew = noteStr.includes("⚠️ New Member") || reasonStr.includes("[New Member]");

        if (isNewFromLastWeek || isTaggedNew || knownNewMembersSet.has(norm)) {
          knownNewMembersSet.add(norm);

          // Automatically attach '⚠️ New Member' hover note if missing!
          if (!noteStr.includes("⚠️ New Member")) {
            memberNotes[r - 1][0] = noteStr ? `${noteStr}\n⚠️ New Member` : "⚠️ New Member";
            notesModified = true;
          }
        }
      }

      if (notesModified) {
        nameRange.setNotes(memberNotes);
        SpreadsheetApp.flush();
      }

      let checkCols = [
        { label: "Mon", match: h => h === "mon" || h.includes("monday") },
        { label: "Tue", match: h => h === "tue" || h.includes("tuesday") },
        { label: "Wed", match: h => h === "wed" || h.includes("wednesday") },
        { label: "Thu", match: h => h === "thu" || h.includes("thursday") },
        { label: "Fri", match: h => h === "fri" || h.includes("friday") },
        { label: "Sat", match: h => h === "sat" || h.includes("saturday") },
        { label: "Strength Ranking - Donations", match: h => h.includes("donation") },
        { label: "Strength Ranking - Power", match: h => h.includes("power") }
      ];

      // Find the first column where each true New Member has a recorded score/power
      let memberFirstSeenCol = {};
      for (let r = 1; r < sheetData.length; r++) {
        let memberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
        if (!memberName) continue;
        let norm = toCanonicalNorm(memberName);
        if (!knownNewMembersSet.has(norm)) continue;

        for (let c = 0; c < checkCols.length; c++) {
          let cIdx = headers.findIndex(checkCols[c].match);
          if (cIdx > -1 && sheetData[r][cIdx] !== "" && sheetData[r][cIdx] !== null && sheetData[r][cIdx] !== undefined) {
            memberFirstSeenCol[norm] = { label: checkCols[c].label, score: sheetData[r][cIdx], row: rosterRange.startRow + r, name: memberName };
            break;
          }
        }
      }

      // Ensure each true New Member is listed once under the category where they first appeared
      Object.keys(memberFirstSeenCol).forEach(norm => {
        let info = memberFirstSeenCol[norm];
        ensureCat(info.label, "Live Sheet Scan");
        let alreadyListed = Object.values(byCategory).some(catObj =>
          catObj.newMembers.some(nm => toCanonicalNorm(nm.name) === norm)
        );
        if (!alreadyListed) {
          byCategory[info.label].newMembers.push({
            category: info.label,
            name: info.name,
            score: info.score,
            row: info.row,
            raw: "⚠️ New Member"
          });
        }
      });

      checkCols.forEach(colDef => {
        let alreadyInAudit = (byCategory[colDef.label] && byCategory[colDef.label].matched.length > 0);
        if (!alreadyInAudit) {
          let cIdx = headers.findIndex(colDef.match);
          if (cIdx === -1) return;

          let popCount = 0;
          for (let r = 1; r < sheetData.length; r++) {
            let v = sheetData[r][cIdx];
            if (v !== "" && v !== null && v !== undefined) popCount++;
          }
          if (popCount < 10) return;

          ensureCat(colDef.label, "Live Sheet Scan");
          for (let r = 1; r < sheetData.length; r++) {
            let memberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
            if (!memberName || memberName.startsWith("---") || memberName.startsWith("⚠️") || memberName.toLowerCase() === "member") continue;

            let norm = toCanonicalNorm(memberName);
            let actualRow = rosterRange.startRow + r;
            let v = sheetData[r][cIdx];
            let noteStr = (memberNotes[r - 1] && memberNotes[r - 1][0]) ? memberNotes[r - 1][0] : "";
            let exempt = (typeof isExemptNote === "function") ? isExemptNote(noteStr) : false;

            if (v === "" || v === null || v === undefined) {
              if (!exempt && !knownNewMembersSet.has(norm)) {
                byCategory[colDef.label].missing.push({ category: colDef.label, name: memberName, row: actualRow });
              }
            } else {
              let isNewInThisCat = byCategory[colDef.label].newMembers.some(nm => toCanonicalNorm(nm.name) === norm);
              if (!isNewInThisCat) {
                byCategory[colDef.label].matched.push({ category: colDef.label, name: memberName, score: v });
              }
            }
          }
        }
      });

      // Clean up any false-positive New or Missing entries
      Object.keys(byCategory).forEach(cat => {
        byCategory[cat].newMembers = byCategory[cat].newMembers.filter(nm => knownNewMembersSet.has(toCanonicalNorm(nm.name)));
        byCategory[cat].missing = byCategory[cat].missing.filter(m => !knownNewMembersSet.has(toCanonicalNorm(m.name)));
      });
    }
  }

  return { byCategory: byCategory, categoryOrder: categoryOrder };
}

/**
 * Menu / Dropdown Action: Opens the Multi-Day Import Audit & Verify Sidebar anytime!
 */
function openAuditSidebarFromMenu() {
  showImportAuditSidebar(null, null, null, null, null, false);
}

/**
 * Displays the interactive Multi-Day Import Audit & Verification Sidebar on the right side of Google Sheets.
 * Lets you switch between "All Processed Days" or individual days (Mon, Tue, Wed, etc.).
 */
function showImportAuditSidebar(label, matched, newMembers, missing, rankChanges, isPaused) {
  let fullData = getFullWeekAuditData_();
  let byCategory = fullData.byCategory;
  let categoryOrder = fullData.categoryOrder;

  // Ensure current run's data is represented if passed directly
  if (label && matched) {
    byCategory[label] = {
      matched: matched.map(m => ({ category: label, name: m.name, score: m.score })),
      missing: (missing || []).map(m => ({ category: label, name: m.name, row: m.row })),
      newMembers: (newMembers || []).map(m => ({ category: label, name: m.name, score: m.score, row: m.row })),
      rankChanges: (rankChanges || []).map(r => ({ category: label, name: r.name, sheetRank: r.sheetRank, ocrRank: r.ocrRank, row: r.row })),
      timestamp: "Just now"
    };
    if (!categoryOrder.includes(label)) categoryOrder.unshift(label);
  }

  let defaultSelection = label || (categoryOrder.length > 0 ? "ALL" : "NONE");

  let statusBanner = isPaused
    ? `<div style="background:#fef9c3;border:1px solid #fde047;color:#854d0e;padding:8px;border-radius:6px;margin-bottom:10px;font-weight:bold;">⏱️ Paused at 5m safety limit. Auto-resuming in 1 minute...</div>`
    : `<div style="background:#dcfce7;border:1px solid #86efac;color:#166534;padding:8px;border-radius:6px;margin-bottom:10px;font-weight:bold;">✅ Multi-Day Import Audit (${categoryOrder.length} folder${categoryOrder.length === 1 ? "" : "s"} tracked)</div>`;

  let html = `
    <!DOCTYPE html><html><body style="font-family:sans-serif;padding:12px;font-size:12px;color:#1e293b;background:#f8fafc;">
    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #cbd5e1;padding-bottom:6px;margin-bottom:8px;">
      <h2 style="font-size:15px;color:#0f172a;margin:0;">📋 Import Audit & Verify</h2>
      <button onclick="clearAudit()" title="Clear Audit History" style="background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;font-size:10px;cursor:pointer;">🧹 Clear Log</button>
    </div>
    ${statusBanner}

    <div style="margin-bottom:10px;">
      <label style="font-size:11px;font-weight:bold;color:#475569;display:block;margin-bottom:4px;">📂 View Day / Category:</label>
      <select id="cat-select" onchange="renderView()" style="width:100%;padding:6px;border:1px solid #94a3b8;border-radius:4px;font-size:12px;font-weight:bold;background:#ffffff;">
        <option value="ALL" ${defaultSelection === "ALL" ? "selected" : ""}>🌎 All Processed Folders (${categoryOrder.join(", ") || "None"})</option>
        ${categoryOrder.map(c => `<option value="${c}" ${defaultSelection === c ? "selected" : ""}>📁 ${c} (${byCategory[c].missing.length} missing, ${byCategory[c].matched.length} matched)</option>`).join("")}
      </select>
    </div>

    <div id="summary-badges" style="display:flex;gap:6px;margin-bottom:12px;font-size:11px;"></div>

    <!-- Quick Member Note Editor -->
    <div style="background:#ffffff;border:1px solid #cbd5e1;padding:10px;border-radius:6px;margin-bottom:14px;">
      <div style="font-weight:bold;color:#334155;margin-bottom:6px;">📝 Quick Add / Update Member Note</div>
      <input id="note-member" type="text" placeholder="Member Name (or click a Missing name below)" style="width:94%;padding:5px;margin-bottom:6px;border:1px solid #cbd5e1;border-radius:4px;font-size:11px;">
      <input id="note-text" type="text" placeholder="Note (e.g. Exempt - Vacation)" style="width:94%;padding:5px;margin-bottom:6px;border:1px solid #cbd5e1;border-radius:4px;font-size:11px;">
      <button onclick="saveNote()" style="background:#2563eb;color:white;border:none;border-radius:4px;padding:5px 10px;font-size:11px;cursor:pointer;width:100%;">Save Note to Member</button>
      <div id="note-status" style="font-size:11px;color:#16a34a;margin-top:4px;display:none;"></div>
    </div>

    <div id="audit-content"></div>

    <script>
      const BY_CAT = ${JSON.stringify(byCategory)};
      const CAT_ORDER = ${JSON.stringify(categoryOrder)};

      function fillMember(name) {
        document.getElementById('note-member').value = name;
        document.getElementById('note-text').focus();
      }

      function renderView() {
        const sel = document.getElementById('cat-select').value;
        const cats = (sel === 'ALL') ? CAT_ORDER : (BY_CAT[sel] ? [sel] : []);
        let matched = [], missing = [], newMembers = [], rankChanges = [];

        cats.forEach(c => {
          const d = BY_CAT[c];
          if (!d) return;
          matched = matched.concat(d.matched || []);
          missing = missing.concat(d.missing || []);
          newMembers = newMembers.concat(d.newMembers || []);
          rankChanges = rankChanges.concat(d.rankChanges || []);
        });

        document.getElementById('summary-badges').innerHTML =
          '<span style="background:#e2e8f0;padding:3px 8px;border-radius:4px;">🟢 Matched: <b>' + matched.length + '</b></span>' +
          '<span style="background:#fee2e2;color:#991b1b;padding:3px 8px;border-radius:4px;">🔴 Missing: <b>' + missing.length + '</b></span>' +
          '<span style="background:#fef3c7;color:#92400e;padding:3px 8px;border-radius:4px;">🟡 New: <b>' + newMembers.length + '</b></span>';

        let out = '';

        if (rankChanges.length > 0) {
          out += '<div style="background:#fffbeb;border:1px solid #fde68a;padding:10px;border-radius:6px;margin-bottom:14px;">' +
                 '<div style="font-weight:bold;color:#b45309;margin-bottom:6px;">⚠️ Rank Discrepancies (' + rankChanges.length + ')</div>' +
                 '<div style="font-size:11px;color:#78350f;margin-bottom:8px;">Manual ranks preserved. Click Apply to use screenshot rank:</div>';
          rankChanges.forEach((r, idx) => {
            out += '<div id="disc-' + idx + '" style="margin-bottom:6px;padding-bottom:6px;border-bottom:1px dashed #fde68a;display:flex;justify-content:space-between;align-items:center;">' +
                   '<div><b>' + r.name + '</b> <span style="font-size:10px;color:#64748b;">[' + r.category + ']</span><br><span style="font-size:11px;color:#92400e;">Sheet: <b>' + r.sheetRank + '</b> → OCR: <b>' + r.ocrRank + '</b></span></div>' +
                   '<button onclick="applyRank(' + r.row + ', \\'' + r.ocrRank + '\\', \\'disc-' + idx + '\\')" style="background:#d97706;color:white;border:none;border-radius:4px;padding:4px 8px;font-size:11px;cursor:pointer;">Apply ' + r.ocrRank + '</button>' +
                   '</div>';
          });
          out += '</div>';
        }

        if (missing.length > 0) {
          out += '<div style="background:#fef2f2;border:1px solid #fecaca;padding:10px;border-radius:6px;margin-bottom:14px;">' +
                 '<div style="font-weight:bold;color:#b91c1c;margin-bottom:4px;">🔴 Missing from Screenshots (' + missing.length + ')</div>' +
                 '<div style="font-size:10px;color:#7f1d1d;margin-bottom:6px;">Tip: Click any name to pre-fill the Note box above</div>' +
                 '<ul style="margin:0;padding-left:16px;color:#991b1b;max-height:180px;overflow-y:auto;">';
          missing.forEach(m => {
            const safeName = m.name.replace(/'/g, "\\\\'");
            out += '<li style="margin-bottom:3px;cursor:pointer;" onclick="fillMember(\\'' + safeName + '\\')"><b>' + m.name + '</b> <span style="font-size:10px;background:#fee2e2;padding:1px 5px;border-radius:3px;border:1px solid #fca5a5;">' + m.category + '</span></li>';
          });
          out += '</ul></div>';
        } else {
          out += '<div style="background:#f0fdf4;border:1px solid #bbf7d0;padding:8px;border-radius:6px;margin-bottom:14px;color:#166534;font-weight:bold;">🎉 Zero Missing Members for this selection!</div>';
        }

        if (newMembers.length > 0) {
          out += '<div style="background:#fefce8;border:1px solid #fef08a;padding:10px;border-radius:6px;margin-bottom:14px;">' +
                 '<div style="font-weight:bold;color:#a16207;margin-bottom:6px;">🟡 New / Unmatched Added to Bottom (' + newMembers.length + ')</div>' +
                 '<ul style="margin:0;padding-left:16px;color:#854d0e;max-height:140px;overflow-y:auto;">';
          newMembers.forEach(m => {
            out += '<li style="margin-bottom:2px;"><b>' + m.name + '</b>: ' + Number(m.score || 0).toLocaleString() + ' <span style="font-size:10px;">[' + m.category + ' - Row ' + m.row + ']</span></li>';
          });
          out += '</ul></div>';
        }

        out += '<div style="background:#ffffff;border:1px solid #e2e8f0;padding:10px;border-radius:6px;margin-bottom:14px;">' +
               '<div style="font-weight:bold;color:#15803d;margin-bottom:6px;">🟢 Matched Members (' + matched.length + ')</div>' +
               '<ul style="margin:0;padding-left:16px;color:#475569;max-height:180px;overflow-y:auto;">';
        matched.forEach(m => {
          out += '<li style="margin-bottom:2px;"><b>' + m.name + '</b>: ' + Number(m.score || 0).toLocaleString() + ' <span style="font-size:10px;color:#94a3b8;">[' + m.category + ']</span></li>';
        });
        out += '</ul></div>';

        document.getElementById('audit-content').innerHTML = out;
      }

      function applyRank(row, newRank, elemId) {
        var el = document.getElementById(elemId);
        if (el) el.style.opacity = '0.5';
        google.script.run.withSuccessHandler(function() {
          if (el) el.innerHTML = '<span style="color:#16a34a;font-weight:bold;">✅ Updated to ' + newRank + '</span>';
        }).updateRankFromDiscrepancy(row, newRank);
      }

      function saveNote() {
        var name = document.getElementById('note-member').value.trim();
        var note = document.getElementById('note-text').value.trim();
        if (!name) return;
        var st = document.getElementById('note-status');
        st.style.display = 'block';
        st.textContent = 'Saving...';
        google.script.run.withSuccessHandler(function(res) {
          st.textContent = res || '✅ Note saved!';
          document.getElementById('note-text').value = '';
        }).updateMemberNoteFromSidebar(name, note);
      }

      function clearAudit() {
        if (!confirm('Clear the Import Audit log sheet?')) return;
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).clearImportAuditLog();
      }

      renderView();
    </script>
    </body></html>`;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("📋 Import Audit & Verify"));
}

function updateRankFromDiscrepancy(row, newRank) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  if (!sheet) return;
  sheet.getRange(row, 1).setValue(newRank);
  SpreadsheetApp.flush();
}

function updateMemberNoteFromSidebar(memberName, noteText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  if (!sheet) return "❌ Sheet not found";
  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(1, 2, lastRow, 1).getValues();
  let targetNorm = normalizeString(stripSquadTag(memberName));
  for (let i = 2; i < names.length; i++) {
    if (normalizeString(stripSquadTag(names[i][0])) === targetNorm) {
      let cell = sheet.getRange(i + 1, 2);
      let existing = cell.getNote();
      let updated = existing ? (existing + "\n" + noteText) : noteText;
      cell.setNote(updated);
      return `✅ Note added to ${names[i][0]}`;
    }
  }
  return `⚠️ Member "${memberName}" not found`;
}

/**
 * Scans folders in sequence and processes the next folder containing screenshots.
 */
function processNextAvailableFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  updateQueueVisualStatus("🔍 Checking category folders...", "#fff2cc");

  const categories = [
    "Strength Ranking - Power",
    "Strength Ranking - Donations",
    "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
  ];

  let targetCategory = null;

  for (let i = 0; i < categories.length; i++) {
    let cat = categories[i];
    let folderId = CONFIG.FOLDER_MAPPINGS[cat];
    if (folderId && !folderId.startsWith("YOUR_")) {
      try {
        let folder = DriveApp.getFolderById(folderId);
        let files = folder.getFiles();
        let hasImages = false;
        while (files.hasNext()) {
          let f = files.next();
          let mime = f.getMimeType();
          if (mime.includes("image") || mime.includes("octet-stream") || f.getName().toLowerCase().endsWith(".jpg") || f.getName().toLowerCase().endsWith(".png")) {
            hasImages = true;
            break;
          }
        }
        if (hasImages) {
          targetCategory = cat;
          break;
        }
      } catch (e) {}
    }
  }

  if (!targetCategory) {
    let timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    updateQueueVisualStatus(`✅ All Folders Empty. (Checked at ${timeStr})`, "#d9ead3");
    return;
  }

  let statusText = `⚡ Processing Folder: ${targetCategory}`;
  updateQueueVisualStatus(statusText, "#fff2cc");
  debugSheet.appendRow([new Date(), `Processing next available folder: "${targetCategory}"`]);

  processFolderDirectly([targetCategory], false, false);
}

/**
 * Uses Gemini to classify an unorganized inbox screenshot into its respective category folder.
 */
function classifyScreenshotWithGemini(file) {
  if (!GEMINI_API_KEY) throw new Error("Gemini API key missing.");

  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL + ":generateContent?key=" + GEMINI_API_KEY;
  
  const prompt = `Classify this mobile game leaderboard screenshot into exactly ONE of the following categories:
- "Strength Ranking - Power" (shows overall combat power or alliance power rankings)
- "Strength Ranking - Donations" (shows alliance donation points or tech donations)
- "Mon" (Monday VS Duel)
- "Tue" (Tuesday VS Duel)
- "Wed" (Wednesday VS Duel)
- "Thu" (Thursday VS Duel)
- "Fri" (Friday VS Duel)
- "Sat" (Saturday VS Duel)

Return STRICT JSON format: {"category": "<exact category string>"}. No markdown, no extra text.`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: file.getMimeType(), data: Utilities.base64Encode(file.getBlob().getBytes()) } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: "application/json"
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);

  let raw = json.candidates[0].content.parts[0].text.trim();
  let match = raw.match(/\{.*\}/s);
  let parsed = JSON.parse(match ? match[0] : raw);
  return parsed.category;
}

/**
 * Universal Start Button & Auto-Watcher Entrypoint:
 * Sorts Inbox and ALWAYS chains into processing the next available folder.
 */
function sortAndProcessInboxFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  if (!CONFIG.INBOX_FOLDER_ID || CONFIG.INBOX_FOLDER_ID.startsWith("YOUR_")) {
    updateQueueVisualStatus("⚠️ Error: INBOX_FOLDER_ID missing.", "#f4cccc");
    throw new Error("INBOX_FOLDER_ID is not configured in Config.gs.");
  }

  updateQueueVisualStatus("📥 Scanning Inbox folder...", "#fff2cc");
  let inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  let files = inboxFolder.getFiles();
  let processedCount = 0;

  let fileList = [];
  while (files.hasNext()) {
    let f = files.next();
    let mime = f.getMimeType();
    if (mime.includes("image") || mime.includes("octet-stream") || f.getName().toLowerCase().endsWith(".jpg") || f.getName().toLowerCase().endsWith(".png")) {
      fileList.push(f);
    }
  }

  let totalFiles = fileList.length;
  if (totalFiles === 0) {
    updateQueueVisualStatus("🔍 Inbox empty. Checking category folders...", "#fff2cc");
    processNextAvailableFolder();
    return { success: true, count: 0 };
  }

  for (let i = 0; i < totalFiles; i++) {
    let f = fileList[i];
    try {
      updateQueueVisualStatus(`📥 Sorting Inbox (${i + 1}/${totalFiles}): ${f.getName()}`, "#fff2cc");
      let category = classifyScreenshotWithGemini(f);
      let targetFolderId = CONFIG.FOLDER_MAPPINGS[category];
      if (targetFolderId && !targetFolderId.startsWith("YOUR_")) {
        let targetFolder = DriveApp.getFolderById(targetFolderId);
        f.moveTo(targetFolder);
        processedCount++;
        debugSheet.appendRow([new Date(), `Routed inbox file "${f.getName()}" to category: ${category}`]);
      }
    } catch (err) {
      debugSheet.appendRow([new Date(), `⚠️ Failed to classify "${f.getName()}": ${err.message}`]);
    }
  }

  updateQueueVisualStatus(`✅ Sorted ${processedCount}/${totalFiles} files! Starting folder import...`, "#d9ead3");
  ss.toast(`Sorted ${processedCount} screenshot(s). Proceeding to import...`, "📥 Inbox Sorted", 4);

  processNextAvailableFolder();
  return { success: true, count: processedCount };
}

function parseLeaderboardBatchWithGemini(batchFiles, currentLabel) {
  if (!GEMINI_API_KEY) throw new Error("Gemini API key is not configured in Script Properties.");

  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL + ":generateContent?key=" + GEMINI_API_KEY;

  let parts = [
    {
      text: `You are an expert OCR parser for a mobile game leaderboard screenshot for category: "${currentLabel}".
Extract all player rows visible in these screenshots.
Return a STRICT JSON array of objects with the following keys:
- "name": The player username (string)
- "score": The numerical score, power, or donation amount (number or string with K/M)
- "rank": The alliance rank if visible (e.g., "R1", "R2", "R3", "R4", "R5"), otherwise empty string "".

Rules:
1. Do not omit partially cut off rows at the top or bottom edges.
2. Return ONLY the JSON array, no commentary, markdown backticks, or extra text.`
    }
  ];

  batchFiles.forEach(file => {
    parts.push({
      inline_data: {
        mime_type: file.getMimeType(),
        data: Utilities.base64Encode(file.getBlob().getBytes())
      }
    });
  });

  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(json.error.message || "Gemini API Error");
  }

  let rawText = json.candidates[0].content.parts[0].text.trim();
  let match = rawText.match(/\[.*\]/s);
  let jsonString = match ? match[0] : rawText.replace(/^```json\s*|^```\s*|\s*```$/g, "");
  return JSON.parse(jsonString);
}

function parseLeaderboardWithGemini(file, currentLabel) {
  return parseLeaderboardBatchWithGemini([file], currentLabel);
}

/**
 * Reverts Data Import to pre-import snapshot and restores screenshots back to input folder.
 */
function rollbackLastImport() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let tempSnap = ss.getSheetByName("_ImportBackup");

  if (!tempSnap || tempSnap.getLastRow() < 2) {
    ui.alert("⚠️ Rollback Unavailable", "No recent backup snapshot was found. Your sheet will not be cleared.", ui.ButtonSet.OK);
    return;
  }

  if (ui.alert("↩ Confirm Rollback", "This will revert your Data Import sheet to the exact state it was in before your last import, and move any processed screenshots back to their input folders. Proceed?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  sheet.clear();
  tempSnap.getDataRange().copyTo(sheet.getRange(1, 1));

  let master;
  try { master = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID); } catch(e) {}
  if (master) {
    let keys = Object.keys(CONFIG.PROCESSED_SUBFOLDER_NAMES);
    keys.forEach(k => {
      let sub = master.getFoldersByName(CONFIG.PROCESSED_SUBFOLDER_NAMES[k]);
      if (sub.hasNext()) {
        let files = sub.next().getFiles();
        let target = DriveApp.getFolderById(CONFIG.FOLDER_MAPPINGS[k]);
        while (files.hasNext()) {
          try { files.next().moveTo(target); } catch(e) {}
        }
      }
    });
  }

  updateQueueVisualStatus("↩️ Rollback Complete - Restored pre-import state", "#d9ead3");
  ui.alert("✅ Rollback Complete!", "Your roster, notes, and screenshots have been successfully restored to their pre-import state.", ui.ButtonSet.OK);
}

function emptyProcessedFolders(silent = false) {
  let ui = SpreadsheetApp.getUi();
  if (!silent && ui.alert("Empty Processed Folders", "Move all processed screenshots to trash?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let masterFolder;
  try { masterFolder = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID); } catch (e) { return; }

  let deleted = 0;
  Object.values(CONFIG.PROCESSED_SUBFOLDER_NAMES).forEach(subName => {
    let subs = masterFolder.getFoldersByName(subName);
    if (subs.hasNext()) {
      let files = subs.next().getFiles();
      while (files.hasNext()) { try { files.next().setTrashed(true); deleted++; } catch (e) {} }
    }
  });
  if (!silent) ui.alert("Cleanup Complete", `Moved ${deleted} files to trash.`, ui.ButtonSet.OK);
}

function restoreProcessedScreenshots() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><body style="font-family:sans-serif;padding:20px;">
    <h3>🔄 Restore Screenshots</h3>
    <select id="sel" style="width:100%;padding:8px;margin-bottom:15px;">
    <option value="ALL">All Processed Folders</option>
    <option value="Strength Ranking - Power">Strength Ranking - Power</option>
    <option value="Strength Ranking - Donations">Strength Ranking - Donations</option>
    <option value="Mon">Monday VS</option>
    <option value="Tue">Tuesday VS</option>
    <option value="Wed">Wednesday VS</option>
    <option value="Thu">Thursday VS</option>
    <option value="Fri">Friday VS</option>
    <option value="Sat">Saturday VS</option>
    </select>
    <button onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).handleDropdownRestoration(document.getElementById('sel').value)" style="background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;">Restore</button>
    </body></html>
  `).setWidth(400).setHeight(220), "Restoration Center");
}

function handleDropdownRestoration(choice) {
  let master;
  try { master = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID); } catch(e) { return; }
  let keys = (choice && choice !== "ALL") ? [choice] : Object.keys(CONFIG.PROCESSED_SUBFOLDER_NAMES);
  let restored = 0;
  keys.forEach(k => {
    let subName = CONFIG.PROCESSED_SUBFOLDER_NAMES[k];
    if (!subName) return;
    let sub = master.getFoldersByName(subName);
    if (sub.hasNext()) {
      let files = sub.next().getFiles();
      let target = DriveApp.getFolderById(CONFIG.FOLDER_MAPPINGS[k]);
      while (files.hasNext()) { try { files.next().moveTo(target); restored++; } catch(e) {} }
    }
  });
  SpreadsheetApp.getUi().alert("Restoration Complete", `Restored ${restored} file(s).`, SpreadsheetApp.getUi().ButtonSet.OK);
}