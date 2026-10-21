/**
 * ScreenshotEngine.gs - Automated Gemini OCR Batch Ingestion Pipeline
 * TWST Roster Management System
 */

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
const GEMINI_MODEL = "gemini-3.8-flash"; // Matches your active API model configuration
const BATCH_SIZE = 5;

/**
 * Main ingestion handler for folder processing.
 */
function processFolderDirectly(headingLabelsArray, returnStats = false, ignoreRanks = false) {
  const START_TIME = Date.now();
  const MAX_RUNTIME_MS = 300000; // 5-minute safety threshold
  let timeoutTriggered = false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  // --- PERSISTENT IMPORT BACKUP CREATION (NEVER DELETED ON SUCCESS) ---
  try {
    let oldSnap = ss.getSheetByName("_ImportBackup");
    if (oldSnap) {
      ss.deleteSheet(oldSnap); // Replace prior rollback point with fresh pre-import snapshot
    }
    let tempSnap = ss.insertSheet("_ImportBackup");
    let srcRange = sheet.getDataRange();
    if (srcRange.getNumRows() > 0 && srcRange.getNumColumns() > 0) {
      srcRange.copyTo(tempSnap.getRange(1, 1));
      srcRange.copyTo(tempSnap.getRange(1, 1), { contentsOnly: true });
    }
    tempSnap.hideSheet();
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
    if (timeoutTriggered) break;

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

    let rosterRange = findRosterRange(sheet);
    let sheetData = sheet.getRange(rosterRange.startRow, 1, rosterRange.rowCount, rosterRange.colCount).getValues();
    let headers = sheetData[0].map(val => val.toString().trim().toLowerCase());

    let nameColIdx = headers.findIndex(x => x.includes("member"));
    if (nameColIdx === -1) nameColIdx = 1;
    let rankColIdx = headers.indexOf("rank");

    let isStrengthRanking = (currentLabel === "Strength Ranking - Power" || currentLabel === "Strength Ranking - Donations");
    let metricColIdx = -1;
    let headerAliases = CONFIG.HEADER_ALIASES[currentLabel] || [];
    metricColIdx = headers.findIndex(headerStr => headerAliases.some(alias => headerStr.includes(alias)));

    // --- BATCH FILE EXECUTION ---
    for (let bIdx = 0; bIdx < allFiles.length; bIdx += BATCH_SIZE) {
      if (Date.now() - START_TIME > MAX_RUNTIME_MS) {
        timeoutTriggered = true;
        debugSheet.appendRow([new Date(), `⚠️ Safety limit reached at file ${bIdx + 1}. Auto-scheduling self-resumption...`]);
        try {
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

      let rosterNames = sheetData.slice(1)
        .map((row, idx) => ({
          rawName: row[nameColIdx] ? row[nameColIdx].toString().trim() : "",
          cleanName: normalizeString(row[nameColIdx]),
          baseName: stripSquadTag(row[nameColIdx]),
          index: idx + 1
        }))
        .filter(obj => obj.cleanName.length >= 1);

      let parsedPlayers = [];
      try {
        parsedPlayers = parseLeaderboardBatchWithGemini(batchFiles, currentLabel);
        debugSheet.appendRow([new Date(), `Gemini extracted ${parsedPlayers.length} entries from batch (${startNum}-${endNum})`]);
      } catch (geminiErr) {
        debugSheet.appendRow([new Date(), `❌ Gemini Parse Error on ${batchLabelStr}: ${geminiErr.message}`]);
        continue;
      }

      let batchMatches = 0;
      parsedPlayers.forEach(p => {
        let cleanCandidateName = p.name ? p.name.trim() : "";
        if (!cleanCandidateName || (cleanCandidateName.toUpperCase().includes("TWISTED") && cleanCandidateName.length < 5)) return;

        let scoreVal = cleanNumericValues(p.score);
        let rawRankVal = p.rank ? p.rank.toString().trim().toUpperCase() : "";
        let rankVal = /^R[1-5]$/i.test(rawRankVal) ? rawRankVal : "";

        let matchedAliasKey = Object.keys(CONFIG.MEMBER_ALIASES).find(
          key => key.toLowerCase() === cleanCandidateName.toLowerCase() || normalizeString(key) === normalizeString(cleanCandidateName)
        );

        if (matchedAliasKey) {
          cleanCandidateName = CONFIG.MEMBER_ALIASES[matchedAliasKey];
        }

        let candidateBase = stripSquadTag(cleanCandidateName);
        let candidateClean = normalizeString(cleanCandidateName);
        let matchedRowIdx = -1;

        for (let r = 0; r < rosterNames.length; r++) {
          let ro = rosterNames[r];
          if (candidateClean === ro.cleanName || (candidateBase.length >= 1 && candidateBase === ro.baseName)) {
            matchedRowIdx = ro.index;
            break;
          }
        }

        if (matchedRowIdx === -1 && candidateClean.length >= 3) {
          for (let r = 0; r < rosterNames.length; r++) {
            let ro = rosterNames[r];
            if (ro.cleanName.startsWith(candidateClean)) {
              matchedRowIdx = ro.index;
              break;
            }
          }
        }

        if (matchedRowIdx === -1 && candidateBase.length >= 3) {
          for (let r = 0; r < rosterNames.length; r++) {
            let ro = rosterNames[r];
            let candCore = candidateBase.replace(/[a-d][a-b]?$/i, "");
            let roCore = ro.baseName.replace(/[a-d][a-b]?$/i, "");
            if (candCore.length >= 3 && candCore === roCore) {
              matchedRowIdx = ro.index;
              break;
            }
          }
        }

        if (matchedRowIdx !== -1) {
          let actualRow = matchedRowIdx + 1;
          sheetData[matchedRowIdx][nameColIdx] = cleanCandidateName;
          if (metricColIdx > -1) sheetData[matchedRowIdx][metricColIdx] = scoreVal;

          cumulativeMatchedRows.add(actualRow);
          batchMatches++;
          auditMatched.push({ name: cleanCandidateName, score: scoreVal, raw: p.name });

          if (rankColIdx > -1 && rankVal && currentLabel === "Strength Ranking - Power") {
            let existingRank = sheetData[matchedRowIdx][rankColIdx] ? sheetData[matchedRowIdx][rankColIdx].toString().trim().toUpperCase() : "";
            if (existingRank && existingRank !== rankVal && existingRank !== "R5" && existingRank !== "R4") {
              rankDiscrepancies.push({
                row: actualRow,
                name: cleanCandidateName,
                sheetRank: existingRank,
                ocrRank: rankVal
              });
            } else if (!existingRank && !ignoreRanks) {
              sheetData[matchedRowIdx][rankColIdx] = rankVal;
            }
          }
        } else {
          let alreadyAdded = sheetData.some(row => stripSquadTag(row[nameColIdx]) === candidateBase);
          if (!alreadyAdded) {
            let newRow = new Array(rosterRange.colCount).fill("");
            newRow[nameColIdx] = cleanCandidateName;
            if (metricColIdx > -1) newRow[metricColIdx] = scoreVal;
            if (rankColIdx > -1 && rankVal && !ignoreRanks) newRow[rankColIdx] = rankVal;

            sheetData.push(newRow);
            let freshRowIdx = sheetData.length;
            cumulativeMatchedRows.add(freshRowIdx);
            auditNewMembers.push({ name: cleanCandidateName, score: scoreVal });
          }
        }
      });

      // Pure raw data commit (no formatting overwrites)
      if (sheetData.length > 1) {
        let nameValues = sheetData.slice(1).map(row => [row[nameColIdx]]);
        sheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, nameValues.length, 1).setValues(nameValues);
      }

      if (metricColIdx > -1 && sheetData.length > 1) {
        let metricValues = sheetData.slice(1).map(row => [row[metricColIdx]]);
        sheet.getRange(rosterRange.startRow + 1, metricColIdx + 1, metricValues.length, 1).setValues(metricValues);
      }

      if (sheetData.length > rosterRange.rowCount) {
        let addedRowsBatch = sheetData.slice(rosterRange.rowCount);
        let insertRow = rosterRange.startRow + rosterRange.rowCount;
        sheet.getRange(insertRow, 1, addedRowsBatch.length, addedRowsBatch[0].length).setValues(addedRowsBatch);
        rosterRange.rowCount = sheetData.length;
      }

      SpreadsheetApp.flush();
      debugSheet.appendRow([new Date(), `Committed Batch (${startNum}-${endNum}) -> Records: ${batchMatches}`]);
      totalProcessedFiles += batchFiles.length;

      if (targetProcessedSubfolder) {
        batchFiles.forEach(f => {
          try { f.moveTo(targetProcessedSubfolder); } catch (e) {}
        });
      }
    }

    // Missing members identification
    let finalRosterRange = findRosterRange(sheet);
    let finalSheetData = sheet.getRange(finalRosterRange.startRow, 1, finalRosterRange.rowCount, finalRosterRange.colCount).getValues();
    let nIdx = finalSheetData[0].map(h => h.toString().trim().toLowerCase()).indexOf("member");
    if (nIdx === -1) nIdx = 1;

    let finalRosterNotes = sheet.getRange(finalRosterRange.startRow, nIdx + 1, finalRosterRange.rowCount, 1).getNotes();
    let metricValues = (metricColIdx > -1) ? sheet.getRange(finalRosterRange.startRow + 1, metricColIdx + 1, finalRosterRange.rowCount - 1, 1).getValues() : [];

    for (let r = 1; r < finalSheetData.length; r++) {
      let memberName = finalSheetData[r][nIdx] ? finalSheetData[r][nIdx].toString().trim() : "";
      if (!memberName || memberName.startsWith("⚠️") || memberName.startsWith("---")) continue;

      let actualRow = finalRosterRange.startRow + r;
      let note = finalRosterNotes[r] ? finalRosterNotes[r][0].toLowerCase() : "";
      let isExempt = typeof isExemptNote === "function" ? isExemptNote(note) : false;
      let scoreCell = (metricValues[r - 1]) ? metricValues[r - 1][0] : "";

      let isBlankScore = (scoreCell === "" || scoreCell === null || scoreCell === undefined);
      let isMatchedInRun = cumulativeMatchedRows.has(actualRow);

      if (!isMatchedInRun && isBlankScore && !isExempt) {
        auditMissing.push(memberName);
      }
    }

    // Write to dedicated 'Import Audit' Tab
    writeToImportAuditTab(ss, currentLabel, auditMatched, auditMissing, auditNewMembers, rankDiscrepancies);
  }

  // Display docked sidebar
  if (!returnStats) {
    showImportAuditSidebar(headingLabelsArray[0] || "Import", auditMatched, auditMissing, auditNewMembers, rankDiscrepancies, timeoutTriggered, totalProcessedFiles);
  }

  return {
    timedOut: timeoutTriggered,
    processedCount: totalProcessedFiles,
    matchedCount: auditMatched.length
  };
}

/**
 * Creates or updates the dedicated 'Import Audit' sheet tab.
 */
function writeToImportAuditTab(ss, label, matched, missing, added, discrepancies) {
  let auditSheet = ss.getSheetByName("Import Audit") || ss.insertSheet("Import Audit");
  auditSheet.clear();

  let rows = [
    ["Category", "Member Name", "Imported Score", "Status / Notes"],
    ["SUMMARY", `Category: ${label}`, `Total Matched: ${matched.length}`, `Date: ${new Date().toLocaleString()}`],
    ["---", "---", "---", "---"]
  ];

  if (discrepancies.length > 0) {
    rows.push(["⚠️ RANK MISMATCH", "Requires Review", "", ""]);
    discrepancies.forEach(d => {
      rows.push(["Rank Discrepancy", d.name, `OCR: ${d.ocrRank}`, `Sheet currently has: ${d.sheetRank}`]);
    });
    rows.push(["---", "---", "---", "---"]);
  }

  if (added.length > 0) {
    rows.push(["🟡 NEW MEMBERS", "Added to bottom", "", ""]);
    added.forEach(a => {
      rows.push(["New Member", a.name, a.score, "Inserted into roster"]);
    });
    rows.push(["---", "---", "---", "---"]);
  }

  if (missing.length > 0) {
    rows.push(["🔴 MISSING", "Not found in screenshots", "", ""]);
    missing.forEach(m => {
      rows.push(["Missing", m, "-", "No screenshot entry"]);
    });
    rows.push(["---", "---", "---", "---"]);
  }

  rows.push(["🟢 MATCHED", `Count: ${matched.length}`, "", ""]);
  matched.forEach(m => {
    rows.push(["Matched", m.name, m.score, `OCR Text: ${m.raw || ""}`]);
  });

  auditSheet.getRange(1, 1, rows.length, 4).setValues(rows);
  auditSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#38bdf8");
}

/**
 * Renders the non-blocking sidebar.
 */
function showImportAuditSidebar(label, matched, missing, added, discrepancies, timedOut, fileCount) {
  let discHtml = "";
  if (discrepancies.length > 0) {
    discHtml = `
      <div style="background:#451a03;border:1px solid #b45309;border-radius:6px;padding:8px;margin-bottom:10px;">
        <div style="color:#fbbf24;font-weight:bold;font-size:11px;text-transform:uppercase;">⚠️ Rank Discrepancies (${discrepancies.length})</div>
        <div style="font-size:11px;color:#fde68a;margin-top:4px;">
          ${discrepancies.map(d => `<div><strong>${d.name}</strong>: Sheet=${d.sheetRank} vs OCR=${d.ocrRank}</div>`).join("")}
        </div>
      </div>
    `;
  }

  let html = `
    <!DOCTYPE html><html><head><base target="_top">
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #f8fafc; padding: 12px; font-size: 12px; margin: 0; }
      h3 { margin-top:0; color: #38bdf8; font-size: 14px; display:flex; justify-content:space-between; align-items:center; }
      .badge { font-size: 10px; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
      .badge-ok { background: #166534; color: #bbf7d0; }
      .badge-paused { background: #854d0e; color: #fef08a; }
      .card { background: #1e293b; border: 1px solid #334155; border-radius: 6px; padding: 8px; margin-bottom: 10px; }
      .title { font-weight: bold; font-size: 11px; text-transform: uppercase; margin-bottom: 4px; }
      ul { margin: 0; padding-left: 16px; max-height: 120px; overflow-y: auto; }
      li { margin-bottom: 2px; }
      .note { font-size: 11px; color: #94a3b8; margin-bottom: 10px; line-height: 1.4; }
    </style></head><body>
      <h3><span>📋 ${label} Audit</span><span class="badge ${timedOut ? 'badge-paused' : 'badge-ok'}">${timedOut ? 'Paused' : 'Done'}</span></h3>
      <p class="note">Processed ${fileCount} file(s). You can freely click and edit the <strong>Data Import</strong> tab right now! Breakdown saved to <strong>Import Audit</strong> tab.</p>
      
      ${discHtml}

      <div class="card">
        <div class="title" style="color:#f87171;">🔴 Missing From Screenshots (${missing.length})</div>
        <ul>${missing.length > 0 ? missing.map(m => `<li>${m}</li>`).join("") : '<li style="color:#4ade80;">None missing</li>'}</ul>
      </div>

      <div class="card">
        <div class="title" style="color:#fbbf24;">🟡 New Members (${added.length})</div>
        <ul>${added.length > 0 ? added.map(a => `<li>${a.name} (${a.score})</li>`).join("") : '<li>No new members</li>'}</ul>
      </div>

      <div class="card">
        <div class="title" style="color:#4ade80;">🟢 Matched Members (${matched.length})</div>
        <ul>${matched.slice(0, 15).map(m => `<li>${m.name}:${m.score}</li>`).join("")}</ul>
        ${matched.length > 15 ? `<div style="font-size:10px;color:#94a3b8;margin-top:4px;">...and ${matched.length - 15} more in 'Import Audit' tab</div>` : ''}
      </div>
    </body></html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle(`Audit: ${label}`));
}

/**
 * Gemini multi-image batch parsing using active model.
 */
function parseLeaderboardBatchWithGemini(batchFiles, currentLabel) {
  if (!GEMINI_API_KEY) throw new Error("Gemini API key is not configured in Script Properties.");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  
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

function stripSquadTag(str) {
  if (!str) return "";
  return str.toString().replace(/\[.*?\]\s*/g, "").trim();
}

/**
 * Scans folders in sequence and processes the next folder containing screenshots.
 */
function processNextAvailableFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

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

  const url = `[https://generativelanguage.googleapis.com/v1beta/models/$](https://generativelanguage.googleapis.com/v1beta/models/$){GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `Classify this mobile game leaderboard screenshot into exactly ONE of the following categories:
- "Strength Ranking - Power"
- "Strength Ranking - Donations"
- "Mon"
- "Tue"
- "Wed"
- "Thu"
- "Fri"
- "Sat"

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
    </body></html>`).setWidth(400).setHeight(220), "Restoration Center");
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

/**
 * Scans the main Inbox folder, classifies unorganized screenshots using Gemini OCR, 
 * and routes them into their respective category folders.
 */
function sortAndProcessInboxFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  if (!CONFIG.INBOX_FOLDER_ID || CONFIG.INBOX_FOLDER_ID.startsWith("YOUR_")) {
    throw new Error("INBOX_FOLDER_ID is not configured in Config.gs.");
  }

  let inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  let files = inboxFolder.getFiles();
  let processedCount = 0;

  while (files.hasNext()) {
    let f = files.next();
    let mime = f.getMimeType();
    if (mime.includes("image") || mime.includes("octet-stream") || f.getName().toLowerCase().endsWith(".jpg") || f.getName().toLowerCase().endsWith(".png")) {
      try {
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
  }

  ss.toast(`Successfully sorted ${processedCount} screenshot(s) from inbox.`, "📥 Inbox Sorted", 5);
  return { success: true, count: processedCount };
}