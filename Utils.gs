/**
 * Utils.gs - Shared Utility Functions, Strike Backfill & Mobile Event Handlers
 * TWST Roster Management System
 */

function updateQueueVisualStatus(message, backgroundColor) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let statusCell = sheet.getRange("F1");
    statusCell.setValue(message);
    statusCell.setBackground(backgroundColor);
    statusCell.setFontWeight("bold");
    SpreadsheetApp.flush();
  } catch (e) {}
}

function isExemptNote(noteText) {
  if (!noteText) return false;
  let lower = noteText.toString().toLowerCase();
  let keywords = (typeof CONFIG !== "undefined" && CONFIG.TARGETS && CONFIG.TARGETS.EXEMPT_KEYWORDS)
    ? CONFIG.TARGETS.EXEMPT_KEYWORDS
    : ["exempt", "vacation", "loa", "personal", "absence", "sick"];
  return keywords.some(kw => lower.includes(kw.toLowerCase()));
}

function deleteExtraColumns(sheetParam) {
  try {
    let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
    if (!sheet) return;
    let maxCols = sheet.getMaxColumns();
    if (maxCols > 13) {
      sheet.deleteColumns(14, maxCols - 13);
    }
  } catch (e) {}
}

function findRosterRange(sheet) {
  let lastRow = sheet.getLastRow();
  if (lastRow < 1) return { startRow: 2, rowCount: 2, colCount: 13 };
  let lastCol = Math.min(sheet.getLastColumn() || 13, 13);
  let values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    let rowStr = values[i].map(v => v.toString().trim().toLowerCase());
    if (rowStr.includes("member") && rowStr.includes("rank")) {
      headerRow = i + 1;
      break;
    }
  }
  if (headerRow === -1) return { startRow: 2, rowCount: Math.max(lastRow - 1, 2), colCount: lastCol };

  let trueLastMemberRow = headerRow;
  for (let i = headerRow; i < values.length; i++) {
    let memberVal = values[i][1] ? values[i][1].toString().trim() : "";
    let rowStr = values[i].map(v => v.toString().trim().toLowerCase());
    if (rowStr.includes("processed summary") || rowStr.includes("processing summary") || rowStr.includes("matched")) break;
    if (memberVal !== "") trueLastMemberRow = i + 1;
  }
  return { startRow: headerRow, rowCount: Math.max(trueLastMemberRow - headerRow + 1, 2), colCount: lastCol };
}

function cleanAndSortRoster(sheetParam) {
  let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  let values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 2)).getValues();
  let rowsToDelete = [];

  for (let i = 0; i < values.length; i++) {
    let colA = values[i][0] ? values[i][0].toString() : "";
    let colB = values[i][1] ? values[i][1].toString() : "";
    if (colA.includes("--- MISSING") || colB.includes("--- MISSING") || colA.includes("--- NEW") || colB.includes("--- NEW")) {
      rowsToDelete.push(i + 1);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

function getTotalRosterCount(sheet) {
  let range = findRosterRange(sheet);
  if (!range || range.rowCount <= 1) return 0;
  let values = sheet.getRange(range.startRow + 1, 1, range.rowCount - 1, range.colCount).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    let name = values[i][1] ? values[i][1].toString().trim() : "";
    if (name !== "" && !name.startsWith("⚠️") && !name.startsWith("MISSING")) count++;
  }
  return count;
}

function normalizeString(str) {
  if (!str) return "";
  let s = str.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  return s.replace(/[^a-z0-9]/g, "").trim();
}

function stripSquadTag(str) {
  if (!str) return "";
  return str.toString().replace(/\[.*?\]\s*/g, "").trim();
}

function cleanNumericValues(val) {
  if (val === null || val === undefined || val === "") return 0;
  let normalized = val.toString().trim().toUpperCase();
  if (normalized === "" || normalized === "-") return 0;
  if (normalized.includes("K")) return Math.round(parseFloat(normalized.replace("K", "").replace(/,/g, ".")) * 1000);
  if (normalized.includes("M")) return Math.round(parseFloat(normalized.replace("M", "").replace(/,/g, ".")) * 1000000);
  let parsed = parseFloat(normalized.replace(/,/g, ""));
  return !isNaN(parsed) ? Math.round(parsed) : 0;
}

function clearDebugLogSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug Log");
  if (sheet) sheet.clear().getRange("A1").setValue("Debug Log Initialized.").setFontStyle("italic");
}

function showDebugSnippet() {
  let debug = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug Log");
  let text = "No debug logs.";
  if (debug && debug.getLastRow() > 0) {
    let lastRow = debug.getLastRow();
    let startRow = Math.max(1, lastRow - 30);
    let values = debug.getRange(startRow, 1, (lastRow - startRow) + 1, 2).getValues();
    text = values.map(row => `${row[0]} | ${row[1]}`).join("\n");
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<textarea readonly style="width:100vw;height:100vh;background:#1e1e1e;color:#9cdcfe;padding:10px;border:none;">${text}</textarea>`).setWidth(800).setHeight(500), 
    "Debug Log"
  );
}

/**
 * SAFE CLEANUP:
 * Clears daily score cells and OCR notes, but PRESERVES:
 * - Amber highlights (#f5a623) on Member names
 * - Strike / Penalty notes (⚠️ Pending Removal, ⚠️ Demoted Last Week)
 * - Exemption notes (LOA, Vacation, etc.)
 */
function clearAllHighlightsAndNotes() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let rRange = findRosterRange(sheet);
  if (!rRange || rRange.rowCount <= 1) return;

  let headers = sheet.getRange(rRange.startRow, 1, 1, rRange.colCount).getValues()[0].map(h => h.toString().toLowerCase());
  let nameColIdx = headers.findIndex(h => h.includes("member"));
  if (nameColIdx === -1) nameColIdx = 1;

  let tableRange = sheet.getRange(rRange.startRow + 1, 1, rRange.rowCount - 1, rRange.colCount);
  let backgrounds = tableRange.getBackgrounds();
  let notes = tableRange.getNotes();

  const AMBER_COLOR = (CONFIG.COLORS.ELIMINATION_ALERT || "#f5a623").toLowerCase();

  for (let r = 0; r < tableRange.getNumRows(); r++) {
    for (let c = 0; c < tableRange.getNumColumns(); c++) {
      let cellColor = (backgrounds[r][c] || "").toLowerCase();
      let noteStr = notes[r][c] || "";

      // Preserve Amber highlight on member column
      if (c === nameColIdx && cellColor === AMBER_COLOR) {
        // Keep Amber!
      } else {
        backgrounds[r][c] = null;
      }

      // Preserve Strike notes and Exemption notes
      if (noteStr) {
        let isStrike = noteStr.includes("⚠️ Pending Removal") || noteStr.includes("⚠️ Demoted Last Week");
        let isExempt = isExemptNote(noteStr);

        if (!isStrike && !isExempt) {
          if (noteStr.includes("Matched & Verified") ||
              noteStr.includes("Missing from all") ||
              noteStr.includes("MEMBER NOT DETECTED") ||
              noteStr.includes("New Member Added") ||
              noteStr.includes("Score:")) {
            notes[r][c] = "";
          }
        }
      }
    }
  }

  tableRange.setBackgrounds(backgrounds);
  tableRange.setNotes(notes);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("Cleared temporary score fills. Amber highlights & strike notes preserved.", "🧹 Cleaned", 4);
}

/**
 * Temporary / Maintenance Tool:
 * Scans past archive spreadsheets from the "Archive Log" tab and backfills:
 * - Demotion counters (x1, x2, ...)
 * - Removal week counters (Week 1, Week 2, ...)
 * - Amber highlights on flagged members
 */
function backfillStrikeNotesFromArchives() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  const archiveSheet = ss.getSheetByName("Archive Log");
  const ui = SpreadsheetApp.getUi();

  if (!archiveSheet || archiveSheet.getLastRow() < 2) {
    ui.alert("Archive Log Empty", "Could not find past archives listed in the 'Archive Log' tab.", ui.ButtonSet.OK);
    return;
  }

  ui.toast("Scanning past archive workbooks...", "🔍 Backfill Running", 15);

  let archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 3).getValues();
  let memberHistory = {};

  // Read each archive file in order
  for (let a = 0; a < archiveData.length; a++) {
    let linkFormula = archiveData[a][2] ? archiveData[a][2].toString() : "";
    let idMatch = linkFormula.match(/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) continue;

    try {
      let archiveFile = SpreadsheetApp.openById(idMatch[1]);
      let snapSheet = archiveFile.getSheetByName("Last Week") || archiveFile.getSheetByName("Data Import") || archiveFile.getSheets()[0];
      if (!snapSheet) continue;

      let rRange = findRosterRange(snapSheet);
      let vals = snapSheet.getRange(rRange.startRow, 1, rRange.rowCount, rRange.colCount).getValues();
      let hdrs = vals[0].map(h => h.toString().toLowerCase());
      let nIdx = hdrs.findIndex(h => h.includes("member"));
      let reasonIdx = hdrs.findIndex(h => h.includes("reason") || h.includes("notes"));

      if (nIdx === -1 || reasonIdx === -1) continue;

      for (let r = 1; r < vals.length; r++) {
        let name = vals[r][nIdx] ? vals[r][nIdx].toString().trim() : "";
        if (!name) continue;
        let reason = vals[r][reasonIdx] ? vals[r][reasonIdx].toString() : "";

        if (!memberHistory[name]) {
          memberHistory[name] = { demotionCount: 0, consecutiveRemoval: 0 };
        }

        if (reason.includes("[Removal]")) {
          memberHistory[name].consecutiveRemoval += 1;
        } else {
          memberHistory[name].consecutiveRemoval = 0;
        }

        if (reason.includes("[Demotion]")) {
          memberHistory[name].demotionCount += 1;
        }
      }
    } catch (e) {
      console.error("Archive parse error: " + e.message);
    }
  }

  // Apply to current Data Import sheet
  let curRange = findRosterRange(sheet);
  let curVals = sheet.getRange(curRange.startRow, 1, curRange.rowCount, curRange.colCount).getValues();
  let curHdrs = curVals[0].map(h => h.toString().toLowerCase());
  let nameCol = curHdrs.findIndex(h => h.includes("member")) + 1;

  let nameRange = sheet.getRange(curRange.startRow + 1, nameCol, curRange.rowCount - 1, 1);
  let notes = nameRange.getNotes();
  let backgrounds = nameRange.getBackgrounds();
  let updatedCount = 0;

  for (let r = 1; r < curVals.length; r++) {
    let name = curVals[r][nameCol - 1] ? curVals[r][nameCol - 1].toString().trim() : "";
    let hist = memberHistory[name];
    if (!hist) continue;

    let existingNote = notes[r - 1][0] || "";
    let cleanNote = existingNote
      .split("\n")
      .filter(line => !line.includes("⚠️ Demoted Last Week") && !line.includes("⚠️ Pending Removal"))
      .join("\n")
      .trim();

    let strikeText = "";
    let flagAmber = false;

    if (hist.consecutiveRemoval > 0) {
      strikeText = `⚠️ Pending Removal (Week ${hist.consecutiveRemoval})`;
      flagAmber = true;
    } else if (hist.demotionCount > 0) {
      strikeText = `⚠️ Demoted Last Week (x${hist.demotionCount})`;
      flagAmber = true;
    }

    if (strikeText) {
      notes[r - 1][0] = cleanNote ? `${cleanNote}\n${strikeText}` : strikeText;
      if (flagAmber) backgrounds[r - 1][0] = CONFIG.COLORS.ELIMINATION_ALERT || "#f5a623";
      updatedCount++;
    }
  }

  nameRange.setNotes(notes);
  nameRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();

  ui.alert("Backfill Complete", `Successfully synced strike history across past archives. Updated ${updatedCount} members with multi-week notes and amber highlights.`, ui.ButtonSet.OK);
}

function menuUpdateRankWithNewRank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  const scanRows = Math.min(lastRow, 10);
  const headerData = sheet.getRange(1, 1, scanRows, lastCol).getValues();

  let headerRow = -1;
  let rankCol = -1;
  let newRankCol = -1;

  for (let r = 0; r < headerData.length; r++) {
    for (let c = 0; c < headerData[r].length; c++) {
      let val = headerData[r][c].toString().trim().toLowerCase();
      if (val === "new rank" || val.includes("new rank") || val.includes("calculated rank")) {
        newRankCol = c + 1;
      } else if (val === "rank" || val === "current rank") {
        rankCol = c + 1;
      }
    }
    if (rankCol !== -1 && newRankCol !== -1) {
      headerRow = r + 1;
      break;
    }
  }

  if (headerRow === -1 || rankCol === -1 || newRankCol === -1) return;

  const numRows = lastRow - headerRow;
  if (numRows <= 0) return;

  const newRankValues = sheet.getRange(headerRow + 1, newRankCol, numRows, 1).getValues();
  const currentRankRange = sheet.getRange(headerRow + 1, rankCol, numRows, 1);
  const currentRankValues = currentRankRange.getValues();

  let updatedCount = 0;
  for (let i = 0; i < numRows; i++) {
    let newVal = newRankValues[i][0];
    if (newVal !== "" && newVal !== null && newVal !== undefined) {
      currentRankValues[i][0] = newVal;
      updatedCount++;
    }
  }

  currentRankRange.setValues(currentRankValues);
  SpreadsheetApp.flush();
  ss.toast(`Updated ${updatedCount} members to their new rank on "${sheet.getName()}".`, "✅ Ranks Updated", 5);
}