/**
 * Utils.gs - Shared Utility Functions & Mobile Event Handlers
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

function menuVerifyAndRestore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();

  cleanAndSortRoster(sheet);
  deleteExtraColumns(sheet);

  let rRange = findRosterRange(sheet);
  if (!rRange || rRange.rowCount <= 1) return;

  let headerValues = sheet.getRange(rRange.startRow, 1, 1, rRange.colCount).getValues()[0];
  let headers = headerValues.map(h => h.toString().trim().toLowerCase());

  let nameColIdx = headers.findIndex(h => h.includes("member"));
  if (nameColIdx === -1) nameColIdx = 1;
  let rankColIdx = headers.indexOf("rank");

  let numRows = rRange.rowCount - 1;
  let nameCol = nameColIdx + 1;

  sheet.getRange(rRange.startRow + 1, nameCol, numRows, 1).setBackground(null);

  let sortCriteria = [];
  if (rankColIdx > -1) {
    sortCriteria.push({ column: rankColIdx + 1, ascending: false });
  }
  sortCriteria.push({ column: nameColIdx + 1, ascending: true });

  sheet.getRange(rRange.startRow + 1, 1, numRows, rRange.colCount).sort(sortCriteria);
  SpreadsheetApp.flush();

  ss.toast("Roster restored: dividers removed, fills cleared, and sorted by Rank (Z-A).", "✅ Roster Restored", 5);
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
  if (s.includes("☥") || s.includes("♀")) return "symbol_ankh";
  return s.replace(/[\s,.\-_👑⚔️🛡️⚠️\[\]\(\)]/g, "").trim();
}

function stripSquadTag(str) {
  if (!str) return "";
  let clean = normalizeString(str);
  if (clean === "symbol_ankh") return "symbol_ankh";
  return clean.replace(/(?:[a-d][1-4][a-b]?)$/i, "").trim();
}

function cleanNumericValues(textStr) {
  if (typeof textStr === 'number') return Math.round(textStr);
  if (!textStr) return 0;
  let normalized = textStr.toString().toUpperCase().trim().replace(/\s+/g, "").replace(/O/g, "0");
  if (normalized === "-" || normalized === "—" || normalized === "") return 0;
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
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<textarea readonly style="width:100vw;height:100vh;background:#1e1e1e;color:#9cdcfe;padding:10px;border:none;">${text}</textarea>`).setWidth(800).setHeight(500), "Debug Log");
}

function clearAllHighlightsAndNotes() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let rRange = findRosterRange(sheet);

  if (!rRange || rRange.rowCount <= 1) return;

  let tableRange = sheet.getRange(rRange.startRow + 1, 1, rRange.rowCount - 1, rRange.colCount);
  tableRange.setBackground(null);

  let notes = tableRange.getNotes();
  let notesModified = false;

  for (let r = 0; r < notes.length; r++) {
    for (let c = 0; c < notes[r].length; c++) {
      let noteStr = notes[r][c];
      if (noteStr) {
        if (noteStr.includes("Matched & Verified") ||
            noteStr.includes("Missing from all") ||
            noteStr.includes("MEMBER NOT DETECTED") ||
            noteStr.includes("New Member Added")) {
          notes[r][c] = "";
          notesModified = true;
        }
      }
    }
  }

  if (notesModified) {
    tableRange.setNotes(notes);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("Background fills and import notes cleared (manual notes preserved).", "🧹 Cleared Fills", 4);
}

function menuUpdateRankWithNewRank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    ss.toast("Sheet appears empty.", "⚠️ Notice", 3);
    return;
  }

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

  if (headerRow === -1 || rankCol === -1 || newRankCol === -1) {
    SpreadsheetApp.getUi().alert(
      "Column Detection Failed",
      `Could not automatically locate both 'Rank' and 'New Rank' headers on "${sheet.getName()}".\n\n` +
      `Found: Rank (Col ${rankCol !== -1 ? rankCol : "Missing"}), New Rank (Col ${newRankCol !== -1 ? newRankCol : "Missing"}).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const numRows = lastRow - headerRow;
  if (numRows <= 0) {
    ss.toast("No roster rows found below the header.", "⚠️ Notice", 3);
    return;
  }

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

  ss.toast(`Updated ${updatedCount} member ranks on "${sheet.getName()}".`, "✅ Ranks Updated", 4);
}

/**
 * Mobile Dropdown Handler (Installable onEdit trigger)
 * Listens for actions selected in cell D1 of the "Data Import" tab.
 */
function mobileMenuHandler(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();

  if (sheet.getName() !== "Data Import") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const val = String(e.value || "").trim();

  // D1 is Row 1, Column 4
  if (row === 1 && col === 4) {
    e.range.setValue("⚙️ Select Action...");
    SpreadsheetApp.flush();

    if (val === "📥 Sort & Process Inbox") {
      sortAndProcessInboxFolder();
    } else if (val === "🚀 Process Next Folder") {
      processNextAvailableFolder();
    } else if (val === "🔮 Calculate Weekly Ranks") {
      calculateWeeklyRankAdjustments();
    } else if (val === "📅 Run Sunday Reset") {
      executeSundayEowReset();
    } else if (val === "⏱️ Toggle Auto Watcher") {
      toggleAutoImportWatcher();
    } else if (val === "🔄 Sync Code to GitHub" && typeof syncProjectToGitHub === "function") {
      syncProjectToGitHub();
    }
  }
}