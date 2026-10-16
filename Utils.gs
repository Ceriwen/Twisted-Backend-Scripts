/**
 * Checks if a note string contains any leave of absence / exemption keywords.
 */
function isExemptNote(noteText) {
  if (!noteText) return false;
  const lower = String(noteText).toLowerCase();
  const keywords = (typeof CONFIG !== "undefined" && CONFIG.TARGETS && CONFIG.TARGETS.EXEMPT_KEYWORDS)
    ? CONFIG.TARGETS.EXEMPT_KEYWORDS
    : ["exempt", "vacation", "loa", "personal", "absence", "sick"];
  return keywords.some(kw => lower.includes(kw.toLowerCase()));
}

/**
 * Utils.gs - Helper Utilities & Custom Mobile Actions
 * TWST Roster Management System
 */

function updateQueueVisualStatus(message, backgroundColor) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let statusCell = sheet.getRange("F1");
    statusCell.setValue(message);
    statusCell.setBackground(backgroundColor);
    statusCell.setFontWeight("bold");
  } catch (e) {}
}

function findRosterRange(sheet) {
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn() || 14;
  let values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    let rowStr = values[i].map(v => v.toString().trim().toLowerCase());
    if (rowStr.includes("member") && rowStr.includes("rank")) { headerRow = i + 1; break; }
  }
  if (headerRow === -1) return { startRow: 1, rowCount: 82, colCount: lastCol };
  
  let trueLastMemberRow = headerRow;
  for (let i = headerRow; i < values.length; i++) {
    let memberVal = values[i] ? values[i].toString().trim() : "";
    let rowStr = values[i].map(v => v.toString().trim().toLowerCase());
    if (rowStr.includes("processed summary") || rowStr.includes("processing summary") || rowStr.includes("matched")) break;
    if (memberVal !== "") trueLastMemberRow = i + 1;
  }
  return { startRow: headerRow, rowCount: Math.max(trueLastMemberRow - headerRow + 1, 2), colCount: lastCol };
}

/**
 * Automatically merges duplicate player rows in the roster by combining daily scores,
 * ranks, and metrics into a single primary row, then deleting secondary duplicate rows.
 */
function deduplicateAndConsolidateRoster(sheetParam) {
  let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let rRange = findRosterRange(sheet);
  if (!rRange || rRange.rowCount <= 1) return 0;

  let headers = sheet.getRange(rRange.startRow, 1, 1, rRange.colCount).getValues().map(h => h.toString().trim().toLowerCase());
  let nameColIdx = headers.findIndex(h => h.includes("member"));
  if (nameColIdx === -1) nameColIdx = 1;

  let dataRange = sheet.getRange(rRange.startRow + 1, 1, rRange.rowCount - 1, rRange.colCount);
  let values = dataRange.getValues();

  let nameMap = {};
  let rowsToDelete = [];
  let mergedCount = 0;

  for (let r = 0; r < values.length; r++) {
    let rawName = values[r][nameColIdx] ? values[r][nameColIdx].toString().trim() : "";
    if (!rawName || rawName.startsWith("---") || rawName.startsWith("⚠️")) continue;

    let cleanKey = normalizeString(rawName);
    if (!cleanKey || cleanKey.length < 3) continue;

    // Check alias dictionary
    let matchedAliasKey = Object.keys(CONFIG.MEMBER_ALIASES || {}).find(
      k => k.toLowerCase() === rawName.toLowerCase() || normalizeString(k) === cleanKey
    );
    if (matchedAliasKey) {
      rawName = CONFIG.MEMBER_ALIASES[matchedAliasKey];
      cleanKey = normalizeString(rawName);
      values[r][nameColIdx] = rawName;
    }

    if (nameMap[cleanKey] !== undefined) {
      // Duplicate found! Merge row r into primaryRowIdx
      let primaryRowIdx = nameMap[cleanKey];
      for (let c = 0; c < rRange.colCount; c++) {
        if (c === nameColIdx) continue;
        let pVal = values[primaryRowIdx][c];
        let dVal = values[r][c];

        if ((pVal === "" || pVal === null || pVal === 0) && (dVal !== "" && dVal !== null && dVal !== 0)) {
          values[primaryRowIdx][c] = dVal;
        } else if (typeof dVal === "number" && typeof pVal === "number") {
          values[primaryRowIdx][c] = Math.max(pVal, dVal);
        }
      }
      rowsToDelete.push(r + rRange.startRow + 1);
      mergedCount++;
    } else {
      nameMap[cleanKey] = r;
    }
  }

  if (mergedCount > 0) {
    dataRange.setValues(values);
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    SpreadsheetApp.flush();
  }
  return mergedCount;
}

function cleanAndSortRoster(sheetParam) {
  let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  let values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 2)).getValues();
  let rowsToDelete = [];
  
  for (let i = 0; i < values.length; i++) {
    let colA = values[i] ? values[i].toString() : "";
    let colB = values[i] ? values[i].toString() : "";
    if (colA.includes("--- MISSING") || colB.includes("--- MISSING") || colA.includes("--- NEW") || colB.includes("--- NEW")) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  deduplicateAndConsolidateRoster(sheet);
}

function menuVerifyAndRestore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  
  cleanAndSortRoster(sheet);
  
  let rRange = findRosterRange(sheet);
  if (!rRange || rRange.rowCount <= 1) return;
  
  let headerValues = sheet.getRange(rRange.startRow, 1, 1, rRange.colCount).getValues();
  let headers = headerValues.map(h => h.toString().trim().toLowerCase());
  
  let nameColIdx = headers.findIndex(h => h.includes("member"));
  if (nameColIdx === -1) nameColIdx = 1;
  let rankColIdx = headers.indexOf("rank");
  
  let numRows = rRange.rowCount - 1;
  
  let tableRange = sheet.getRange(rRange.startRow + 1, 1, numRows, rRange.colCount);
  tableRange.setBackground(null);
  
  let sortCriteria = [];
  if (rankColIdx > -1) {
    sortCriteria.push({ column: rankColIdx + 1, ascending: false });
  }
  sortCriteria.push({ column: nameColIdx + 1, ascending: true });  
  tableRange.sort(sortCriteria);
  try {
    deleteExtraColumns(sheet);
    applyVerificationDropdowns(sheet);
  } catch (e) {}
  SpreadsheetApp.flush();
  
  ss.toast("Roster restored: dividers removed, duplicates merged, all daily fills cleared, and sorted by Rank (Z-A).", "✅ Roster Restored", 5);
}

function getTotalRosterCount(sheet) {
  let range = findRosterRange(sheet);
  if (!range) return 0;
  let values = sheet.getRange(range.startRow + 1, 1, range.rowCount - 1, range.colCount).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    let name = values[i] ? values[i].toString().trim() : "";
    if (name !== "" && !name.startsWith("⚠️") && !name.startsWith("MISSING")) count++;
  }
  return count;
}

function normalizeString(str) {
  if (!str) return "";
  let s = str.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  if (s.includes("☥") || s.includes("♀")) return "symbol_ankh";
  return s.replace(/[\s,._\-👑⚔️🛡️⚠\\[\\]()]/g, "").trim();
}

function stripSquadTag(str) {
  if (!str) return "";
  let clean = normalizeString(str);
  if (clean === "symbol_ankh") return "symbol_ankh";
  return clean.replace(/(?:[a-d][a-b]?)\$/i, "").trim();
}

/**
 * Calculates Levenshtein edit distance between two strings.
 */
function getLevenshteinDistance(a, b) {
  if (!a) return b ? b.length : 0;
  if (!b) return a ? a.length : 0;
  let matrix = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

/**
 * Replaces common leetspeak characters to improve fuzzy OCR matching.
 */
function normalizeLeetspeak(str) {
  if (!str) return "";
  return str.toString().toLowerCase()
    .replace(/0/g, "o")
    .replace(/3/g, "e")
    .replace(/1/g, "i")
    .replace(/5/g, "s")
    .replace(/4/g, "a")
    .replace(/@/g, "a")
    .replace(/\$/g, "s");
}

/**
 * Calculates fuzzy similarity percentage (0.0 to 1.0) between two player names,
 * incorporating word token overlap matching for core handle retention.
 */
function calculateNameSimilarity(str1, str2) {
  let s1 = stripSquadTag(str1);
  let s2 = stripSquadTag(str2);
  if (!s1 || !s2) return 0.0;
  if (s1 === s2) return 1.0;
  
  let l1 = normalizeLeetspeak(s1);
  let l2 = normalizeLeetspeak(s2);
  if (l1 === l2) return 0.95;
  
  if (s1.includes(s2) || s2.includes(s1) || l1.includes(l2) || l2.includes(l1)) {
    let minLen = Math.min(s1.length, s2.length);
    let maxLen = Math.max(s1.length, s2.length);
    if (minLen >= 3) return minLen / maxLen;
  }

  // Token word overlap matching (e.g. "Apostle T Hoarder" vs "Thot Hoarder" sharing "Hoarder")
  let extractWords = function(str) {
    if (!str) return [];
    return str.toString().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9\s]/g, " ")
      .split(/\s+/)
      .filter(w => w.length >= 3 && !/^[a-d][a-b]?\$/i.test(w));
  };

  let w1 = extractWords(str1);
  let w2 = extractWords(str2);
  for (let i = 0; i < w1.length; i++) {
    for (let j = 0; j < w2.length; j++) {
      if (w1[i] === w2[j] || (w1[i].length >= 4 && w2[j].length >= 4 && (w1[i].includes(w2[j]) || w2[j].includes(w1[i])))) {
        let maxWordLen = Math.max(w1[i].length, w2[j].length);
        if (maxWordLen >= 4) return 0.85;
        if (maxWordLen >= 3) return 0.80;
      }
    }
  }
  
  let dist = Math.min(getLevenshteinDistance(s1, s2), getLevenshteinDistance(l1, l2));
  let maxLen = Math.max(s1.length, s2.length);
  return maxLen > 0 ? (maxLen - dist) / maxLen : 0.0;
}

function cleanNumericValues(textStr) {
  if (typeof textStr === "number") return Math.round(textStr);
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
            noteStr.includes("New Member Added") ||
            noteStr.includes("MEMBER NOT DETECTED") ||
            noteStr.includes("Score missing")) {
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

function mobileMenuHandler(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  if (sheet.getName() !== "Data Import") return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const val = e.value;
  
  if (row === 1 && col === 4) {
    // Immediately reset D1 back to placeholder label so onEdit triggers reliably on repeated clicks
    sheet.getRange(1, 4).setValue("📱 Mobile Actions...");
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
    }
  }
}

/**
 * Safely removes temporary divider banner rows ('--- NEW MEMBERS ---' and '--- MISSING / UNMATCHED ---')
 * while keeping all player rows intact before running a new import batch.
 */
function cleanUpDividerBanners(sheet) {
  if (!sheet) return;
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  let range = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 2));
  let values = range.getValues();
  let rowsToDelete = [];
  
  for (let i = 0; i < values.length; i++) {
    let colA = values[i] ? values[i].toString().trim() : "";
    let colB = values[i] ? values[i].toString().trim() : "";
    if (colA.includes("--- MISSING") || colB.includes("--- MISSING") || 
        colA.includes("--- NEW") || colB.includes("--- NEW")) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
  SpreadsheetApp.flush();
}

/**
 * Global spreadsheet edit handler for interactive row-by-row verification and mobile dropdown menu.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Data Import") return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const val = String(e.value || "").trim();
  
  if (row === 1 && col === 4) {
    mobileMenuHandler(e);
    return;
  }
  
  if (row > 2 && col === 13) {
    if (val.includes("Verify Member") || val === "✅ Verify") {
      verifySingleRow(sheet, row);
    } else if (val.includes("Merge") || val === "🔀 Merge") {
      mergeSingleRow(sheet, row);
    } else if (val.includes("Remove Row") || val === "❌ Remove") {
      sheet.deleteRow(row);
      SpreadsheetApp.getActiveSpreadsheet().toast("Deleted unverified row.", "🗑️ Row Removed", 3);
    }
  }
}

/**
 * Verifies a single appended row, moving them into standard active roster status.
 */
function verifySingleRow(sheet, rowIndex) {
  let lastCol = Math.max(sheet.getLastColumn(), 13);
  let rowRange = sheet.getRange(rowIndex, 1, 1, lastCol);
  rowRange.setBackground(null);
  
  let reasonCell = sheet.getRange(rowIndex, 13);
  reasonCell.setValue("Active");
  
  let memberCell = sheet.getRange(rowIndex, 2);
  let memberName = String(memberCell.getValue() || "").trim();
  memberCell.setNote("");
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Verified "${memberName}" as an active member!`, "✅ Member Verified", 4);
}

/**
 * Merges an appended duplicate/alias row into the suggested existing member row.
 */
function mergeSingleRow(sheet, rowIndex) {
  let lastCol = sheet.getLastColumn();
  let rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  let memberName = String(rowValues[1] || "").trim();
  let reasonStr = String(rowValues[12] || "");
  
  let matchStr = "";
  if (reasonStr.includes("Match? ->")) {
    matchStr = reasonStr.split("Match? ->")[1].split("(")[0].trim();
  }
  
  if (!matchStr) {
    SpreadsheetApp.getUi().alert("No suggested candidate match found in Reason column for this row.");
    return;
  }
  
  let lastRow = sheet.getLastRow();
  let allNames = sheet.getRange(1, 2, lastRow, 1).getValues();
  let targetRowIndex = -1;
  
  for (let r = 0; r < allNames.length; r++) {
    if (r + 1 === rowIndex) continue;
    let name = String(allNames[r][0] || "").trim();
    if (normalizeString(name) === normalizeString(matchStr)) {
      targetRowIndex = r + 1;
      break;
    }
  }
  
  if (targetRowIndex === -1) {
    SpreadsheetApp.getUi().alert(`Could not locate target member "${matchStr}" in the roster to merge into.`);
    return;
  }
  
  let targetRange = sheet.getRange(targetRowIndex, 1, 1, lastCol);
  let targetValues = targetRange.getValues()[0];
  
  for (let c = 2; c <= 9; c++) {
    let sourceScore = cleanNumericValues(rowValues[c]);
    let targetScore = cleanNumericValues(targetValues[c]);
    if (sourceScore > 0 && targetScore === 0) {
      targetValues[c] = sourceScore;
    }
  }
  
  targetRange.setValues([targetValues]);
  sheet.deleteRow(rowIndex);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Merged scores from "${memberName}" into "${matchStr}"!`, "🔀 Merged Successfully", 5);
}

/**
 * Automatically purges and deletes any extra columns beyond Column M (Col 13)
 * such as temporary helper columns N:S (Cols 14-19).
 */
function deleteExtraColumns(sheetParam) {
  let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;
  let maxCols = sheet.getMaxColumns();
  if (maxCols > 13) {
    try {
      sheet.deleteColumns(14, maxCols - 13);
      SpreadsheetApp.flush();
    } catch (e) {}
  }
}

/**
 * Applies native Google Sheets Data Validation dropdowns (✅ Verify, 🔀 Merge, ❌ Remove)
 * to Column M (Reason/Action) for all newly appended or match-candidate member rows.
 */
function applyVerificationDropdowns(sheetParam) {
  let sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import") || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;
  
  let rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["✅ Verify", "🔀 Merge", "❌ Remove"], true)
    .setAllowInvalid(true)
    .build();
    
  let rosterRange = findRosterRange(sheet);
  let values = sheet.getRange(rosterRange.startRow + 1, 1, rosterRange.rowCount - 1, 13).getValues();
  let nameNotes = sheet.getRange(rosterRange.startRow + 1, 2, rosterRange.rowCount - 1, 1).getNotes();
  
  let appliedCount = 0;
  for (let r = 0; r < values.length; r++) {
    let rowIndex = rosterRange.startRow + 1 + r;
    let colA = String(values[r][0] || "").trim();
    let colB = String(values[r][1] || "").trim();
    let reason = String(values[r][12] || "").trim();
    let note = String(nameNotes[r][0] || "").trim();
    
    if (reason.includes("Match? ->") || reason.includes("New Member") || note.includes("New Member") || colA.includes("---") || colB.includes("---") || reason === "" || reason.includes("Unverified")) {
      sheet.getRange(rowIndex, 13).setDataValidation(rule);
      appliedCount++;
    }
  }
  return appliedCount;
}
