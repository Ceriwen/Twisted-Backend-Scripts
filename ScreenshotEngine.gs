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
      : headers.findIndex(item => item === currentLabel.toLowerCase() || item.includes(currentLabel.toLowerCase()));
    let rankColIdx = headers.findIndex(item => item === "rank");

    if (nameColIdx === -1) {
      debugSheet.appendRow([new Date(), `❌ 'Member' column not found in sheet.`]);
      continue;
    }

    let flaggedRows = [];
    if (typeof findEliminationFlaggedRows === "function") {
      flaggedRows = findEliminationFlaggedRows(sheet, sheetData, nameColIdx);
    }

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

      let ocrResults = [];
      try {
        ocrResults = parseLeaderboardBatchWithGemini(batchFiles, currentLabel);
      } catch (geminiErr) {
        debugSheet.appendRow([new Date(), `❌ Gemini Parse Error on ${batchLabelStr}: ${geminiErr.message}`]);
        continue;
      }

      let matchedRows = new Set();
      let verificationLogs = {};
      let updatedProfileCells = [];
      let dynamicRankChanges = [];

      ocrResults.forEach(item => {
        let cleanCandidateName = stripSquadTag(item.name);
        let scoreVal = cleanNumericValues(item.score);
        let rankVal = item.rank ? item.rank.toString().trim().toUpperCase() : "";

        let matchedRowIdx = -1;
        for (let r = 1; r < sheetData.length; r++) {
          let rowMemberName = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
          if (normalizeString(rowMemberName) === normalizeString(cleanCandidateName)) {
            matchedRowIdx = r;
            break;
          }
        }

        if (matchedRowIdx > -1) {
          let actualRow = startRow + matchedRowIdx;
          matchedRows.add(actualRow);
          sheetData[matchedRowIdx][nameColIdx] = cleanCandidateName;
          if (metricColIdx > -1) sheetData[matchedRowIdx][metricColIdx] = scoreVal;

          cumulativeMatchedRows.add(actualRow);
          auditMatched.push({ name: cleanCandidateName, score: scoreVal, raw: item.name });

          if (rankColIdx > -1 && rankVal && currentLabel === "Strength Ranking - Power") {
            let existingRank = sheetData[matchedRowIdx][rankColIdx] ? sheetData[matchedRowIdx][rankColIdx].toString().trim().toUpperCase() : "";
            if (existingRank && existingRank !== rankVal && existingRank !== "R5" && existingRank !== "R4") {
              rankDiscrepancies.push({ row: actualRow, name: cleanCandidateName, sheetRank: existingRank, ocrRank: rankVal });
            } else if (!existingRank && !ignoreRanks) {
              sheetData[matchedRowIdx][rankColIdx] = rankVal;
            }
          }
        } else {
          auditMissing.push({ name: cleanCandidateName, score: scoreVal, category: currentLabel });
        }
      });

      sheet.getRange(startRow, 1, totalCurrentRows, rosterRange.colCount).setValues(sheetData);

      if (typeof paintBatchFormatting === "function") {
        paintBatchFormatting(sheet, sheetData, nameColIdx, metricColIdx, rankColIdx, flaggedRows, matchedRows, cumulativeMatchedRows, isStrengthRanking, updatedProfileCells, verificationLogs, dynamicRankChanges, currentLabel);
      }

      batchFiles.forEach(file => {
        try {
          if (targetProcessedSubfolder) file.moveTo(targetProcessedSubfolder);
          else file.setTrashed(true);
        } catch (e) {}
      });

      totalProcessedFiles += batchFiles.length;
    }
  }

  try {
    if (typeof writeToImportAuditTab === "function") {
      writeToImportAuditTab(headingLabelsArray[0], auditMatched, auditNewMembers, auditMissing, rankDiscrepancies);
    }
  } catch (e) {}

  if (returnStats) {
    return { timedOut: timeoutTriggered, processedCount: totalProcessedFiles };
  }
}

/**
* Scans folders in sequence and processes the next folder containing screenshots with F1 visual updates.
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

  updateQueueVisualStatus(`⚡ Processing Folder: ${targetCategory}`, "#fff2cc");
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
* Scans the main Inbox folder, classifies unorganized screenshots using Gemini OCR, 
* and routes them into their respective category folders with F1 visual updates.
*/
function sortAndProcessInboxFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");

  if (!CONFIG.INBOX_FOLDER_ID || CONFIG.INBOX_FOLDER_ID.startsWith("YOUR_")) {
    updateQueueVisualStatus("⚠️ Error: INBOX_FOLDER_ID missing.", "#f4cccc");
    throw new Error("INBOX_FOLDER_ID is not configured in Config.gs.");
  }

  updateQueueVisualStatus("📥 Scanning Inbox folder...", "#fff2cc");
  let inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  let files = inboxFolder.getFiles();
  let processedCount = 0;
  let totalFiles = 0;

  let fileList = [];
  while (files.hasNext()) {
    let f = files.next();
    let mime = f.getMimeType();
    if (mime.includes("image") || mime.includes("octet-stream") || f.getName().toLowerCase().endsWith(".jpg") || f.getName().toLowerCase().endsWith(".png")) {
      fileList.push(f);
    }
  }

  totalFiles = fileList.length;
  if (totalFiles === 0) {
    updateQueueVisualStatus("✅ Inbox is empty - Ready", "#d9ead3");
    ss.toast("Inbox folder has no images to sort.", "📥 Inbox Empty", 4);
    return { success: true, count: 0 };
  }

  for (let i = 0; i < fileList.length; i++) {
    let f = fileList[i];
    try {
      updateQueueVisualStatus(`📥 Sorting (${i + 1}/${totalFiles}): ${f.getName()}`, "#fff2cc");
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

  updateQueueVisualStatus(`✅ Sorted ${processedCount} of ${totalFiles} screenshots!`, "#d9ead3");
    ss.toast(`Successfully sorted ${processedCount} screenshots from inbox.`, "📥 Inbox Sorted", 5);

  // AUTOMATIC CHAINING: Immediately proceed to process the sorted folders!
  if (processedCount > 0 && typeof processNextAvailableFolder === "function") {
    updateQueueVisualStatus("⚡ Moving to process category folders...", "#fff2cc");
    Utilities.sleep(1000);
    processNextAvailableFolder();
  }

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
  return parseLeaderboardBatchWithGeminstats([file], currentLabel);
}

function stripSquadTag(str) {
  if (!str) return "";
  return str.toString().replace(/\[.*?\]\s*/g, "").trim();
}

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
    <button onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).restoreProcessedScreenshotsAction(document.getElementById('sel').value)" style="background:#2563eb;color:white;padding:10px 15px;border:none;border-radius:5px;cursor:pointer;">Restore</button>
    </body></html>
  `).setTitle("Restore Processed Screenshots").setWidth(350).setHeight(200));
}

function restoreProcessedScreenshotsAction(categoryKey) {
  let master;
  try { master = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID); } catch (e) { return; }
  let restored = 0;
  let keys = categoryKey === "ALL" ? Object.keys(CONFIG.PROCESSED_SUBFOLDER_NAMES) : [categoryKey];
  keys.forEach(k => {
    let subName = CONFIG.PROCESSED_SUBFOLDER_NAMES[k] || k;
    let sub = master.getFoldersByName(subName);
    if (sub.hasNext()) {
      let files = sub.next().getFiles();
      let target = DriveApp.getFolderById(CONFIG.FOLDER_MAPPINGS[k]);
      while (files.hasNext()) { try { files.next().moveTo(target); restored++; } catch(e) {} }
    }
  });
  SpreadsheetApp.getUi().alert("Restoration Complete", `Restored ${restored} file(s).`, SpreadsheetApp.getUi().ButtonSet.OK);
}