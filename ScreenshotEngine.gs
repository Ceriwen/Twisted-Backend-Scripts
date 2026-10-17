/**
 * ScreenshotEngine.gs - Automated screenshot sorting and Gemini OCR ingestion pipeline
 * TWST Roster Management System
 */

function sortAndProcessInboxFolder() {
  if (CONFIG.INBOX_FOLDER_ID && !CONFIG.INBOX_FOLDER_ID.startsWith("YOUR_")) {
    let inboxFolder;
    try { 
      inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID); 
      let files = inboxFolder.getFiles();
      let sortedCount = 0;
      
      updateQueueVisualStatus("⚡ Sorting Inbox...", "#fff2cc");
      
      while (files.hasNext()) {
        let f = files.next();
        let mime = f.getMimeType();
        if (mime === MimeType.PNG || mime === MimeType.JPEG || mime.includes("image")) {
          try {
            let cat = classifyScreenshotWithGemini(f);
            if (cat && CONFIG.FOLDER_MAPPINGS[cat]) {
              f.moveTo(DriveApp.getFolderById(CONFIG.FOLDER_MAPPINGS[cat]));
              sortedCount++;
            }
          } catch (e) {}
        }
      }
      
      let timeStr = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      
      if (sortedCount > 0) {
        try { SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully sorted ${sortedCount} screenshots from Inbox. Moving to process...`, "📥 Inbox Sorted", 4); } catch (e) {}
        updateQueueVisualStatus(`✅ Sorted ${sortedCount} files. Processing next folder...`, "#d9ead3");
      } else {
        updateQueueVisualStatus(`👀 Watcher Active - Checked Inbox at ${timeStr}. Checking for folders to process...`, "#e8f0fe"); 
      }
      
      // Chain directly into processing the next available folder
      processNextAvailableFolder();
      
    } catch (e) {
      updateQueueVisualStatus("⚠️ Error accessing Inbox.", "#f4cccc");
      processNextAvailableFolder();
    }
  } else {
    processNextAvailableFolder();
  }
}

function classifyScreenshotWithGemini(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return null;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const blob = file.getBlob();
  const payload = {
    contents: [{ parts: [{ text: `Classify this Last War alliance screenshot into ONE: "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Strength Ranking - Power", "Strength Ranking - Donations". Return ONLY the exact string.` }, { inlineData: { mimeType: blob.getContentType() || "image/jpeg", data: Utilities.base64Encode(blob.getBytes()) } }] }],
    generationConfig: { temperature: 0.1 }
  };
  try {
    let res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
    let json = JSON.parse(res.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      let cat = json.candidates[0].content.parts[0].text.trim().replace(/^["']|["']$/g, "");
      if (CONFIG.FOLDER_MAPPINGS[cat]) return cat;
    }
  } catch (e) {}
  return null;
}

function processNextAvailableFolder() {
  const potentialFolders = ["Strength Ranking - Power", "Strength Ranking - Donations", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");
  
  let targetCategory = null;
  
  for (let i = 0; i < potentialFolders.length; i++) {
    let cat = potentialFolders[i];
    let folderId = CONFIG.FOLDER_MAPPINGS[cat];
    if (folderId && !folderId.startsWith("YOUR_")) {
      try {
        let folder = DriveApp.getFolderById(folderId);
        let files = folder.getFiles();
        let hasImages = false;
        while (files.hasNext()) {
          let f = files.next();
          let mime = f.getMimeType();
          if (mime === MimeType.PNG || mime === MimeType.JPEG || mime.includes("image")) {
            hasImages = true;
            break;
          }
        }
        if (hasImages) { targetCategory = cat; break; }
      } catch (e) {}
    }
  }
  
  if (!targetCategory) {
    let timeStr = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    updateQueueVisualStatus(`✅ Inbox and Folders Empty. (Last run: ${timeStr})`, "#d9ead3");
    return;
  }
  
  let statusText = `⚡ Processing Folder: ${targetCategory}`;
  updateQueueVisualStatus(statusText, "#fff2cc");
  try { ss.toast(statusText, "🔄 Processing Next Folder", 5); } catch (e) {}
  debugSheet.appendRow([new Date(), `Processing next available folder: "${targetCategory}"`]);
  
  try {
    let result = processFolderDirectly([targetCategory], true, false);
    if (!result || !result.timedOut) {
      debugSheet.appendRow([new Date(), `Successfully finished "${targetCategory}". Matched: ${result ? result.matched : 'N/A'}`]);
      runRankCalculationEngine(["mon", "tue", "wed", "thu", "fri", "sat", "donations"], "Post-Import Auto-Calculation");
      updateQueueVisualStatus("✅ Completed and Ready for next Batch", "#d9ead3");
    } else {
      updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
    }
  } catch (err) {
    debugSheet.appendRow([new Date(), `CRITICAL ERROR / TIMEOUT on "${targetCategory}": ` + err.message]);
    updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
    try { ss.toast(`Processing interrupted on "${targetCategory}". Click 'Process Next Folder' again to resume remaining files.`, "⚠️ Interrupted", 10); } catch (e) {}
  }
}

function processFolderDirectly(headingLabelsArray, returnStats = false, ignoreRanks = false) {
  const START_TIME = Date.now();
  const MAX_RUNTIME_MS = 210000; // 3.5 minutes safety guardrail
  let timeoutTriggered = false;  

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  
  try {
    let tempSnap = ss.getSheetByName("_ImportBackup");
    if (!tempSnap) {
      tempSnap = ss.insertSheet("_ImportBackup");
      let srcRange = sheet.getDataRange();
      srcRange.copyTo(tempSnap.getRange(1, 1));
      srcRange.copyTo(tempSnap.getRange(1, 1), {contentsOnly: true});
      tempSnap.hideSheet();
    }
  } catch(e) {}
  
  // Strip any previous divider rows before starting import
  if (typeof cleanAndSortRoster === "function") {
    cleanAndSortRoster(sheet);
  }
  
  let startRange = findRosterRange(sheet);
  if (startRange && startRange.rowCount > 1) {
    let headers = sheet.getRange(startRange.startRow, 1, 1, startRange.colCount).getValues()[0].map(h => h.toString().trim().toLowerCase());
    let nIdx = headers.findIndex(h => h.includes("member"));
    if (nIdx === -1) nIdx = 1;
    sheet.getRange(startRange.startRow + 1, nIdx + 1, startRange.rowCount - 1, 1).setBackground(null);
  }
  
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");
  let totalProcessedFiles = 0;
  let batchUniqueMatched = new Set();
  let batchAppended = 0;
  
  for (let h = 0; h < headingLabelsArray.length; h++) {
    if (timeoutTriggered) break; 

    let currentLabel = headingLabelsArray[h];
    let folderId = CONFIG.FOLDER_MAPPINGS[currentLabel];
    if (!folderId || folderId.startsWith("YOUR_")) continue;
    
    let folder;
    try { folder = DriveApp.getFolderById(folderId); } catch(e) { continue; }
    
    let files = folder.getFiles();
    let folderFiles = [];
    while (files.hasNext()) {
      let f = files.next();
      let mime = f.getMimeType();
      if (mime === MimeType.PNG || mime === MimeType.JPEG || mime.includes("image")) folderFiles.push(f);
    }
    
    if (folderFiles.length === 0) continue;
    folderFiles.sort((a, b) => a.getName().localeCompare(b.getName()));
    
    let targetProcessedSubfolder = null;
    try {
      let masterProcessedFolder = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID);
      let targetSubfolderName = CONFIG.PROCESSED_SUBFOLDER_NAMES[currentLabel] || currentLabel;
      let existingSubs = masterProcessedFolder.getFoldersByName(targetSubfolderName);
      targetProcessedSubfolder = existingSubs.hasNext() ? existingSubs.next() : masterProcessedFolder.createFolder(targetSubfolderName);
    } catch (err) {}
    
    let isStrengthRanking = (currentLabel === "Strength Ranking - Power" || currentLabel === "Strength Ranking - Donations");
    let cumulativeMatchedRows = new Set();
    let cumulativeVerificationLogs = {};
    let cumulativeUpdatedRows = new Set();
    let cumulativeFlaggedRows = [];
    
    let addedNamesList = [];
    let metricColIdx = -1;
    
    let rosterRange = findRosterRange(sheet);
    if (!rosterRange) continue;
    
    let sheetData = sheet.getRange(rosterRange.startRow, 1, rosterRange.rowCount, rosterRange.colCount).getValues();
    const headers = sheetData[0].map(h => h.toString().trim().toLowerCase());
    
    let nameColIdx = headers.findIndex(h => h.includes("member"));
    if (nameColIdx === -1) nameColIdx = 1; 
    let rankColIdx = headers.indexOf("rank");
    
    let possibleAliases = CONFIG.HEADER_ALIASES[currentLabel] || [currentLabel.toLowerCase()];
    metricColIdx = -1;
    for (let a = 0; a < possibleAliases.length; a++) {
      metricColIdx = headers.indexOf(possibleAliases[a]);
      if (metricColIdx > -1) break;
    }
    if (metricColIdx === -1) continue;

    for (let fIdx = 0; fIdx < folderFiles.length; fIdx++) {
      if (Date.now() - START_TIME > MAX_RUNTIME_MS) {
        timeoutTriggered = true;
        updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
        try { ss.toast("Nearing Google's execution limit. Saving current progress...", "⏳ Graceful Pause", 8); } catch (e) {}
        debugSheet.appendRow([new Date(), `⚠️ Graceful Exit Triggered: Hit 3.5m limit on file ${fIdx + 1}`]);
        break; 
      }

      let file = folderFiles[fIdx];
      let fileName = file.getName();
      let fileNumStr = `[${currentLabel}] File ${fIdx + 1}/${folderFiles.length}: ${fileName}`;
      
      updateQueueVisualStatus(`⚡ Processing: ${fileNumStr}`, "#fff2cc");
      debugSheet.appendRow([new Date(), `Scanning file -> ${fileNumStr}`]);
      
      let rosterNames = sheetData.slice(1)
        .map((row, idx) => ({ 
          rawName: row[nameColIdx] ? row[nameColIdx].toString().trim() : "", 
          cleanName: normalizeString(row[nameColIdx]), 
          baseName: stripSquadTag(row[nameColIdx]),
          index: idx + 1 
        }))
        .filter(obj => obj.cleanName.length >= 1);
      
      let parsedPlayers;
      try {
        parsedPlayers = parseLeaderboardWithGemini(file, currentLabel);
        debugSheet.appendRow([new Date(), `Gemini extracted ${parsedPlayers.length} entries`]);
      } catch (geminiErr) {
        debugSheet.appendRow([new Date(), `❌ Gemini Parse Error on ${fileName}: ${geminiErr.message}`]);
        continue;
      }
      
      let fileMatches = 0;
      parsedPlayers.forEach(p => {
        let cleanCandidateName = p.name ? p.name.trim() : "";
        if (!cleanCandidateName) return;
        
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
            matchedRowIdx = ro.index; break;
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
              matchedRowIdx = ro.index; break;
            }
          }
        }
        
        if (matchedRowIdx !== -1) {
          let actualRow = matchedRowIdx + 1; 
          sheetData[matchedRowIdx][nameColIdx] = cleanCandidateName;
          if (metricColIdx > -1) sheetData[matchedRowIdx][metricColIdx] = scoreVal;
          
          cumulativeMatchedRows.add(actualRow);
          cumulativeUpdatedRows.add(actualRow);
          batchUniqueMatched.add(actualRow);
          fileMatches++;
          
          cumulativeVerificationLogs[actualRow] = `✅ Matched & Verified via ${currentLabel} (${fileName})\nScore: ${scoreVal}`;
          
          if (rankColIdx > -1 && rankVal && currentLabel !== "Strength Ranking - Donations" && !ignoreRanks) {
            let existingRank = sheetData[matchedRowIdx][rankColIdx] ? sheetData[matchedRowIdx][rankColIdx].toString().trim().toUpperCase() : "";
            if (existingRank !== "R5" && existingRank !== "R4") {
              sheetData[matchedRowIdx][rankColIdx] = rankVal;
            }
          }
        } else {
          let alreadyAdded = sheetData.some(row => stripSquadTag(row[nameColIdx]) === candidateBase);
          if (!alreadyAdded && cleanCandidateName.length >= 1) {
            let newRowTemplate = Array(sheetData[0].length).fill("");
            newRowTemplate[nameColIdx] = cleanCandidateName;
            if (rankColIdx > -1 && !ignoreRanks) newRowTemplate[rankColIdx] = rankVal || "R3";
            if (metricColIdx > -1) newRowTemplate[metricColIdx] = scoreVal;
            
            sheetData.push(newRowTemplate);
            let freshRowIdx = sheetData.length;
            cumulativeFlaggedRows.push(freshRowIdx);
            cumulativeMatchedRows.add(freshRowIdx);
            cumulativeUpdatedRows.add(freshRowIdx);
            batchUniqueMatched.add(freshRowIdx);
            batchAppended++;
            fileMatches++;
            
            addedNamesList.push(cleanCandidateName); 
            cumulativeVerificationLogs[freshRowIdx] = `⚠️ New Member Added via ${currentLabel} (${fileName})\nScore: ${scoreVal}`;
          }
        }
      });
      
      // --- PER-FILE INCREMENTAL SPREADSHEET WRITES ---
      if (sheetData.length > 1) {
        let nameValues = sheetData.slice(1).map(row => [row[nameColIdx]]);
        sheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, nameValues.length, 1).setValues(nameValues);
      }
      
      if (metricColIdx > -1 && sheetData.length > 1) {
        let metricValues = sheetData.slice(1).map(row => [row[metricColIdx]]);
        sheet.getRange(rosterRange.startRow + 1, metricColIdx + 1, metricValues.length, 1).setValues(metricValues);
      }
      
      if (isStrengthRanking && rankColIdx > -1 && sheetData.length > 1 && !ignoreRanks) {
        let rankValues = sheetData.slice(1).map(row => [row[rankColIdx]]);
        sheet.getRange(rosterRange.startRow + 1, rankColIdx + 1, rankValues.length, 1).setValues(rankValues);
      }
      
      if (sheetData.length > rosterRange.rowCount) {
        let addedRowsBatch = sheetData.slice(rosterRange.rowCount);
        let insertRow = rosterRange.startRow + rosterRange.rowCount;
        sheet.getRange(insertRow, 1, addedRowsBatch.length, addedRowsBatch[0].length).setValues(addedRowsBatch);
        rosterRange.rowCount = sheetData.length;
      }
      
      paintBatchFormatting(sheet, sheetData, nameColIdx, metricColIdx, rankColIdx, cumulativeFlaggedRows, cumulativeMatchedRows, cumulativeUpdatedRows, isStrengthRanking, [], cumulativeVerificationLogs, [], currentLabel);
      SpreadsheetApp.flush(); // Ensure Sheet commits changes after each screenshot
      
      debugSheet.appendRow([new Date(), `Committed ${fileName} -> Matched/Added: ${fileMatches} players.`]);
      totalProcessedFiles++;
      
      if (targetProcessedSubfolder) {
        try { file.moveTo(targetProcessedSubfolder); } catch (e) {}
      }
    }

    // --- 100% FINISHED SORTING & GROUPING ---
    if (!timeoutTriggered) {
      let finalRosterRange = findRosterRange(sheet);
      let finalSheetData = sheet.getRange(finalRosterRange.startRow, 1, finalRosterRange.rowCount, finalRosterRange.colCount).getValues();
      let nIdx = finalSheetData[0].map(h => h.toString().trim().toLowerCase()).indexOf("member");
      if (nIdx === -1) nIdx = 1;
      
      let finalRosterNotes = sheet.getRange(finalRosterRange.startRow, nIdx + 1, finalRosterRange.rowCount, 1).getNotes();
      let metricValues = (metricColIdx > -1) ? sheet.getRange(finalRosterRange.startRow + 1, metricColIdx + 1, finalRosterRange.rowCount - 1, 1).getValues() : [];
      
      let globalMissingNames = [];
      let globalAddedNames = [];
      
      for (let r = 1; r < finalSheetData.length; r++) {
        let memberName = finalSheetData[r][nIdx] ? finalSheetData[r][nIdx].toString().trim() : "";
        if (!memberName || memberName.startsWith("⚠️") || memberName.startsWith("---")) continue;
        
        let note = finalRosterNotes[r] ? finalRosterNotes[r][0].toLowerCase() : "";
        let isExempt = isExemptNote(note);
        let scoreCell = (metricValues[r - 1]) ? metricValues[r - 1][0] : "";
        
        let isNew = addedNamesList.some(an => normalizeString(an) === normalizeString(memberName));
        
        if (isNew) {
          globalAddedNames.push(memberName);
        } else if ((scoreCell === "" || scoreCell === null || scoreCell === 0) && !isExempt) {
          globalMissingNames.push(memberName);
          if (metricColIdx > -1 && !isStrengthRanking) {
             sheet.getRange(finalRosterRange.startRow + r, metricColIdx + 1).setBackground(CONFIG.COLORS.MISSING_DATA).setNote(`❌ Missing from all ${currentLabel} screenshots`);
          }
        }
      }
      
      if (globalAddedNames.length > 0 || globalMissingNames.length > 0) {
        let sortKeys = [];
        for (let r = 1; r < finalSheetData.length; r++) {
          let name = finalSheetData[r][nIdx] ? finalSheetData[r][nIdx].toString().trim() : "";
          if (!name) { sortKeys.push([0]); continue; }
          
          if (globalAddedNames.includes(name)) {
            sortKeys.push([3]); 
          } else if (globalMissingNames.includes(name)) {
            sortKeys.push([2]); 
          } else {
            sortKeys.push([1]); 
          }
        }
        
        let sortCol = finalRosterRange.colCount + 1;
        let maxCols = sheet.getMaxColumns();
        if (sortCol > maxCols) sheet.insertColumnAfter(maxCols); 
        
        sheet.getRange(finalRosterRange.startRow + 1, sortCol, sortKeys.length, 1).setValues(sortKeys);
        sheet.getRange(finalRosterRange.startRow + 1, 1, finalRosterRange.rowCount - 1, sortCol).sort(sortCol);
        sheet.getRange(finalRosterRange.startRow + 1, sortCol, finalRosterRange.rowCount - 1, 1).clearContent();
        
        if (typeof deleteExtraColumns === "function") {
          deleteExtraColumns(sheet);
        }
        
        let sortedNames = sheet.getRange(finalRosterRange.startRow + 1, nIdx + 1, finalRosterRange.rowCount - 1, 1).getValues();
        let firstMissingRow = -1;
        let firstNewRow = -1;
        
        for (let i = 0; i < sortedNames.length; i++) {
           let n = sortedNames[i][0].toString().trim();
           if (n && globalMissingNames.includes(n) && firstMissingRow === -1) firstMissingRow = finalRosterRange.startRow + 1 + i;
           if (n && globalAddedNames.includes(n) && firstNewRow === -1) firstNewRow = finalRosterRange.startRow + 1 + i;
        }
        
        if (firstNewRow !== -1) {
           sheet.insertRowBefore(firstNewRow);
           sheet.getRange(firstNewRow, 1, 1, finalRosterRange.colCount)
               .setBackground("#d9ead3").setFontColor("black").setFontWeight("bold")
               .setValues([Array(finalRosterRange.colCount).fill("").fill("--- NEW MEMBERS ---", nIdx, nIdx+1)]);
        }
        if (firstMissingRow !== -1) {
           sheet.insertRowBefore(firstMissingRow);
           sheet.getRange(firstMissingRow, 1, 1, finalRosterRange.colCount)
               .setBackground("#f4cccc").setFontColor("black").setFontWeight("bold")
               .setValues([Array(finalRosterRange.colCount).fill("").fill("--- MISSING / UNMATCHED ---", nIdx, nIdx+1)]);
        }
      }
    }
  }
  
  if (!returnStats && headingLabelsArray.length === 1) {
    try {
      let ui = SpreadsheetApp.getUi();
      if (timeoutTriggered) {
        ui.alert("⚠️ Batch Paused", `Processed ${totalProcessedFiles} screenshots for "${headingLabelsArray[0]}".\n\nGoogle limits scripts to 6 minutes. The script gracefully paused to save your progress. Run again from the menu to resume!`, ui.ButtonSet.OK);
      } else {
        ui.alert("✅ Import Complete", `Successfully processed "${headingLabelsArray[0]}". Unmatched and New members have been grouped for your review. Use "Verify & Restore Roster Order" to return to normal!`, ui.ButtonSet.OK);
      }
    } catch (e) {}
  } else if (!returnStats && timeoutTriggered) {
    try { SpreadsheetApp.getActiveSpreadsheet().toast(`Batch paused early due to time limits. Run again to resume!`, "⚠️ Paused", 8); } catch (e) {}
  }
  
  return { processedFiles: totalProcessedFiles, matched: batchUniqueMatched.size, appended: batchAppended, timedOut: timeoutTriggered };
}

function rollbackLastImport() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let tempSnap = ss.getSheetByName("_ImportBackup");
  
  if (!tempSnap) {
    ui.alert("⚠️ Rollback Unavailable", "No recent import backup found in this session.", ui.ButtonSet.OK);
    return;
  }
  
  if (ui.alert("↩ Confirm Rollback", "This will revert your Data Import sheet to the exact state it was in before your last import, and move any processed screenshots back to their input folders. Proceed?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  sheet.clear();
  tempSnap.getDataRange().copyTo(sheet.getRange(1, 1));
  
  ss.deleteSheet(tempSnap);
  
  let master;
  try { master = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID); } catch(e) {}
  if (master) {
    let keys = Object.keys(CONFIG.PROCESSED_SUBFOLDER_NAMES);
    keys.forEach(k => {
      let sub = master.getFoldersByName(CONFIG.PROCESSED_SUBFOLDER_NAMES[k]);
      if (sub.hasNext()) {
        let files = sub.next().getFiles();
        let target = DriveApp.getFolderById(CONFIG.FOLDER_MAPPINGS[k]);
        while (files.hasNext()) { try { files.next().moveTo(target); } catch(e) {} }
      }
    });
  }
  
  ui.alert("✅ Rollback Complete!", "Your roster, notes, and screenshots have been successfully restored to their pre-import state.", ui.ButtonSet.OK);
}

function parseLeaderboardWithGemini(file, currentLabel) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("GEMINI_API_KEY missing in Script Properties.");
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const blob = file.getBlob();
  const payload = {
    contents: [{ parts: [{ text: `Analyze this Last War mobile game alliance leaderboard screenshot (${currentLabel}). Extract every commander/player name, their rank (if visible, e.g. R1, R2, R3, R4, R5), and their point score or metric value. Return ONLY a valid JSON array of objects with keys: 'rank', 'name', 'score'. No markdown.` }, { inlineData: { mimeType: blob.getContentType() || "image/jpeg", data: Utilities.base64Encode(blob.getBytes()) } }] }],
    generationConfig: { temperature: 0.1 }
  };
  
  let res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
  let json = JSON.parse(res.getContentText());
  
  if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
    throw new Error(json.error ? json.error.message : "Empty response from Gemini API");
  }
  
  let rawText = json.candidates[0].content.parts[0].text.trim();
  let match = rawText.match(/\[.*\]/s);
  let jsonString = match ? match[0] : rawText.replace(/^```json\s*|^```\s*|\s*```$/g, "");
  return JSON.parse(jsonString);
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