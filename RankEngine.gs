/**
 * RankEngine.gs - Rank Evaluation & EOW Reset Routines
 * TWST Roster Management System
 */

function paintBatchFormatting(sheet, sheetData, nameColIdx, metricColIdx, rankColIdx, flaggedRows, matchedRows, updatedRows, isStrengthRanking, updatedProfileCells, verificationLogs, dynamicRankChanges, currentLabel) {
  let totalCurrentRows = sheetData.length;
  if (totalCurrentRows <= 1) return;
  let rosterRange = findRosterRange(sheet);
  let existingNameBg = sheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, totalCurrentRows - 1, 1).getBackgrounds();
  let existingNotes = sheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, totalCurrentRows - 1, 1).getNotes();
  
  let existingMetricBg = (metricColIdx > -1) ? sheet.getRange(rosterRange.startRow + 1, metricColIdx + 1, totalCurrentRows - 1, 1).getBackgrounds() : [];
  let existingMetricNotes = (metricColIdx > -1) ? sheet.getRange(rosterRange.startRow + 1, metricColIdx + 1, totalCurrentRows - 1, 1).getNotes() : [];
  
  let nameBackgrounds = [];
  let columnBackgrounds = [];
  let columnNotes = [];
  
  for (let k = 1; k < totalCurrentRows; k++) {
    let actualRow = k + 1;
    let priorNameColor = (existingNameBg[k - 1] && existingNameBg[k - 1][0]) ? existingNameBg[k - 1][0].toLowerCase() : "";
    let nameNote = (existingNotes[k - 1] && existingNotes[k - 1][0]) ? existingNotes[k - 1][0].toLowerCase() : "";
    let isExemptNoteFlag = isExemptNote(nameNote);
    let isAlreadyAmber = (priorNameColor === CONFIG.COLORS.ELIMINATION_ALERT.toLowerCase());
    
    let priorMetricColor = (existingMetricBg[k - 1] && existingMetricBg[k - 1][0]) ? existingMetricBg[k - 1][0].toLowerCase() : "";
    let priorMetricNote = (existingMetricNotes[k - 1] && existingMetricNotes[k - 1][0]) ? existingMetricNotes[k - 1][0] : "";
    let cellScore = (metricColIdx > -1) ? cleanNumericValues(sheetData[k][metricColIdx]) : 0;

    if (!sheetData[k][nameColIdx] || sheetData[k][nameColIdx].toString().trim() === "") {
      nameBackgrounds.push([priorNameColor || ""]);
      if (metricColIdx > -1) {
        columnBackgrounds.push([priorMetricColor || ""]);
        columnNotes.push([priorMetricNote || ""]);
      }
      continue;
    }
    
    if (isExemptNoteFlag) {
      nameBackgrounds.push([priorNameColor || CONFIG.COLORS.MATCHED_MEMBER]);
    } else if (flaggedRows.includes(actualRow) || isAlreadyAmber) {
      nameBackgrounds.push([CONFIG.COLORS.ELIMINATION_ALERT]);
    } else if (matchedRows.has(actualRow)) {
      nameBackgrounds.push([CONFIG.COLORS.MATCHED_MEMBER]);
    } else {
      nameBackgrounds.push([priorNameColor || ""]);
    }
    
    if (metricColIdx > -1) {
      let isCumulativeFolder = (currentLabel === "Strength Ranking - Power" || currentLabel === "Strength Ranking - Donations");
      
      if (matchedRows.has(actualRow) && !isStrengthRanking) {
        columnBackgrounds.push([CONFIG.COLORS.DAILY_SUCCESS]);
        columnNotes.push([verificationLogs[actualRow] || priorMetricNote]);
      } else if (priorMetricColor === CONFIG.COLORS.DAILY_SUCCESS.toLowerCase() || cellScore > 0) {
        // PRESERVE existing successful match or score from earlier batch run!
        columnBackgrounds.push([priorMetricColor || CONFIG.COLORS.DAILY_SUCCESS]);
        columnNotes.push([priorMetricNote]);
      } else if (!isStrengthRanking && !isCumulativeFolder && currentLabel) {
        columnBackgrounds.push([CONFIG.COLORS.MISSING_DATA]);
        columnNotes.push([verificationLogs[actualRow] || `⚠️ MEMBER NOT DETECTED (${currentLabel}):
Score missing.`]);
      } else {
        columnBackgrounds.push([priorMetricColor || ""]);
        columnNotes.push([priorMetricNote || ""]);
      }
    }
  }
  
  sheet.getRange(rosterRange.startRow + 1, nameColIdx + 1, totalCurrentRows - 1, 1).setBackgrounds(nameBackgrounds);
  let isCumulativeFolderCheck = (currentLabel === "Strength Ranking - Power" || currentLabel === "Strength Ranking - Donations");
  if (metricColIdx > -1 && !isStrengthRanking && !isCumulativeFolderCheck) {
    sheet.getRange(rosterRange.startRow + 1, metricColIdx + 1, totalCurrentRows - 1, 1).setBackgrounds(columnBackgrounds).setNotes(columnNotes);
  }
}

function calculateWeeklyRankAdjustments() {
  runRankCalculationEngine(["mon", "tue", "wed", "thu", "fri", "sat", "donations"], "Post-Import Auto-Calculation");
}

function runRankCalculationEngine(targetDayLabels, contextLabelMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let rosterRange = findRosterRange(sheet);
  if (!rosterRange) return;
  
  let sheetRange = sheet.getRange(rosterRange.startRow, 1, rosterRange.rowCount, rosterRange.colCount);
  let sheetData = sheetRange.getValues();
  let sheetNotes = sheet.getRange(rosterRange.startRow + 1, 2, rosterRange.rowCount - 1, 1).getNotes();
  const headers = sheetData[0].map(h => h.toString().trim().toLowerCase());
  
  const rankColIdx = headers.indexOf("rank");
  const nameColIdx = headers.findIndex(h => h.includes("member"));
  const donationsColIdx = headers.findIndex(h => h.includes("donation"));
  const newRankColIdx = headers.findIndex(h => h.includes("new rank"));
  const reasonColIdx = headers.findIndex(h => h.includes("reason") || h.includes("notes"));
  const weeklyColIdx = headers.findIndex(h => h.includes("weekly"));
  const mondayColIdx = headers.findIndex(h => h.includes("mon"));

  let missing = [];
  if (rankColIdx === -1) missing.push("Rank");
  if (nameColIdx === -1) missing.push("Member");
  if (donationsColIdx === -1) missing.push("Donations");
  if (newRankColIdx === -1) missing.push("New Rank");

  if (missing.length > 0) {
    ss.toast(`Calculation aborted. Missing columns: ${missing.join(", ")}. Please check headers!`, "⚠️ Header Not Found", 10);
    return;
  }
  
  const dayIndices = targetDayLabels.map(day => headers.findIndex(h => h.includes(day.replace("donations", "donation")))).filter(idx => idx > -1);
  let processedDayIndices = [];
  let totalValidMembers = getTotalRosterCount(sheet);
  
  dayIndices.forEach(idx => {
    let populatedCount = 0;
    for (let r = 1; r < sheetData.length; r++) {
      let val = sheetData[r][idx];
      let nameVal = sheetData[r][nameColIdx] ? sheetData[r][nameColIdx].toString().trim() : "";
      if (nameVal !== "" && !nameVal.startsWith("⚠️") && val !== null && val !== undefined && val !== "") populatedCount++;
    }
    if (totalValidMembers > 0 && (populatedCount / totalValidMembers) >= 0.25) processedDayIndices.push(idx);
  });

  let isWeekFinished = (processedDayIndices.length >= 7);
  let isMondayProcessed = processedDayIndices.includes(mondayColIdx);
  
  let outputRangeValues = [];
  let weeklyScoresToWrite = [];
  
  for (let i = 1; i < sheetData.length; i++) {
    let row = sheetData[i];
    let name = row[nameColIdx] ? row[nameColIdx].toString().trim() : "";
    let currentRank = row[rankColIdx] ? row[rankColIdx].toString().trim().toUpperCase() : "";
    let nameNote = (sheetNotes[i - 1] && sheetNotes[i - 1][0]) ? sheetNotes[i - 1][0].toLowerCase() : "";
    
    if (!name || name.startsWith("⚠️") || name.startsWith("MISSING")) {
      outputRangeValues.push(["", ""]);
      weeklyScoresToWrite.push([""]);
      continue;
    }
    
    let isGracePeriod = false;
    if (isMondayProcessed && mondayColIdx > -1) {
      let monVal = row[mondayColIdx];
      if (monVal === "" || monVal === null || monVal === undefined) {
        isGracePeriod = true;
      }
    }
    
    let hasAnyActivity = dayIndices.some(idx => row[idx] !== null && row[idx] !== undefined && row[idx] !== "");
    
    let calculatedWeeklySum = 0;
    let vsDayIndices = dayIndices.filter(idx => !headers[idx].includes("donation"));
    
    vsDayIndices.forEach(idx => {
      let cellRawValue = row[idx];
      if (cellRawValue !== null && cellRawValue !== undefined && cellRawValue !== "") {
        calculatedWeeklySum += cleanNumericValues(cellRawValue);
      }
    });
    
    weeklyScoresToWrite.push([calculatedWeeklySum > 0 ? calculatedWeeklySum : ""]);

    if (!hasAnyActivity && (currentRank === "R3" || currentRank === "")) {
      outputRangeValues.push(["R3", `[New Member] ${CONFIG.TARGETS.MSG_NEW_MEMBER}`]);
      continue;
    }
    
    let missedVsDaysCount = 0, perfectVsDaysCount = 0;
    vsDayIndices.forEach(idx => {
      let cellRawValue = row[idx];
      if (cellRawValue !== null && cellRawValue !== undefined && cellRawValue !== "") {
        if (cleanNumericValues(cellRawValue) < CONFIG.TARGETS.DEMOTION_MIN_DAILY_VS_POINTS) missedVsDaysCount++;
        else perfectVsDaysCount++;
      }
    });
    
    let missedDonationsGoal = false;
    let totalDonations = donationsColIdx > -1 ? cleanNumericValues(row[donationsColIdx]) : 0;
    if (isWeekFinished && donationsColIdx > -1) {
      missedDonationsGoal = totalDonations < (CONFIG.TARGETS.DEMOTION_MIN_DAILY_DONATIONS * 6);
    }
    
    let allVsDaysHitElite = true;
    vsDayIndices.forEach(idx => {
      if (cleanNumericValues(row[idx]) < CONFIG.TARGETS.CRUSHED_IT_MIN_VS_SCORE) allVsDaysHitElite = false;
    });

    let isEliteWeek = isWeekFinished && allVsDaysHitElite && (perfectVsDaysCount === vsDayIndices.length) && (!missedDonationsGoal) && ((totalDonations / 6) >= CONFIG.TARGETS.CRUSHED_IT_ELITE_DONATIONS);
    let isPerfectWeek = isWeekFinished && (perfectVsDaysCount === vsDayIndices.length) && (!missedDonationsGoal);
    
    let failureDetails = [];
    if (missedVsDaysCount > 0) failureDetails.push(`${missedVsDaysCount} VS day(s)`);
    if (missedDonationsGoal) failureDetails.push(`Daily Donations Goal`);
    
    let totalFailureCount = missedVsDaysCount + (missedDonationsGoal ? 1 : 0);
    let triggerDemotion = totalFailureCount >= CONFIG.TARGETS.REQUIRED_MISSED_COUNT_FOR_DEMOTION;
    
    let calculatedRank = currentRank;
    let rankTags = [];
    let reasonDetails = "";
    
    let isProtectedManagement = CONFIG.TARGETS.PROTECT_RANKS.includes(currentRank);
    let isExempt = isExemptNote(nameNote) || isGracePeriod;
    
    if (isProtectedManagement) {
      calculatedRank = currentRank;
    }
    
    if (isGracePeriod) {
      rankTags.push("[New Member]");
      calculatedRank = currentRank;
      reasonDetails = "Joined mid-week (Monday score missing)";
    } else if (isExempt && !isGracePeriod) {
      rankTags.push("[Exempt]");
      calculatedRank = currentRank;
    }
    
    if (triggerDemotion && !isProtectedManagement && !isExempt) {
      if (currentRank === "R1") {
        calculatedRank = "R1";
        rankTags.push(CONFIG.TARGETS.TAG_IMMEDIATE_REMOVAL);
      } else {
        calculatedRank = (currentRank === "R3" ? "R2" : "R1");
        rankTags.push(CONFIG.TARGETS.TAG_DEMOTION);
      }
      reasonDetails = `Missed: ${failureDetails.join(" and ")}`;
      
    } else if (!isWeekFinished && !isGracePeriod) {
      let baseTag = rankTags.length > 0 ? rankTags.join(" ") : CONFIG.TARGETS.TAG_MAINTAINED;
      outputRangeValues.push([calculatedRank, `${baseTag} ${CONFIG.TARGETS.MSG_MID_WEEK_STATUS}`]);
      continue;
      
    } else if (!isGracePeriod) {
      let earnedElite = false, earnedPerfect = false;
      if (isEliteWeek) {
        earnedElite = true;
        if (!isProtectedManagement && !isExempt) calculatedRank = "R3";
        rankTags.push(CONFIG.TARGETS.TAG_CRUSHED_IT);
        reasonDetails = CONFIG.TARGETS.MSG_CRUSHED_IT;
      } else if (isPerfectWeek) {
        earnedPerfect = true;
        if (!isProtectedManagement && !isExempt) calculatedRank = "R3";
        rankTags.push(CONFIG.TARGETS.TAG_PERFECT);
        reasonDetails = `100% VS & Donations Met`;
      }
      
      let eligibleForPromotion = (currentRank === "R1" || currentRank === "R2");
      if (CONFIG.TARGETS.PROMOTION_RESTORE_TO_R3 && eligibleForPromotion && !isProtectedManagement && !isExempt) {
        calculatedRank = "R3";
        rankTags.push(CONFIG.TARGETS.TAG_PROMOTION);
        reasonDetails = earnedElite ? `${CONFIG.TARGETS.MSG_CRUSHED_IT} Restored to R3.` : earnedPerfect ? `100% VS & Donations Met. Restored to R3.` : `Goals met (VS & Donations). Restored to R3.`;
      }
      
      if (!earnedElite && !earnedPerfect && !rankTags.includes(CONFIG.TARGETS.TAG_PROMOTION)) {
        rankTags.push(CONFIG.TARGETS.TAG_MAINTAINED);
        reasonDetails = reasonDetails || `Safe, Allowed ${totalFailureCount} missed target(s)`;
      }
    }
    
    let uniqueTags = [...new Set(rankTags)].join(" ");
    outputRangeValues.push([calculatedRank, reasonDetails ? `${uniqueTags} ${reasonDetails}` : uniqueTags]);
  }
  
  if (reasonColIdx > -1 && reasonColIdx !== newRankColIdx + 1) {
    sheet.getRange(rosterRange.startRow + 1, newRankColIdx + 1, outputRangeValues.length, 1).setValues(outputRangeValues.map(r => [r[0]]));
    sheet.getRange(rosterRange.startRow + 1, reasonColIdx + 1, outputRangeValues.length, 1).setValues(outputRangeValues.map(r => [r[1]]));
  } else {
    sheet.getRange(rosterRange.startRow + 1, newRankColIdx + 1, outputRangeValues.length, 2).setValues(outputRangeValues);
  }
  
  if (weeklyColIdx > -1) {
    sheet.getRange(rosterRange.startRow + 1, weeklyColIdx + 1, weeklyScoresToWrite.length, 1).setValues(weeklyScoresToWrite);
  }
  
    // Ensure extra helper columns are purged and row verification dropdowns applied automatically
  try {
    if (typeof deleteExtraColumns === "function") deleteExtraColumns(sheet);
    if (typeof applyVerificationDropdowns === "function") applyVerificationDropdowns(sheet);
  } catch (e) {}

  // Auto-trigger Discord Live Card update
  try {
    if (typeof updateLiveDiscordDashboard === "function") {
      updateLiveDiscordDashboard();
    }
  } catch (e) {
    console.error("Discord Auto-Update error: " + e.message);
  }
  
  ss.toast(`Rank calculations updated successfully (${contextLabelMode}).`, "🔮 Ranks Updated", 4);
}

function executeSundayEowReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
  let ui = SpreadsheetApp.getUi();
  
  if (Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "EEEE") !== "Sunday" && ui.alert("⚠️ Boundary Warning", "Today is not Sunday. Proceed anyway?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (ui.alert("Confirm Sunday Reset", "Overwrite 'Last Week' tab, backup workbook, apply New Ranks, append Penalty Notes, and clear points?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  let today = new Date();
  let dayOfWeek = today.getDay();
  let targetSunday = new Date(today);
  if (dayOfWeek !== 0) {
    targetSunday.setDate(today.getDate() - dayOfWeek);
  }
  let dateStr = Utilities.formatDate(targetSunday, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  let rosterRange = findRosterRange(sheet);
  let values = sheet.getRange(rosterRange.startRow, 1, rosterRange.rowCount, rosterRange.colCount).getValues();
  let headers = values[0].map(h => h.toString().trim().toLowerCase());
  
  let nameIdx = headers.findIndex(h => h.includes("member"));
  let rankIdx = headers.indexOf("rank");
  let newRankIdx = headers.findIndex(h => h.includes("new rank"));
  let reasonIdx = headers.findIndex(h => h.includes("reason") || h.includes("notes"));
  
  let nameNotes = sheet.getRange(rosterRange.startRow, nameIdx + 1, rosterRange.rowCount, 1).getNotes();
  let preservedData = [];
  
  for (let r = 1; r < values.length; r++) {
    let name = values[r][nameIdx];
    if (!name || name.toString().trim() === "") continue;
    
    let currentRank = rankIdx > -1 ? values[r][rankIdx] : "";
    let newRank = newRankIdx > -1 ? values[r][newRankIdx] : "";
    let reason = reasonIdx > -1 ? values[r][reasonIdx].toString() : "";
    let currentNote = nameNotes[r][0];
    
    let finalRank = (newRank && newRank.toString().trim() !== "") ? newRank : currentRank;
    
    let keepColor = false;
    let finalNote = currentNote;
    
    if (reason.includes("[Removal]") || reason.includes("[Demotion]")) {
      let strikeText = reason.includes("[Removal]") ? "⚠️ Pending Removal" : "⚠️ Demoted Last Week";
      if (!finalNote.includes(strikeText)) {
        finalNote = finalNote ? finalNote + "\n" + strikeText : strikeText;
      }
      keepColor = true;
    }
    
    preservedData.push({
      row: r + rosterRange.startRow,
      rank: finalRank,
      note: finalNote,
      keepColor: keepColor
    });
  }

  // 1. Apply New Rank to the Rank column FIRST so the archives reflect updated ranks
  if (rankIdx > -1) {
    let startRow = rosterRange.startRow + 1;
    let numRows = rosterRange.rowCount - 1;
    let rankRange = sheet.getRange(startRow, rankIdx + 1, numRows, 1);
    let rankValues = rankRange.getValues();
    preservedData.forEach(p => {
      let rOffset = p.row - startRow;
      if (rOffset >= 0 && rOffset < rankValues.length) rankValues[rOffset][0] = p.rank;
    });
    rankRange.setValues(rankValues);
    SpreadsheetApp.flush();
  }

  // 2. Freeze updated state into 'Last Week' tab
  let lastWeekSheet = ss.getSheetByName("Last Week");
  if (!lastWeekSheet) lastWeekSheet = ss.insertSheet("Last Week");
  
  lastWeekSheet.clear();
  try {
    let sourceRange = sheet.getDataRange();
    let vals = sourceRange.getDisplayValues();
    let bgs = sourceRange.getBackgrounds();
    let fcs = sourceRange.getFontColors();
    let fws = sourceRange.getFontWeights();
    
    let targetRange = lastWeekSheet.getRange(1, 1, vals.length, vals[0].length);
    targetRange.setValues(vals);
    targetRange.setBackgrounds(bgs);
    targetRange.setFontColors(fcs);
    targetRange.setFontWeights(fws);
  } catch (e) {
    ss.toast("Minor error freezing formatting, but data was backed up to Drive.", "Warning", 4);
  }

  // 3. Create Drive archive backup workbook
  let backup = DriveApp.getFileById(ss.getId()).makeCopy(`TWST Roster Archive - ${dateStr}`);
  let archiveTab = ss.getSheetByName("Archive Log") || ss.insertSheet("Archive Log");
  if (archiveTab.getLastRow() === 0) archiveTab.getRange("A1:C1").setValues([["Archive Date", "Workbook Name", "Direct Link"]]).setFontWeight("bold");
  archiveTab.appendRow([dateStr, backup.getName(), `=HYPERLINK("${backup.getUrl()}", "Open Backup File")`]);

  // 4. Clear scores and requirements for the fresh week
  let colsToClear = ["mon", "tue", "wed", "thu", "fri", "sat", "donation", "weekly", "new rank", "reason", "notes"];
  let colIndicesToClear = [];
  headers.forEach((h, idx) => {
    if (colsToClear.some(c => h.includes(c))) colIndicesToClear.push(idx + 1);
  });
  
  colIndicesToClear.forEach(col => {
    sheet.getRange(rosterRange.startRow + 1, col, rosterRange.rowCount - 1, 1).clearContent().setBackground(null).clearNote();
  });
  
  sheet.getRange(rosterRange.startRow + 1, nameIdx + 1, rosterRange.rowCount - 1, 1).setBackground(null).clearNote();
  
  // 5. Reapply penalty notes and amber fills on Data Import
  if (preservedData.length > 0) {
    let startRow = rosterRange.startRow + 1;
    let numRows = rosterRange.rowCount - 1;
    let nameRange = sheet.getRange(startRow, nameIdx + 1, numRows, 1);
    let existingNotes = nameRange.getNotes();
    let existingBgs = nameRange.getBackgrounds();
    preservedData.forEach(p => {
      let rOffset = p.row - startRow;
      if (rOffset >= 0 && rOffset < numRows) {
        if (p.note) existingNotes[rOffset][0] = p.note;
        if (p.keepColor) existingBgs[rOffset][0] = CONFIG.COLORS.ELIMINATION_ALERT;
      }
    });
    nameRange.setNotes(existingNotes).setBackgrounds(existingBgs);
  }
  
  emptyProcessedFolders(true);
  SpreadsheetApp.flush();

  // Send Discord EOW Broadcast
  try {
    if (typeof sendSundayEOWReport === "function") {
      sendSundayEOWReport();
    }
  } catch (e) {
    console.error("Discord webhook failed: " + e.message);
  }

  ui.alert("Reset Complete", `Data Import safely frozen to 'Last Week'.\nNew Ranks applied, penalties moved to Name hover notes, and points cleared.`, ui.ButtonSet.OK);
}
