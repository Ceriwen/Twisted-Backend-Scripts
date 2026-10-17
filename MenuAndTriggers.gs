/**
 * MenuAndTriggers.gs - Spreadsheet Custom Menu & Auto-Import Trigger Handlers
 * TWST Roster Management System
 */

function runSingleImport(dayLabel) {
  updateQueueVisualStatus(`⚡ Importing ${dayLabel}...`, "#fff2cc");
  try {
    let res = processFolderDirectly([dayLabel]);
    if (!res || !res.timedOut) {
      // Rank calculation removed from post-import auto-trigger
      updateQueueVisualStatus("✅ Import Complete - Ready for Verification", "#d9ead3");
    }
  } catch (err) {
    updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
  }
}

function importMonOnly() { runSingleImport("Mon"); }
function importTueOnly() { runSingleImport("Tue"); }
function importWedOnly() { runSingleImport("Wed"); }
function importThuOnly() { runSingleImport("Thu"); }
function importFriOnly() { runSingleImport("Fri"); }
function importSatOnly() { runSingleImport("Sat"); }

function importPowerRankingOnly() { 
  updateQueueVisualStatus("⚡ Importing Power Ranking Only...", "#fff2cc");
  try {
    let res = processFolderDirectly(["Strength Ranking - Power"], false, true); 
    if (!res || !res.timedOut) {
      // Rank calculation removed from post-import auto-trigger
      updateQueueVisualStatus("✅ Import Complete - Ready for Verification", "#d9ead3");
    }
  } catch (err) {
    updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
  }
}

function importPowerRankingWithRanks() { 
  updateQueueVisualStatus("⚡ Importing Power Ranking & Ranks...", "#fff2cc");
  try {
    let res = processFolderDirectly(["Strength Ranking - Power"], false, false); 
    if (!res || !res.timedOut) {
      // Rank calculation removed from post-import auto-trigger
      updateQueueVisualStatus("✅ Import Complete - Ready for Verification", "#d9ead3");
    }
  } catch (err) {
    updateQueueVisualStatus("⚠️ Interrupted/Timed Out. Run again to resume.", "#f4cccc");
  }
}

function importDonationsRanking() { 
  runSingleImport("Strength Ranking - Donations");
}

function toggleAutoImportWatcher() {
  let triggers = ScriptApp.getProjectTriggers();
  let isWatcherActive = triggers.some(t => t.getHandlerFunction() === "sortAndProcessInboxFolder");
  
  if (isWatcherActive) {
    disableAutoImportWatcher();
  } else {
    enableAutoImportWatcher();
  }
}

function enableAutoImportWatcher() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sortAndProcessInboxFolder") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("sortAndProcessInboxFolder").timeBased().everyMinutes(15).create();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");
  if (sheet) {
    sheet.getRange("C1").setValue("Watcher On").setBackground("#d9ead3").setFontColor("black").setFontWeight("bold");
    sheet.getRange("F1").setValue("⏱️ Watcher Started. Next check in 15m...").setBackground("#e8f0fe").setFontWeight("bold");
  }
  
  ss.toast("Auto-import watcher enabled! Running every 15 minutes.", "⏱️ Automation Active", 6);
}

function disableAutoImportWatcher() {
  let triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sortAndProcessInboxFolder") {
      ScriptApp.deleteTrigger(triggers[i]);
      removedCount++;
    }
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");
  if (sheet) {
    sheet.getRange("C1").setValue("Watcher Off").setBackground("#f4cccc").setFontColor("black").setFontWeight("bold");
    sheet.getRange("F1").setValue("").setBackground(null);
  }
  
  ss.toast(`Auto-import watcher disabled.`, "🛑 Automation Paused", 6);
}

function onOpen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data Import");
    if (sheet) {
      let triggers = ScriptApp.getProjectTriggers();
      let isWatcherActive = triggers.some(t => t.getHandlerFunction() === "sortAndProcessInboxFolder");
      if (isWatcherActive) {
        sheet.getRange("C1").setValue("Watcher On").setBackground("#d9ead3").setFontColor("black").setFontWeight("bold");
      } else {
        sheet.getRange("C1").setValue("Watcher Off").setBackground("#f4cccc").setFontColor("black").setFontWeight("bold");
      }
    }
  } catch (e) {}

  SpreadsheetApp.getUi().createMenu('🌪️ TWST Tools')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('📥Process Imports')
      .addItem('📥 Sort Inbox', 'sortAndProcessInboxFolder')
      .addSeparator()
      .addItem('🚀 Process Next Folder', 'processNextAvailableFolder')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('📂 Manual Single Imports')
        .addItem('📸 Import Strength Ranking - Power & Ranks', 'importPowerRankingWithRanks')
        .addItem('💰 Import Strength Ranking - Donations', 'importDonationsRanking')
        .addSeparator()
        .addItem('Monday Only', 'importMonOnly').addItem('Tuesday Only', 'importTueOnly').addItem('Wednesday Only', 'importWedOnly')
        .addItem('Thursday Only', 'importThuOnly').addItem('Friday Only', 'importFriOnly').addItem('Saturday Only', 'importSatOnly')
      )
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⏱️ Auto-Import Watcher')
      .addItem('⏱️ Toggle Auto-Import Watcher', 'toggleAutoImportWatcher')
    )
    .addSeparator()
    .addItem('🧹 Verify & Restore Roster Order', 'menuVerifyAndRestore')
    .addItem('🧹 Clear Table Highlights & Notes', 'clearAllHighlightsAndNotes')
    .addSeparator()
    .addItem('🔮 Calculate Weekly Rank Adjustments', 'calculateWeeklyRankAdjustments')
    .addItem('📢 Send/Update Discord Live Card', 'updateLiveDiscordDashboard')
    .addItem('📅 Run Sunday End-of-Week Reset', 'executeSundayEowReset')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🐞 Debug Menu')
      .addItem('↩️ Rollback Last Import & Restore Screenshots', 'rollbackLastImport')
      .addItem('🔄 Restore Processed Screenshots', 'restoreProcessedScreenshots')
      .addItem('🗑️ Empty Processed Folders', 'emptyProcessedFolders')
      .addItem('🔄 Update Rank with New Rank (Active Sheet)', 'menuUpdateRankWithNewRank')
      .addItem('🐞 Clear Debug Log', 'clearDebugLogSheet')
      .addItem('📋 Show Debug Log Snippet', 'showDebugSnippet')
      .addItem('📸 Import Strength Ranking - Power Only', 'importPowerRankingOnly')
      .addSeparator()
      .addItem('🔄 Sync Code to GitHub', 'syncProjectToGitHub')
    )
  .addToUi();
}