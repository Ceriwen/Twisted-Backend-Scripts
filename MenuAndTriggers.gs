/**
 * MenuAndTriggers.gs - Spreadsheet Custom Menu, Mobile Dropdown & Automation Triggers
 * TWST Roster Management System
 */

function runSingleImport(dayLabel) {
  updateQueueVisualStatus(`⚡ Importing ${dayLabel}...`, "#fff2cc");
  try {
    let res = processFolderDirectly([dayLabel]);
    if (!res || !res.timedOut) {
      updateQueueVisualStatus("✅ Import Complete - Ready", "#d9ead3");
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
      updateQueueVisualStatus("✅ Import Complete - Ready", "#d9ead3");
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
      updateQueueVisualStatus("✅ Import Complete - Ready", "#d9ead3");
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

  ScriptApp.newTrigger("sortAndProcessInboxFolder")
    .timeBased()
    .everyMinutes(15)
    .create();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");
  if (sheet) {
    sheet.getRange("C1").setValue("Watcher On").setBackground("#d9ead3").setFontColor("black").setFontWeight("bold");
    sheet.getRange("F1").setValue("Auto-watcher running every 15m").setBackground("#d9ead3");
  }

  ss.toast("Auto-import watcher enabled (checks Inbox every 15m).", "⚡ Watcher Active", 6);
}

function disableAutoImportWatcher() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sortAndProcessInboxFolder") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");
  if (sheet) {
    sheet.getRange("C1").setValue("Watcher Off").setBackground("#f4cccc").setFontColor("black").setFontWeight("bold");
    sheet.getRange("F1").setValue("").setBackground(null);
  }

  ss.toast("Auto-import watcher disabled.", "🛑 Automation Paused", 6);
}

/**
 * Composite Nightly Maintenance Trigger.
 * Runs rank recalculations, refreshes the Discord live card, and pushes full codebase sync to GitHub & Drive.
 */
function executeNightlyMaintenance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log");
  debugSheet.appendRow([new Date(), "🌙 Starting Nightly Maintenance..."]);

  try {
    if (typeof calculateWeeklyRankAdjustments === "function") {
      calculateWeeklyRankAdjustments();
    }
  } catch (e) {
    debugSheet.appendRow([new Date(), "⚠️ Nightly Rank Calc Error: " + e.message]);
  }

  try {
    if (typeof updateLiveDiscordDashboard === "function") {
      updateLiveDiscordDashboard();
    }
  } catch (e) {
    debugSheet.appendRow([new Date(), "⚠️ Nightly Discord Sync Error: " + e.message]);
  }

  try {
    if (typeof syncProjectToGitHub === "function") {
      syncProjectToGitHub();
    }
  } catch (e) {
    debugSheet.appendRow([new Date(), "⚠️ Nightly GitHub Sync Error: " + e.message]);
  }

  debugSheet.appendRow([new Date(), "✅ Nightly Maintenance Completed."]);
}

function installNightlyMaintenanceTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "executeNightlyMaintenance") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("executeNightlyMaintenance")
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Nightly Maintenance scheduled for 12:00 AM daily.",
    "🌙 Trigger Configured",
    5
  );
}

/**
 * Desktop Custom Menu
 */
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
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📥 Process Imports')
      .addItem('📥 Sort Inbox', 'sortAndProcessInboxFolder')
      .addSeparator()
      .addItem('🚀 Process Next Folder', 'processNextAvailableFolder')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('📂 Manual Single Imports')
        .addItem('📸 Import Strength Ranking - Power & Ranks', 'importPowerRankingWithRanks')
        .addItem('📸 Import Strength Ranking - Power Only', 'importPowerRankingOnly')
        .addItem('💰 Import Strength Ranking - Donations', 'importDonationsRanking')
        .addSeparator()
        .addItem('Monday Only', 'importMonOnly')
        .addItem('Tuesday Only', 'importTueOnly')
        .addItem('Wednesday Only', 'importWedOnly')
        .addItem('Thursday Only', 'importThuOnly')
        .addItem('Friday Only', 'importFriOnly')
        .addItem('Saturday Only', 'importSatOnly')
      )
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⏱️ Auto-Import Watcher')
      .addItem('⏱️ Toggle Auto-Import Watcher', 'toggleAutoImportWatcher')
    )
    .addSeparator()
    .addItem('🧹 Clear Table Highlights & Notes', 'clearAllHighlightsAndNotes')
    .addItem('🔮 Calculate Weekly Rank Adjustments', 'calculateWeeklyRankAdjustments')
    .addItem('📢 Send/Update Discord Live Card', 'updateLiveDiscordDashboard')
    .addItem('📅 Run Sunday End-of-Week Reset', 'executeSundayEowReset')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🐞 Debug & Cloud')
      .addItem('🔍 Backfill Strike Notes from Archives', 'backfillStrikeNotesFromArchives')
      .addSeparator()
      .addItem('🌙 Run Nightly Maintenance Now', 'executeNightlyMaintenance')
      .addItem('⚙️ Install Midnight Auto-Trigger', 'installNightlyMaintenanceTrigger')
      .addSeparator()
      .addItem('↩️ Rollback Last Import & Restore Screenshots', 'rollbackLastImport')
      .addItem('🔄 Restore Processed Screenshots', 'restoreProcessedScreenshots')
      .addItem('🗑️ Empty Processed Folders', 'emptyProcessedFolders')
      .addItem('🔄 Update Rank with New Rank (Active Sheet)', 'menuUpdateRankWithNewRank')
      .addItem('🐞 Clear Debug Log', 'clearDebugLogSheet')
      .addItem('📋 Show Debug Log Snippet', 'showDebugSnippet')
      .addSeparator()
      .addItem('🔄 Sync Code to GitHub & Drive', 'syncProjectToGitHub')
    )
    .addToUi();
}

/**
 * Mobile Dropdown Handler in cell D1 of "Data Import"
 */
function mobileMenuHandler(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Data Import") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1 && col === 4) {
    let val = String(e.range.getValue() || "").trim();
    if (!val || val === "⚙️ Select Action...") return;

    updateQueueVisualStatus(`⚡ Running: ${val}...`, "#fff2cc");

    try {
      if (val === "📥 Sort & Process Inbox") {
        sortAndProcessInboxFolder();
      } else if (val === "🚀 Process Next Folder") {
        processNextAvailableFolder();
      } else if (val === "🧹 Clear Highlights & Notes") {
        if (typeof clearAllHighlightsAndNotes === "function") clearAllHighlightsAndNotes();
      } else if (val === "🔮 Calculate Weekly Ranks") {
        if (typeof calculateWeeklyRankAdjustments === "function") calculateWeeklyRankAdjustments();
      } else if (val === "📢 Update Discord Live Card") {
        if (typeof updateLiveDiscordDashboard === "function") updateLiveDiscordDashboard();
      } else if (val === "📅 Run Sunday Reset") {
        if (typeof executeSundayEowReset === "function") executeSundayEowReset();
      } else if (val === "⏱️ Toggle Auto Watcher") {
        toggleAutoImportWatcher();
      } else if (val === "🌙 Run Nightly Maintenance") {
        executeNightlyMaintenance();
      } else if (val === "🔄 Sync Code to GitHub") {
        if (typeof syncProjectToGitHub === "function") syncProjectToGitHub();
      }
    } catch (err) {
      updateQueueVisualStatus(`❌ Error: ${err.message}`, "#f4cccc");
      let debugSheet = sheet.getParent().getSheetByName("Debug Log");
      if (debugSheet) {
        debugSheet.appendRow([new Date(), `Dropdown Error on [${val}]: ${err.stack || err.message}`]);
      }
    }

    Utilities.sleep(300);
    e.range.setValue("⚙️ Select Action...");
    SpreadsheetApp.flush();
  }
}