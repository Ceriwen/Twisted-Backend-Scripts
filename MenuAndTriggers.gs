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
 * Runs rank recalculations, refreshes Discord live card, and syncs code to GitHub & Drive.
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
    .addItem('⏱️ Toggle Auto Watcher', 'toggleAutoImportWatcher')
    .addItem('📋 Open Audit Sidebar', 'openAuditSidebarFromMenu')
    .addItem('📥 Sort & Process Folders', 'sortAndProcessInboxFolder')
    .addItem('🔄 Sync Ranks, Discord & Code', 'executeNightlyMaintenance')
    .addItem('🧹 Clear Highlights & Notes', 'clearAllHighlightsAndNotes')
    .addItem('📅 Run Sunday Reset', 'executeSundayEowReset')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🐞 Debugging Imports')
      .addItem('🚀 Process Next Folder', 'processNextAvailableFolder')
      .addSeparator()
      .addItem('📥 Strength: Power & Ranks', 'importPowerRankingWithRanks')
      .addItem('📥 Strength: Power Only', 'importPowerRankingOnly')
      .addItem('📥 Strength: Donations', 'importDonationsRanking')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('📅 Daily VS Days')
        .addItem('📥 Monday VS', 'importMonOnly')
        .addItem('📥 Tuesday VS', 'importTueOnly')
        .addItem('📥 Wednesday VS', 'importWedOnly')
        .addItem('📥 Thursday VS', 'importThuOnly')
        .addItem('📥 Friday VS', 'importFriOnly')
        .addItem('📥 Saturday VS', 'importSatOnly')
      )
      .addSubMenu(SpreadsheetApp.getUi().createMenu('↩️ Rollback & Restore')
        .addItem('↩️ Rollback Last Import', 'rollbackLastImport')
        .addItem('↩️ Restore Screenshots', 'restoreProcessedScreenshots')
        .addItem('🗑️ Empty Processed Folders', 'emptyProcessedFolders')
      )
    )

    // ==========================================
    // SUBMENU 2: System, Logs & Cloud Tools
    // ==========================================
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠️ System Maintenance')
      .addItem('🧮 Calculate Weekly Rank Adjustments', 'calculateWeeklyRankAdjustments')
      .addItem('🔄 Update Rank with New Rank', 'menuUpdateRankWithNewRank')
      .addItem('📋 Show Debug Snippet', 'showDebugSnippet')
      .addSeparator()
      .addItem('☁️ Send/Update Discord Live Card', 'updateLiveDiscordDashboard')
      .addItem('☁️ Sync Code to GitHub & Drive', 'syncProjectToGitHub')
      .addSeparator()
      .addItem('🧹 Clear Audit History', 'clearImportAuditLog')
      .addItem('🧹 Clear Debug Log', 'clearDebugLogSheet')
      
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

  // Cell D1 is Row 1, Column 4
  if (row === 1 && col === 4) {
    let val = String(e.range.getValue() || "").trim();
    if (!val || val === "⚙️ Select Action...") return;

    updateQueueVisualStatus(`⚡ Running: ${val}...`, "#fff2cc");

    try {
      if (val === "⏱️ Toggle Auto Watcher") {
        toggleAutoImportWatcher();
      } else if (val === "📋 Open Audit Sidebar") {
        openAuditSidebarFromMenu();
      } else if (val === "📥 Sort & Process Folders") {
        sortAndProcessInboxFolder();
      } else if (val === "🔄 Sync Ranks, Discord & Code") {
        executeNightlyMaintenance();
      } else if (val === "🧹 Clear Highlights & Notes") {
        if (typeof clearAllHighlightsAndNotes === "function") clearAllHighlightsAndNotes();
      } else if (val === "📅 Run Sunday Reset") {
        if (typeof executeSundayEowReset === "function") executeSundayEowReset();
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