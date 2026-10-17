/**
 * GitHubSync.gs
 * Automated cloud sync pushing all project files to both GitHub AND a Google Drive mirror folder.
 */

const GITHUB_CONFIG = {
  BRANCH: "main",
  DRIVE_MIRROR_FOLDER_NAME: "TWST Codebase Mirror",
  MASTER_FILE_NAME: "TWST_Master_Codebase.txt",
  COMMIT_AUTHOR: {
    name: "TWST Cloud Sync",
    email: "automation@twst.internal"
  }
};

function syncProjectToGitHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("GITHUB_TOKEN");
  const owner = props.getProperty("GITHUB_OWNER");
  const repo = props.getProperty("GITHUB_REPO");

  if (typeof updateQueueVisualStatus === "function") {
    updateQueueVisualStatus("🔄 Syncing project to GitHub & Drive...", "#fff2cc");
  }
  try { ss.toast("Syncing project files to GitHub and Google Drive...", "🔄 Cloud Sync", 5); } catch (e) {}

  const scriptId = ScriptApp.getScriptId();
  const files = fetchProjectFiles(scriptId);
  Logger.log(`Found ${files.length} files in project.`);

  if (!files || files.length === 0) {
    const errMsg = "⚠️ Sync Failed: Could not read Apps Script project files (Check OAuth Scopes in appsscript.json).";
    Logger.log(errMsg);
    if (typeof updateQueueVisualStatus === "function") updateQueueVisualStatus(errMsg, "#f4cccc");
    try { ss.toast(errMsg, "❌ API Error", 10); } catch (e) {}
    return;
  }

  // 1. Sync to Google Drive Mirror Folder (Plain-text .txt for seamless Gemini access)
  let driveSyncedCount = 0;
  try {
    driveSyncedCount = syncFilesToDriveMirror(files);
  } catch (driveErr) {
    Logger.log("Drive Mirror Error: " + driveErr.message);
  }

  // 2. Sync to GitHub Repository
  if (!token || !owner || !repo) {
    const msg = `✅ Saved ${driveSyncedCount} files to Drive ("${GITHUB_CONFIG.DRIVE_MIRROR_FOLDER_NAME}"). ⚠️ Skipped GitHub (Missing Script Properties).`;
    Logger.log(msg);
    if (typeof updateQueueVisualStatus === "function") updateQueueVisualStatus(msg, "#fff2cc");
    try { ss.toast(msg, "⚠️ Partial Sync", 8); } catch (e) {}
    return;
  }

  const branch = GITHUB_CONFIG.BRANCH;
  let pushedCount = 0;
  let failedFiles = [];

  files.forEach(file => {
    let ext = ".gs";
    if (file.type === "HTML") ext = ".html";
    else if (file.type === "JSON") ext = ".json";

    const filename = file.name.endsWith(ext) ? file.name : (file.name + ext);
    const path = encodeURIComponent(filename);
    const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}?ref=${branch}`;

    // Get current SHA if file exists
    let sha = null;
    const getRes = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/vnd.github+json",
        "User-Agent": "Google-Apps-Script-Sync"
      },
      muteHttpExceptions: true
    });

    if (getRes.getResponseCode() === 200) {
      sha = JSON.parse(getRes.getContentText()).sha;
    }

    const payload = {
      message: `Cloud Sync: Update ${filename}`,
      content: Utilities.base64Encode(file.source, Utilities.Charset.UTF_8),
      branch: branch,
      committer: GITHUB_CONFIG.COMMIT_AUTHOR
    };
    if (sha) payload.sha = sha;

    const putRes = UrlFetchApp.fetch(`https://api.github.com/repos/${owner}/${repo}/contents/${path}`, {
      method: "put",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/vnd.github+json",
        "User-Agent": "Google-Apps-Script-Sync"
      },
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = putRes.getResponseCode();
    Logger.log(`Push ${filename}: HTTP ${code}`);
    if (code === 200 || code === 201) {
      pushedCount++;
    } else {
      failedFiles.push(`${filename} (HTTP ${code})`);
      Logger.log(`Failed ${filename}: ${putRes.getContentText()}`);
    }
  });

  let timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  if (failedFiles.length === 0) {
    const successMsg = `✅ Synced ${pushedCount} files to GitHub & Drive at ${timeStr}`;
    Logger.log(successMsg);
    if (typeof updateQueueVisualStatus === "function") updateQueueVisualStatus(successMsg, "#d9ead3");
    try { ss.toast(`Pushed ${pushedCount} files to GitHub & updated "${GITHUB_CONFIG.MASTER_FILE_NAME}" in Drive!`, "✅ Cloud Sync Complete", 6); } catch (e) {}
  } else {
    const warnMsg = `⚠️ GitHub: ${pushedCount} ok, Failed: ${failedFiles.join(", ")} | Drive: ${driveSyncedCount} ok`;
    Logger.log(warnMsg);
    if (typeof updateQueueVisualStatus === "function") updateQueueVisualStatus(warnMsg, "#f4cccc");
    try { ss.toast(warnMsg, "⚠️ Partial Sync", 10); } catch (e) {}
  }
}

/**
 * Saves each script as a readable .txt file AND builds a single combined Master file
 * in Google Drive so Gemini can read the entire live project from Drive anytime.
 */
function syncFilesToDriveMirror(files) {
  let folderIter = DriveApp.getFoldersByName(GITHUB_CONFIG.DRIVE_MIRROR_FOLDER_NAME);
  let folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(GITHUB_CONFIG.DRIVE_MIRROR_FOLDER_NAME);

  let masterContent = `================================================================================\n`;
  masterContent += `TWST ROSTER MANAGEMENT SYSTEM - MASTER CODEBASE SNAPSHOT\n`;
  masterContent += `Last Synced: ${new Date().toLocaleString()}\n`;
  masterContent += `================================================================================\n\n`;

  let count = 0;

  files.forEach(file => {
    let ext = ".gs";
    if (file.type === "HTML") ext = ".html";
    else if (file.type === "JSON") ext = ".json";

    const rawName = file.name.endsWith(ext) ? file.name : (file.name + ext);
    const txtFileName = rawName + ".txt"; // Plain text extension so Gemini never blocks it

    masterContent += `\n\n// =============================================================================\n`;
    masterContent += `// FILE: ${rawName}\n`;
    masterContent += `// =============================================================================\n\n`;
    masterContent += file.source + `\n`;

    upsertDriveTextFile(folder, txtFileName, file.source);
    count++;
  });

  // Save the combined single-file snapshot
  upsertDriveTextFile(folder, GITHUB_CONFIG.MASTER_FILE_NAME, masterContent);

  return count;
}

function upsertDriveTextFile(folder, filename, content) {
  let existing = folder.getFilesByName(filename);
  if (existing.hasNext()) {
    let file = existing.next();
    file.setContent(content);
    while (existing.hasNext()) {
      try { existing.next().setTrashed(true); } catch (e) {}
    }
  } else {
    folder.createFile(filename, content, MimeType.PLAIN_TEXT);
  }
}

function fetchProjectFiles(scriptId) {
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`API Error: ${res.getResponseCode()}\n${res.getContentText()}`);
    return [];
  }

  return JSON.parse(res.getContentText()).files || [];
}

/**
 * Run this function ONCE from the Apps Script editor to install 
 * a nightly automated trigger that syncs your project every night between midnight and 1 AM.
 */
function setupNightlyGitHubSyncTrigger() {
  // Clear any existing triggers for this function to prevent duplicates
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "syncProjectToGitHub") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create a new time-driven trigger to run every night
  ScriptApp.newTrigger("syncProjectToGitHub")
    .timeBased()
    .everyDays(1)
    .atHour(0) // Midnight
    .create();

  Logger.log("Nightly GitHub & Drive sync trigger successfully installed!");
}

/**
 * Run this if you ever want to disable the nightly auto-sync.
 */
function removeNightlyGitHubSyncTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "syncProjectToGitHub") {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log("Nightly sync trigger removed.");
}