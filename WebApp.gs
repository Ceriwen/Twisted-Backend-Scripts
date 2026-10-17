/**
 * WebApp.gs - Backend routing and dashboard request handlers
 * TWST Roster Management System
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('App');
  template.forceArchive = (e && e.parameter && e.parameter.archive === 'true') ? true : false;
  
  return template.evaluate()
    .setTitle('TWST Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getDashboardData(viewLastWeek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isArchive = (viewLastWeek === true || viewLastWeek === "true");
  const sheetName = isArchive ? "Last Week" : "Data Import";
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found. Please ensure it exists.");
  }

  const lastRow = Math.max(3, sheet.getLastRow());
  const dataRange = sheet.getRange(3, 1, lastRow - 2, 13);
  const data = dataRange.getValues();
  const notes = dataRange.getNotes(); 

  let members = [];
  let totalPower = 0;
  let totalWeekly = 0;
  let activeMembers = 0;

  const parseNum = (val) => {
    if (typeof val === "number") return val;
    return parseFloat(String(val).replace(/[^0-9.-]+/g, "")) || 0;
  };

  data.forEach((row, index) => {
    const name = String(row[1] || "").trim();
    if (!name || name.startsWith("---") || name.startsWith("⚠️") || name.toLowerCase() === "member") return;

    activeMembers++;
    
    // Column indices matching Sheet layout (Cols A-M):
    const rank = String(row[0] || "-").trim();
    const power = parseNum(row[2]);
    const mon = parseNum(row[3]);
    const tue = parseNum(row[4]);
    const wed = parseNum(row[5]);
    const thu = parseNum(row[6]);
    const fri = parseNum(row[7]);
    const sat = parseNum(row[8]);
    const weekly = parseNum(row[9]);
    const donations = parseNum(row[10]);
    const newRank = String(row[11] || "-").trim();

    totalPower += power;
    totalWeekly += weekly;

    let reason = String(row[12] || "").trim();
    const nameNote = String(notes[index][1] || "");

    if (nameNote.includes("Pending Removal") && !reason.includes("[Removal]")) {
      reason = (reason ? reason + " " : "") + "[Removal]";
    }
    if (nameNote.includes("Demoted Last Week") && !reason.includes("[Demotion]")) {
      reason = (reason ? reason + " " : "") + "[Demotion]";
    }

    members.push({
      id: "m" + index,
      rank: rank,
      name: name,
      power: power.toLocaleString(),
      mon: mon.toLocaleString(),
      tue: tue.toLocaleString(),
      wed: wed.toLocaleString(),
      thu: thu.toLocaleString(),
      fri: fri.toLocaleString(),
      sat: sat.toLocaleString(),
      weekly: weekly.toLocaleString(),
      donations: donations.toLocaleString(),
      newRank: newRank,
      reason: reason || "Active"
    });
  });

  const avgWeekly = activeMembers > 0 
    ? (totalWeekly / activeMembers / 1000000).toFixed(1) + "M" 
    : "0M";

  return {
    members: members,
    stats: {
      totalPower: totalPower >= 1000000000 
        ? (totalPower / 1000000000).toFixed(2) + "B" 
        : (totalPower / 1000000).toFixed(2) + "M",
      activeMembers: activeMembers,
      avgWeekly: avgWeekly
    }
  };
}

function uploadFilesToDrive(filesData, password) {
  try {
    if (password !== "Twisted4Life") {
      return { success: false, error: "Unauthorized: Invalid leadership password." };
    }

    let inboxFolderId = CONFIG.INBOX_FOLDER_ID;
    let targetFolder = DriveApp.getFolderById(inboxFolderId);

    let count = 0;
    filesData.forEach(file => {
      const decodedData = Utilities.base64Decode(file.data);
      const blob = Utilities.newBlob(decodedData, file.mimeType, file.name);
      targetFolder.createFile(blob);
      count++;
    });

    return { success: true, count: count };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function runAIBatchFromApp(password) {
  if (password !== "Twisted4Life") throw new Error("Unauthorized.");
  
  if (typeof sortAndProcessInboxFolder === "function") {
    return sortAndProcessInboxFolder();
  } else if (typeof processNextAvailableFolder === "function") {
    return processNextAvailableFolder();
  } else {
    throw new Error("Batch processing functions were not found.");
  }
}