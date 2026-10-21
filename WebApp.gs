/**
 * WebApp.gs - Google Apps Script Backend for TWST Command Center Web App
 * TWST Roster Management System
 */

function doGet(e) {
  let forceArchive = (e && e.parameter && e.parameter.view === 'archive') ? true : false;
  let template = HtmlService.createTemplateFromFile("App");
  template.forceArchive = forceArchive;
  
  return template.evaluate()
    .setTitle("TWST Ops Command Center")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Universal Archive Analytics Compiler.
 * Scans canonical 'Data Import' across all past archive workbooks to calculate
 * power growth, average VS, demotions, reliability rate, and weeks tracked.
 */
function compileArchiveAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let archiveSheet = ss.getSheetByName("Archive Log") || ss.getSheets().find(s => s.getName().trim().toLowerCase() === "archive log");
  if (!archiveSheet) {
    console.error("❌ 'Archive Log' sheet tab not found.");
    return;
  }

  const lastRow = archiveSheet.getLastRow();
  if (lastRow < 2) {
    console.warn("⚠️ 'Archive Log' has no archive entries below the header.");
    return;
  }

  const numRows = lastRow - 1;
  const range = archiveSheet.getRange(2, 1, numRows, 3);
  const archiveVals = range.getValues();
  const archiveFormulas = range.getFormulas();
  const archiveRichTexts = range.getRichTextValues();

  let historyMap = {};
  let processedFiles = 0;

  console.log(`🔍 Found ${numRows} archive entries in "Archive Log". Beginning deep scan...`);

  const REMOVAL_REGEX = /[immediate removal]|[removal]|\bpending removal\b/i;
  const DEMOTION_REGEX = /[demotion]|\bdemoted\b|missed:\s*\d+/i;
  const EXEMPT_REGEX = /[exempt]|\bexempt\b|grace period|\(rn\)|loa\b|vacation|absence|sick|personal/i;

  for (let a = 0; a < numRows; a++) {
    let dateStr = archiveVals[a][0] ? Utilities.formatDate(new Date(archiveVals[a][0]), "GMT-4", "yyyy-MM-dd") : `Row ${a + 2}`;
    let workbookName = archiveVals[a][1] ? archiveVals[a][1].toString().trim() : "";
    let rawFormula = archiveFormulas[a][2] || "";
    let rawVal = archiveVals[a][2] ? archiveVals[a][2].toString().trim() : "";

    let fileId = "";
    let rt = archiveRichTexts[a][2];
    if (rt && rt.getLinkUrl()) {
      let m = rt.getLinkUrl().match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) fileId = m[1];
    }
    if (!fileId && rawFormula) {
      let m = rawFormula.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) fileId = m[1];
    }
    if (!fileId && rawVal) {
      let m = rawVal.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) fileId = m[1];
    }

    let file = null;
    if (fileId) {
      try { file = SpreadsheetApp.openById(fileId); } catch (e) {}
    }
    if (!file && workbookName) {
      try {
        let matches = DriveApp.getFilesByName(workbookName);
        if (matches.hasNext()) file = SpreadsheetApp.open(matches.next());
      } catch (e) {}
    }

    if (!file) {
      console.warn(`⚠️ Row ${a + 2} (${dateStr}): Unable to open workbook "${workbookName}"`);
      continue;
    }

    let snapSheet = file.getSheetByName("Data Import") || file.getSheetByName("Last Week") || file.getSheets()[0];
    let allData = snapSheet.getDataRange().getValues();
    if (allData.length < 2) continue;

    let headerRowIdx = -1;
    let nIdx = -1;
    let pIdx = -1;
    let wIdx = -1;
    let rIdx = -1;

    for (let r = 0; r < Math.min(allData.length, 5); r++) {
      let rowLower = allData[r].map(c => c ? c.toString().trim().toLowerCase() : "");
      let foundMember = rowLower.findIndex(c => c === "member" || c.includes("member"));
      if (foundMember > -1) {
        headerRowIdx = r;
        nIdx = foundMember;
        pIdx = rowLower.findIndex(c => c.includes("power"));
        wIdx = rowLower.findIndex(c => c.includes("weekly"));
        rIdx = rowLower.findIndex(c => c.includes("reason") || c.includes("notes"));
        break;
      }
    }

    if (nIdx === -1) continue;

    let firstDataRowSheet = headerRowIdx + 2; 
    let totalDataRows = allData.length - (headerRowIdx + 1);
    let memberNotes = snapSheet.getRange(firstDataRowSheet, nIdx + 1, totalDataRows, 1).getNotes();
    let rowCountThisFile = 0;

    for (let r = headerRowIdx + 1; r < allData.length; r++) {
      let name = allData[r][nIdx] ? allData[r][nIdx].toString().trim() : "";
      if (!name || name.startsWith("---") || name.startsWith("⚠️") || name.toLowerCase() === "member") continue;

      if (typeof CONFIG !== "undefined" && CONFIG.MEMBER_ALIASES && CONFIG.MEMBER_ALIASES[name]) {
        name = CONFIG.MEMBER_ALIASES[name];
      }

      let power = (pIdx > -1 && typeof allData[r][pIdx] === "number") ? allData[r][pIdx] : (typeof cleanNumericValues === "function" ? cleanNumericValues(allData[r][pIdx]) : 0);
      let weekly = (wIdx > -1 && typeof allData[r][wIdx] === "number") ? allData[r][wIdx] : (typeof cleanNumericValues === "function" ? cleanNumericValues(allData[r][wIdx]) : 0);
      let reasonStr = (rIdx > -1 && allData[r][rIdx]) ? allData[r][rIdx].toString() : "";
      
      let noteIndex = r - (headerRowIdx + 1);
      let noteText = (memberNotes[noteIndex] && memberNotes[noteIndex][0]) ? memberNotes[noteIndex][0] : "";

      let combinedRowStatus = `${reasonStr} ${noteText}`.trim();

      if (!historyMap[name]) {
        historyMap[name] = {
          demotions: 0,
          removalCount: 0,
          weeksTracked: 0,
          compliantWeeks: 0,
          totalVsScore: 0,
          powerStart: power
        };
      }

      historyMap[name].weeksTracked += 1;
      historyMap[name].totalVsScore += weekly;

      let isExempt = EXEMPT_REGEX.test(combinedRowStatus);
      let isRemoval = REMOVAL_REGEX.test(combinedRowStatus);
      let isDemotion = DEMOTION_REGEX.test(combinedRowStatus);

      if (isRemoval) {
        historyMap[name].removalCount += 1;
      } else if (isDemotion && !isExempt) {
        historyMap[name].demotions += 1;
      } else {
        historyMap[name].compliantWeeks += 1;
      }

      rowCountThisFile++;
    }

    processedFiles++;
    console.log(`✅ [${dateStr}] Scanned "${snapSheet.getName()}" in "${file.getName()}" -> Extracted ${rowCountThisFile} members.`);
  }

  // Cross-check active sheet notes
  try {
    let curSheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
    let curVals = curSheet.getDataRange().getValues();
    let curHdr = curVals.findIndex(row => row.some(c => c && c.toString().toLowerCase().includes("member")));
    if (curHdr > -1) {
      let curNIdx = curVals[curHdr].findIndex(c => c && c.toString().toLowerCase().includes("member"));
      let activeFirstDataRow = curHdr + 2;
      let activeDataRowsCount = curVals.length - (curHdr + 1);
      let activeNotes = curSheet.getRange(activeFirstDataRow, curNIdx + 1, activeDataRowsCount, 1).getNotes();

      for (let i = curHdr + 1; i < curVals.length; i++) {
        let name = curVals[i][curNIdx] ? curVals[i][curNIdx].toString().trim() : "";
        let noteIdx = i - (curHdr + 1);
        let note = activeNotes[noteIdx] && activeNotes[noteIdx][0] ? activeNotes[noteIdx][0].toLowerCase() : "";
        if (!name || !historyMap[name]) continue;

        let demoMatch = note.match(/demoted last week \(x(\d+)\)/i);
        let removalMatch = note.match(/pending removal \(week (\d+)\)/i);

        if (demoMatch) {
          historyMap[name].demotions = Math.max(historyMap[name].demotions, parseInt(demoMatch[1], 10));
        } else if (removalMatch) {
          historyMap[name].removalCount = Math.max(historyMap[name].removalCount, parseInt(removalMatch[1], 10));
        } else if (note.includes("pending removal")) {
          historyMap[name].removalCount = Math.max(historyMap[name].removalCount + 1, 1);
        }
      }
    }
  } catch (e) {
    console.warn("Active sheet check: " + e.message);
  }

  // Final summary calculations
  Object.keys(historyMap).forEach(name => {
    let m = historyMap[name];
    m.avgVsScore = m.weeksTracked > 0 ? Math.round(m.totalVsScore / m.weeksTracked) : 0;
    m.reliabilityPct = m.weeksTracked > 0 ? Math.min(100, Math.round((m.compliantWeeks / m.weeksTracked) * 100)) : 100;
  });

  PropertiesService.getScriptProperties().setProperty("MEMBER_ARCHIVE_HISTORY", JSON.stringify(historyMap));
  console.log(`🎉 Complete! Successfully processed ${processedFiles} of ${numRows} archive workbooks for ${Object.keys(historyMap).length} members.`);
}

/**
 * Returns roster records combined with compiled multi-week metrics.
 */
function getDashboardData(viewLastWeek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isArchive = (viewLastWeek === true || viewLastWeek === "true");
  const sheetName = isArchive ? "Last Week" : "Data Import";
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found. Please ensure it exists.");
  }

  let historyMap = {};
  try {
    let rawHist = PropertiesService.getScriptProperties().getProperty("MEMBER_ARCHIVE_HISTORY");
    if (rawHist) historyMap = JSON.parse(rawHist);
  } catch (e) {
    console.warn("Could not read history cache: " + e.message);
  }

  const lastRow = Math.max(3, sheet.getLastRow());
  const dataRange = sheet.getRange(3, 1, lastRow - 2, 13);
  const data = dataRange.getValues();

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
    const reason = String(row[12] || "").trim();

    totalPower += power;
    totalWeekly += weekly;

    let cleanName = name.toLowerCase();
    let memberHist = historyMap[name] || null;
    if (!memberHist) {
      let matchedKey = Object.keys(historyMap).find(k => k.trim().toLowerCase() === cleanName);
      if (matchedKey) memberHist = historyMap[matchedKey];
    }

    if (!memberHist) {
      memberHist = {
        demotions: 0,
        removalCount: 0,
        weeksTracked: 1,
        compliantWeeks: 1,
        totalVsScore: weekly,
        avgVsScore: weekly,
        reliabilityPct: 100,
        powerStart: power
      };
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
      reason: reason || "Active",
      history: memberHist
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
  if (password !== "Twisted4Life") {
    return { success: false, error: "Unauthorized: Invalid leadership password." };
  }

  try {
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