/**
 * WebApp.gs - Google Apps Script Backend for TWST Command Center Web App
 * TWST Roster Management System
 */

function doGet(e) {
  let forceArchive = (e && e.parameter && (e.parameter.view === 'archive' || e.parameter.archive === 'true')) ? true : false;
  let template = HtmlService.createTemplateFromFile("App");
  template.forceArchive = forceArchive;
  
  return template.evaluate()
    .setTitle("TWST Ops Command Center")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Normalizes member names and strips squad tags (" D1", " C2", " A4b", etc.)
 * so a player's multi-week history stays connected across all weeks.
 */
function normalizeMemberKey_(rawName) {
  if (!rawName) return "";
  let s = rawName.toString().trim();
  if (typeof CONFIG !== "undefined" && CONFIG.MEMBER_ALIASES && CONFIG.MEMBER_ALIASES[s]) {
    s = CONFIG.MEMBER_ALIASES[s];
  }
  s = s.replace(/\s+[A-Da-d][1-4][a-bA-B]?$/i, "").trim();
  return s.toLowerCase();
}

/**
 * Plain, strict string-based status checker (Zero Regex!).
 */
function classifyRowStatus_(reasonStr, noteStr, curRank, newRank) {
  let r = (reasonStr || "").toString().toLowerCase();
  let n = (noteStr || "").toString().toLowerCase();
  let cr = (curRank || "").toString().trim().toUpperCase();
  let nr = (newRank || "").toString().trim().toUpperCase();

  let isExempt = r.includes("[exempt]") || r.includes("grace period") || r.includes("(rn)") ||
                 n.includes("exempt") || n.includes("vacation") || n.includes("loa") || n.includes("sick") || n.includes("personal");

  let isRemoval = r.includes("[removal]") || 
                  r.includes("[immediate removal]") || 
                  n.includes("pending removal") ||
                  nr === "R0";

  let isDemotion = !isRemoval && (
                   r.includes("[demotion]") || 
                   r.includes("[demoted]") || 
                   n.includes("demoted last week") ||
                   (cr === "R3" && (nr === "R2" || nr === "R1")) ||
                   (cr === "R2" && nr === "R1")
                 );

  let isRestored = r.includes("[promotion]") || r.includes("restored to r3");

  return { isRemoval, isDemotion, isExempt, isRestored };
}

/**
 * Universal Archive Analytics Compiler.
 * Exclusively scans the 'Data Import' tab across all archive workbooks in chronological order.
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

  console.log(`🔍 Found ${numRows} archive entries in "Archive Log". Scanning "Data Import" per archive...`);

  const parseNum = (val) => {
    if (!val) return 0;
    if (typeof val === "number") return val;
    let n = parseFloat(val.toString().replace(/[^0-9.-]+/g, ""));
    return isNaN(n) ? 0 : Math.round(n);
  };

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

    // Exclusively read "Data Import"
    let snapSheet = file.getSheetByName("Data Import");
    if (!snapSheet) {
      console.warn(`⚠️ Row ${a + 2} (${dateStr}): "Data Import" tab not found in "${file.getName()}".`);
      continue;
    }

    let allData = snapSheet.getDataRange().getValues();
    if (allData.length < 2) continue;

    let headerRowIdx = -1;
    let rankIdx = -1;
    let nIdx = -1;
    let pIdx = -1;
    let wIdx = -1;
    let newRankIdx = -1;
    let rIdx = -1;

    for (let r = 0; r < Math.min(allData.length, 5); r++) {
      let rowLower = allData[r].map(c => c ? c.toString().trim().toLowerCase() : "");
      let foundMember = rowLower.findIndex(c => c === "member" || c.includes("member"));
      if (foundMember > -1) {
        headerRowIdx = r;
        rankIdx = rowLower.findIndex(c => c === "rank");
        nIdx = foundMember;
        pIdx = rowLower.findIndex(c => c.includes("power"));
        wIdx = rowLower.findIndex(c => c.includes("weekly"));
        newRankIdx = rowLower.findIndex(c => c.includes("new rank"));
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
      let rawName = allData[r][nIdx] ? allData[r][nIdx].toString().trim() : "";
      if (!rawName || rawName.startsWith("---") || rawName.startsWith("⚠️") || rawName.toLowerCase() === "member") continue;

      let key = normalizeMemberKey_(rawName);
      if (!key) continue;

      let curRank = (rankIdx > -1 && allData[r][rankIdx]) ? allData[r][rankIdx].toString().trim().toUpperCase() : "";
      let newRank = (newRankIdx > -1 && allData[r][newRankIdx]) ? allData[r][newRankIdx].toString().trim().toUpperCase() : "";
      let power = pIdx > -1 ? parseNum(allData[r][pIdx]) : 0;
      let weekly = wIdx > -1 ? parseNum(allData[r][wIdx]) : 0;
      let reasonStr = (rIdx > -1 && allData[r][rIdx]) ? allData[r][rIdx].toString() : "";

      let noteIndex = r - (headerRowIdx + 1);
      let noteText = (memberNotes[noteIndex] && memberNotes[noteIndex][0]) ? memberNotes[noteIndex][0] : "";

      if (!historyMap[key]) {
        historyMap[key] = {
          displayName: rawName,
          demotions: 0,
          removalCount: 0,
          removalStreak: 0,
          weeksTracked: 0,
          vsWeeksCount: 0,
          compliantWeeks: 0,
          totalVsScore: 0,
          // Strictly record their starting power from the first archive they joined/appeared in!
          powerStart: power > 0 ? power : 0,
          lastRank: curRank
        };
      }

      let h = historyMap[key];
      // Only set powerStart if it was 0 in their first week; never overwrite it later!
      if ((!h.powerStart || h.powerStart === 0) && power > 0) {
        h.powerStart = power;
      }

      h.weeksTracked += 1;
      if (weekly > 0) {
        h.totalVsScore += weekly;
        h.vsWeeksCount += 1;
      }

      let st = classifyRowStatus_(reasonStr, noteText, curRank, newRank);

      if ((curRank === "R2" || curRank === "R1" || st.isRestored) && h.demotions === 0 && !st.isDemotion) {
        h.demotions = 1;
      }

      if (st.isRemoval && !st.isExempt) {
        h.removalCount += 1;
        h.removalStreak += 1;
      } else if (st.isDemotion && !st.isExempt) {
        h.demotions += 1;
        h.removalStreak = 0;
      } else {
        h.removalStreak = 0;
        h.compliantWeeks += 1;
      }

      h.lastRank = newRank || curRank;
      rowCountThisFile++;
    }

    processedFiles++;
    console.log(`✅ [${dateStr}] Scanned "Data Import" in "${file.getName()}" -> Extracted ${rowCountThisFile} members.`);
  }

  // Incorporate current live Data Import sheet (Current Week)
  try {
    let curSheet = ss.getSheetByName("Data Import") || ss.getActiveSheet();
    let curVals = curSheet.getDataRange().getValues();
    let curHdr = curVals.findIndex(row => row.some(c => c && c.toString().toLowerCase().includes("member")));
    if (curHdr > -1) {
      let rowLower = curVals[curHdr].map(c => c ? c.toString().trim().toLowerCase() : "");
      let rankIdx = rowLower.findIndex(c => c === "rank");
      let curNIdx = rowLower.findIndex(c => c.includes("member"));
      let pIdx = rowLower.findIndex(c => c.includes("power"));
      let wIdx = rowLower.findIndex(c => c.includes("weekly"));
      let newRankIdx = rowLower.findIndex(c => c.includes("new rank"));
      let rIdx = rowLower.findIndex(c => c.includes("reason") || c.includes("notes"));

      let activeFirstDataRow = curHdr + 2;
      let activeDataRowsCount = curVals.length - (curHdr + 1);
      let activeNotes = curSheet.getRange(activeFirstDataRow, curNIdx + 1, activeDataRowsCount, 1).getNotes();

      for (let i = curHdr + 1; i < curVals.length; i++) {
        let rawName = curVals[i][curNIdx] ? curVals[i][curNIdx].toString().trim() : "";
        if (!rawName || rawName.startsWith("---") || rawName.startsWith("⚠️") || rawName.toLowerCase() === "member") continue;

        let key = normalizeMemberKey_(rawName);
        let noteIdx = i - (curHdr + 1);
        let note = activeNotes[noteIdx] && activeNotes[noteIdx][0] ? activeNotes[noteIdx][0] : "";
        let reasonStr = (rIdx > -1 && curVals[i][rIdx]) ? curVals[i][rIdx].toString() : "";
        let curRank = (rankIdx > -1 && curVals[i][rankIdx]) ? curVals[i][rankIdx].toString().trim().toUpperCase() : "";
        let newRank = (newRankIdx > -1 && curVals[i][newRankIdx]) ? curVals[i][newRankIdx].toString().trim().toUpperCase() : "";
        let power = pIdx > -1 ? parseNum(curVals[i][pIdx]) : 0;
        let weekly = wIdx > -1 ? parseNum(curVals[i][wIdx]) : 0;

        let st = classifyRowStatus_(reasonStr, note, curRank, newRank);

        if (!historyMap[key]) {
          // Brand new member this week: starting power is current power
          historyMap[key] = {
            displayName: rawName,
            demotions: (st.isDemotion || curRank === "R2" || curRank === "R1" || st.isRestored) ? 1 : 0,
            removalCount: st.isRemoval ? 1 : 0,
            removalStreak: st.isRemoval ? 1 : 0,
            weeksTracked: 1,
            vsWeeksCount: weekly > 0 ? 1 : 0,
            compliantWeeks: (st.isRemoval || st.isDemotion) ? 0 : 1,
            totalVsScore: weekly,
            powerStart: power,
            lastRank: newRank || curRank
          };
        } else {
          let h = historyMap[key];
          if ((!h.powerStart || h.powerStart === 0) && power > 0) {
            h.powerStart = power;
          }

          if (weekly > 0 || st.isRemoval || st.isDemotion) {
            h.weeksTracked += 1;
            if (weekly > 0) {
              h.totalVsScore += weekly;
              h.vsWeeksCount += 1;
            }
            if (st.isRemoval && !st.isExempt) {
              h.removalCount += 1;
              h.removalStreak += 1;
            } else if (st.isDemotion && !st.isExempt) {
              h.demotions += 1;
              h.removalStreak = 0;
            } else {
              h.removalStreak = 0;
              h.compliantWeeks += 1;
            }
          }

          if ((st.isRestored || curRank === "R2" || curRank === "R1") && h.demotions === 0) {
            h.demotions = 1;
          }

          // Check explicit note counters
          let demoMatch = note.match(/demoted last week \(x(\d+)\)/i);
          let removalMatch = note.match(/pending removal \(week (\d+)\)/i);
          if (demoMatch) h.demotions = Math.max(h.demotions, parseInt(demoMatch[1], 10));
          if (removalMatch) {
            h.removalCount = Math.max(h.removalCount, parseInt(removalMatch[1], 10));
            h.removalStreak = Math.max(h.removalStreak, parseInt(removalMatch[1], 10));
          }
        }
      }
    }
  } catch (e) {
    console.warn("Active sheet check: " + e.message);
  }

  // Final summary calculations
  Object.keys(historyMap).forEach(key => {
    let m = historyMap[key];
    let divisor = m.vsWeeksCount > 0 ? m.vsWeeksCount : Math.max(1, m.weeksTracked);
    m.avgVsScore = m.totalVsScore > 0 ? Math.round(m.totalVsScore / divisor) : 0;
    
    let totalWks = Math.max(1, m.weeksTracked);
    let goodWks = Math.min(totalWks, Math.max(0, m.compliantWeeks || 0));
    m.reliabilityPct = Math.min(100, Math.max(0, Math.round((goodWks / totalWks) * 100)));
  });

  PropertiesService.getScriptProperties().setProperty("MEMBER_ARCHIVE_HISTORY", JSON.stringify(historyMap));
  console.log(`🎉 Complete! Successfully processed ${processedFiles} of ${numRows} archive workbooks + current week for ${Object.keys(historyMap).length} members.`);
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

    let normKey = normalizeMemberKey_(name);
    let memberHist = historyMap[normKey] || historyMap[name] || historyMap[name.toLowerCase()] || null;
    if (!memberHist) {
      let matchedKey = Object.keys(historyMap).find(k => normalizeMemberKey_(k) === normKey);
      if (matchedKey) memberHist = historyMap[matchedKey];
    }

    if (!memberHist) {
      let isRem = reason.toLowerCase().includes("[removal]");
      let isDem = reason.toLowerCase().includes("[demotion]") || reason.toLowerCase().includes("restored to r3") || rank === "R2" || rank === "R1";
      memberHist = {
        demotions: isDem ? 1 : 0,
        removalCount: isRem ? 1 : 0,
        removalStreak: isRem ? 1 : 0,
        weeksTracked: 1,
        compliantWeeks: (isRem || isDem) ? 0 : 1,
        totalVsScore: weekly,
        avgVsScore: weekly,
        reliabilityPct: (isRem || isDem) ? 0 : 100,
        powerStart: power
      };
    } else {
      if ((reason.toLowerCase().includes("restored to r3") || rank === "R2" || rank === "R1") && (!memberHist.demotions || memberHist.demotions === 0)) {
        memberHist.demotions = 1;
      }
      if ((!memberHist.powerStart || memberHist.powerStart === 0) && power > 0) {
        memberHist.powerStart = power;
      }
      if (!memberHist.avgVsScore || memberHist.avgVsScore === 0) {
        memberHist.avgVsScore = weekly;
      }
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