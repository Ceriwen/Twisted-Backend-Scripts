/**
 * Discord.gs
 * Handles all Webhook communications between the TWST Roster Data Management sheet and Discord.
 */

const MASTER_WEBHOOK_URL = "https://discordapp.com/api/webhooks/1559018614131003552/PTmKCLPebPVnTGimYad5tgX5htH29az3v9uMa3VJAz2qoYT6p_Go8r0VrDdCl9cjy7-l";

function updateLiveDiscordDashboard() {
  const webhookUrl = MASTER_WEBHOOK_URL;
  const webAppUrl = "https://sites.google.com/view/twisted-operations/twst-ops";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");
  
  if (!sheet) return;

  let rRange = (typeof findRosterRange === "function") ? findRosterRange(sheet) : { startRow: 2, rowCount: sheet.getLastRow(), colCount: 13 };
  let startRow = rRange.startRow + 1;
  let numRows = Math.max(1, rRange.rowCount - 1);
  const data = sheet.getRange(startRow, 1, numRows, Math.max(13, rRange.colCount)).getValues();

  let memberCount = 0;
  let totalPower = 0;
  let compliantCount = 0; 
  
  let crushedIt = [];
  let perfect = [];
  let promotions = [];
  let maintained = [];
  let newMember = [];
  let exempt = [];
  let demotions = [];
  let removals = [];

  const parseNum = (val) => parseFloat(String(val).replace(/[^0-9.-]+/g, "")) || 0;

  data.forEach(row => {
    const name = String(row[1] || "").trim();
    if (!name || name.startsWith("---") || name.startsWith("⚠️")) return;

    memberCount++;
    totalPower += parseNum(row[2]); 
    
    const reason = String(row[12] || "").trim(); 
    
    if (!reason.includes("[Demotion]") && !reason.includes("[Removal]")) {
      compliantCount++;
    }
    
    if (reason.includes("[Crushed It!]")) crushedIt.push(name);
    if (reason.includes("[Perfect]")) perfect.push(name);
    if (reason.includes("[Promotion]")) promotions.push(name);
    if (reason.includes("[Maintained]")) maintained.push(name);
    if (reason.includes("[New Member]")) newMember.push(name);
    if (reason.includes("[Exempt]")) exempt.push(name);
    if (reason.includes("[Demotion]")) demotions.push(name);
    if (reason.includes("[Removal]")) removals.push(name);
  });

  const formatSuffix = (num) => {
    if (num >= 1000000000) return (num / 1000000000).toFixed(2) + "B";
    if (num >= 1000000) return (num / 1000000).toFixed(2) + "M";
    if (num >= 1000) return (num / 1000).toFixed(1) + "K";
    return num.toString();
  };

  const healthPct = memberCount > 0 ? Math.round((compliantCount / memberCount) * 100) : 0;
  
  const formatList = (arr) => {
    if (arr.length === 0) return "None";
    let str = arr.join(", ");
    return str.length > 1024 ? str.substring(0, 1020) + "..." : str;
  };

  const makeField = (title, emoji, arr) => {
    let fieldName = `${emoji} ${title} (${arr.length})`;
    return { "name": fieldName, "value": formatList(arr), "inline": false };
  };

  const payload = {
    "embeds": [
      {
        "title": "📊 TWISTED OPERATIONS",
        "url": webAppUrl,
        "description": "Live Alliance Performance Metrics • Policy Compliance • Roster Tracking\n👆 *(Click the header above to open the full interactive dashboard)*\n\u200B",
        "color": 2045558,
        "fields": [
          { "name": "🏥 Alliance Health", "value": `**${healthPct}%**\n${healthPct >= 90 ? "Optimal" : "Needs Review"}`, "inline": true },
          { "name": "👥 Roster Capacity", "value": `**${memberCount} / 100**\n${memberCount === 100 ? "Full Capacity" : "Recruiting"}`, "inline": true },
          { "name": "⚔️ Total Power", "value": `**${formatSuffix(totalPower)}**\nActive Combat Power`, "inline": true },
          
          makeField("Crushed It!", "🔥", crushedIt),
          makeField("Perfect", "🌟", perfect),
          makeField("Promoted / Restored", "⬆️", promotions),
          makeField("Maintained", "🛡️", maintained),
          makeField("New Member", "👋", newMember),
          makeField("Exempt", "⏸️", exempt),
          makeField("Pending Demotions", "⚠️", demotions),
          makeField("Immediate Removal", "⛔", removals)
        ],
        "footer": { "text": "Last Sync: " + new Date().toLocaleString() }
      }
    ]
  };

  let props = PropertiesService.getScriptProperties();
  let messageId = props.getProperty("DISCORD_DASHBOARD_MSG_ID");
  
  let endpoint = webhookUrl;
  let method = "post";
  
  if (messageId) {
    endpoint = webhookUrl + "/messages/" + messageId;
    method = "patch";
  } else {
    endpoint = webhookUrl + "?wait=true"; 
  }

  let options = {
    "method": method,
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  let res = UrlFetchApp.fetch(endpoint, options);
  let code = res.getResponseCode();

  if (method === "patch" && code === 404) {
    endpoint = webhookUrl + "?wait=true";
    options.method = "post";
    res = UrlFetchApp.fetch(endpoint, options);
    code = res.getResponseCode();
  }

  if (code === 200 || code === 204) {
    if (options.method === "post") {
      let json = JSON.parse(res.getContentText());
      props.setProperty("DISCORD_DASHBOARD_MSG_ID", json.id);
    }
    console.log("Success! Discord dashboard updated.");
    try { SpreadsheetApp.getActiveSpreadsheet().toast("Live dashboard updated successfully!", "✅ Success"); } catch(e) {}
  } else {
    console.error("Discord Error " + code + ": " + res.getContentText());
    try { SpreadsheetApp.getActiveSpreadsheet().toast("Discord Error: " + code, "❌ Error"); } catch(e) {}
  }
}

function sendSundayEOWReport() {
  const webhookUrl = MASTER_WEBHOOK_URL;
  const webAppUrl = "https://script.google.com/macros/s/AKfycbwRXwImXVcxMh8O6_QxPiKM3ErrGCxdqRCQ1-692fzjXY0upU1H7ZrXNUTLQAJ9t0hF-Q/exec";
  const archiveLink = webAppUrl + "?archive=true";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Last Week");
  if (!sheet) return;

  const data = sheet.getRange(3, 1, Math.max(3, sheet.getLastRow()) - 2, 13).getValues();
  
  let compliantCount = 0;
  let activeMembers = 0;
  let totalPower = 0;
  let totalWeekly = 0;

  const parseNum = (val) => parseFloat(String(val).replace(/[^0-9.-]+/g, "")) || 0;

  data.forEach(row => {
    const name = String(row[1] || "").trim();
    if (!name || name.startsWith("---") || name.startsWith("⚠️")) return;
    
    activeMembers++;
    totalPower += parseNum(row[2]);
    totalWeekly += parseNum(row[9]);
    
    const reason = String(row[12] || "");
    if (!reason.includes("[Demotion]") && !reason.includes("[Removal]")) {
      compliantCount++;
    }
  });

  const healthPct = activeMembers > 0 ? Math.round((compliantCount / activeMembers) * 100) : 0;
  const formatSuffix = (num) => {
    if (num >= 1000000000) return (num / 1000000000).toFixed(2) + "B";
    if (num >= 1000000) return (num / 1000000).toFixed(2) + "M";
    return num.toLocaleString();
  };

  const payload = {
    content: "📢 **Sunday End-of-Week Reset Complete!**",
    embeds: [{
      title: "📊 TWST Alliance Weekly Performance",
      description: `The week has officially closed. Ranks have been updated and penalties applied. Click the link below to view the locked-in data.\n\n[**👉 View Final Dashboard Here**](${archiveLink})`,
      color: 3447003,
      fields: [
        { name: "Alliance Health", value: `${healthPct}% Compliant`, inline: true },
        { name: "Roster Capacity", value: `${activeMembers} / 100`, inline: true },
        { name: "Total Power", value: formatSuffix(totalPower), inline: true },
        { name: "Weekly VS Score", value: formatSuffix(totalWeekly), inline: true }
      ],
      footer: { text: "TWSTOps Automation Engine" },
      timestamp: new Date().toISOString()
    }]
  };

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });
}
