/**
 * Discord.gs
 * Handles all Webhook communications between the TWST Roster Data Management sheet and Discord.
 */

const MASTER_WEBHOOK_URL = "https://discord.com/api/webhooks/1559018614131003552/PTmKCLPebPVnTGimYad5tgX5htH29az3v9uMa3VJAz2qoYT6p_Go8r0VrDdCl9cjy7-l";
const DASHBOARD_SITE_URL = "https://sites.google.com/view/twisted-operations/twst-ops";

/**
 * Shared helper that parses roster rows from either "Data Import" or "Last Week"
 * using the exact same Row 3 range logic as WebApp.gs.
 */
function buildRosterEmbedFields_(sheet, isFinalReport) {
  const lastRow = Math.max(3, sheet.getLastRow());
  const data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();

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

  const parseNum = (val) => {
    if (typeof val === "number") return val;
    return parseFloat(String(val).replace(/[^0-9.-]+/g, "")) || 0;
  };

  data.forEach(row => {
    const name = String(row[1] || "").trim();
    if (!name || name.startsWith("---") || name.startsWith("⚠️") || name.toLowerCase() === "member") return;

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

  const makeNameField = (title, emoji, arr) => {
    return {
      "name": `${emoji} ${title} (${arr.length})`,
      "value": formatList(arr),
      "inline": false
    };
  };

  const makeCountField = (title, emoji, arr) => {
    return {
      "name": `${emoji} ${title}`,
      "value": `**${arr.length}** ${arr.length === 1 ? "Member" : "Members"}`,
      "inline": false
    };
  };

  const demotionLabel = isFinalReport ? "Demotions Applied" : "Pending Demotions";
  const removalLabel = isFinalReport ? "Removals Issued" : "Immediate Removal";

  return [
    { "name": "🏥 Alliance Health", "value": `**${healthPct}%**\n${healthPct >= 90 ? "Optimal" : "Needs Review"}`, "inline": true },
    { "name": "👥 Roster Capacity", "value": `**${memberCount} / 100**\n${memberCount === 100 ? "Full Capacity" : "Recruiting"}`, "inline": true },
    { "name": "⚔️ Total Power", "value": `**${formatSuffix(totalPower)}**\nActive Combat Power`, "inline": true },

    // Visual Separator Field
    { "name": "\u200B", "value": "▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬", "inline": false },

    makeNameField("Crushed It!", "🔥", crushedIt),
    makeNameField("Perfect", "🌟", perfect),
    makeNameField("Promoted / Restored", "⬆️", promotions),
    makeCountField("Maintained", "🛡️", maintained),
    makeNameField("New Member", "👋", newMember),
    makeNameField("Exempt", "⏸️", exempt),
    makeNameField(demotionLabel, "⚠️", demotions),
    makeNameField(removalLabel, "⛔", removals)
  ];
}

function updateLiveDiscordDashboard() {
  const webhookUrl = MASTER_WEBHOOK_URL;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Import");

  if (!sheet) return;

  const fields = buildRosterEmbedFields_(sheet, false);

  const payload = {
    "embeds": [
      {
        "title": "📊 TWISTED OPERATIONS",
        "url": DASHBOARD_SITE_URL,
        "description": "Live Alliance Performance Metrics • Policy Compliance • Roster Tracking\n👆 *(Click the header above to open the full interactive dashboard)*\n\u200B",
        "color": 2045558,
        "fields": fields,
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

/**
 * Helper to build the Sunday EOW payload from the "Last Week" sheet.
 */
function buildSundayEOWPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Last Week");
  if (!sheet) return null;

  const fields = buildRosterEmbedFields_(sheet, true);

  return {
    "content": "📢 **Sunday End-of-Week Reset Complete!**",
    "embeds": [
      {
        "title": "📊 TWST Alliance Weekly Performance (Final Status)",
        "url": DASHBOARD_SITE_URL,
        "description": `The week has officially closed. Ranks have been updated and penalties applied. Below is last week's final status breakdown.\n👆 *(Click the header above to open the full interactive dashboard)*\n\u200B`,
        "color": 3447003,
        "fields": fields,
        "footer": { "text": "TWSTOps Automation Engine • Week Closed: " + new Date().toLocaleString() }
      }
    ]
  };
}

function sendSundayEOWReport() {
  const payload = buildSundayEOWPayload_();
  if (!payload) return;

  UrlFetchApp.fetch(MASTER_WEBHOOK_URL, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  });
}

// ============================================================================
// TEMPORARY ONE-TIME UPDATE FUNCTION (Delete after running once)
// ============================================================================
function tempPatchExistingSundayCard() {
  const targetInput = "1559638475324526683";

  const payload = buildSundayEOWPayload_();
  if (!payload) return;

  const cleanMessageId = targetInput.trim().split("/").pop();
  const modernWebhookUrl = MASTER_WEBHOOK_URL.replace("discordapp.com", "discord.com").replace(/\/$/, "");
  const endpoint = modernWebhookUrl + "/messages/" + cleanMessageId;

  const res = UrlFetchApp.fetch(endpoint, {
    "method": "patch",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  });

  console.log("Response Code: " + res.getResponseCode() + " | " + res.getContentText());
}
