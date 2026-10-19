/**
 * Config.gs - Global Configuration & Constants
 * TWST Roster Management System
 */

const CONFIG = {
  INBOX_FOLDER_ID: "1HQ4-OWvAzlkIVE4gzfLzkhCpo18ePoB8", // 📥 Drop unorganized screenshots here
  FOLDER_MAPPINGS: {
    "Mon": "1zRN-WmsTmUggDzToInvYGxckLnU-_3VC",
    "Tue": "1okmGGXrlBHeYBAhPbWa5v1PUvR-2-ynW",
    "Wed": "17QusA-UjVGudgxpuA-a8uHlSl1UFeZv5",
    "Thu": "1l7Hfj8TMvXMcYw1MU2yrK7iMPZg9Ypcj",
    "Fri": "1oA64SF_yAeGGXGOsWvIaC66h-sYBf7i-",
    "Sat": "1YqdFajIdHu-BeI2Y7jwAP-fv6v1EJSPN",
    "Strength Ranking - Power": "1BQGf-BGqDPFWegAOm4XjR2Bqrj1LN1pK",
    "Strength Ranking - Donations": "151krneCLJzPtMALV57A7nrnjU-4ypijn"
  },
  PROCESSED_FOLDER_ID: "18XRS9GqMEwnFef5qBLPQQro1MviJcmQm",
  ARCHIVE_FOLDER_ID: "", // Optional Drive Folder ID for Sunday weekly backup copies
  PROCESSED_SUBFOLDER_NAMES: {
    "Mon": "Monday VS",
    "Tue": "Tuesday VS",
    "Wed": "Wednesday VS",
    "Thu": "Thursday VS",
    "Fri": "Friday VS",
    "Sat": "Saturday VS",
    "Strength Ranking - Power": "Members",
    "Strength Ranking - Donations": "Donations"
  },
  SUMMARY_LABEL_MAP: {
    "Strength Ranking - Power": "Pow",
    "Strength Ranking - Donations": "Don",
    "Mon": "Mon",
    "Tue": "Tue",
    "Wed": "Wed",
    "Thu": "Thu",
    "Fri": "Fri",
    "Sat": "Sat"
  },
  HEADER_ALIASES: {
    "Strength Ranking - Power": ["power"],
    "Strength Ranking - Donations": ["donations"],
    "Mon": ["mon", "monday"],
    "Tue": ["tue", "tuesday"],
    "Wed": ["wed", "wednesday"],
    "Thu": ["thu", "thursday"],
    "Fri": ["fri", "friday"],
    "Sat": ["sat", "saturday"]
  },
  
  // Member Alias Dictionary 
  // Map shortened/alternate OCR screenshot names (left) to their EXACT Sheet name (right)
  MEMBER_ALIASES: {
    "Red": "Reddeadeye123",
    "ZioooooooooooooooON": "ZioooooooooooooooN",
    "Apostle T Hoarder": "Thot Hoarder",
    "ZioooooooooooooooON ": "ZioooooooooooooooN",
    "ZiooooooooooooooooooN ": "ZioooooooooooooooN",
    "Not the one": "N0t th3 0n3",
    "N0t th3 On3": "NOt th3 On3",
    "N0t th3 0n3": "NOt th3 On3",
    "McNasty D1": "McNastyyyyy",
    "McNasty": "McNastyyyyy",
    "Loocslydna": "Loocsiydna",
    "At hutto D1": "At hutto D2",
    "SoulCollector0 D1": "Soulcollector0",
    "M3ATLOAF": "DONK333Y",
    "WasiWarrior2": "WasiWarrior"
  },
  
  // ==========================================
  // RANK CALCULATION & REQUIREMENT CONFIGURATION
  // ==========================================
  TARGETS: {
    DEMOTION_MIN_DAILY_VS_POINTS: 3600000,      
    DEMOTION_MIN_DAILY_DONATIONS: 3000,         
    REQUIRED_MISSED_COUNT_FOR_DEMOTION: 3,      
    TAG_DEMOTION: "[Demotion]",                 

    PROMOTION_RESTORE_TO_R3: true,              
    PROMOTION_MIN_VS_POINTS: 3600000,           
    PROMOTION_MIN_DAILY_DONATIONS: 3000,        
    TAG_PROMOTION: "[Promotion]",               

    CRUSHED_IT_MIN_VS_SCORE: 7200000,           
    CRUSHED_IT_ELITE_DONATIONS: 6000,           
    TAG_CRUSHED_IT: "[Crushed It!]",            
    TAG_MAINTAINED: "[Maintained]",             
    TAG_PERFECT: "[Perfect]",                   

    TAG_IMMEDIATE_REMOVAL: "[Removal]",         
    PROTECT_RANKS: ["R4", "R5"],
    EXEMPT_KEYWORDS: ["exempt", "vacation", "loa", "personal", "absence", "sick"],                

    MSG_NEW_MEMBER: "Initialized to R3",
    MSG_CRUSHED_IT: "Opened all 9 VS boxes and doubled daily donations!",
    MSG_PERFECT: "100% VS & Donations targets met successfully.",
    MSG_MID_WEEK_STATUS: "Mid-week progress tracking"
  },

  COLORS: {
    MATCHED_MEMBER: "#d9ead3",
    ELIMINATION_ALERT: "#f5a623",
    MISSING_DATA: "#ffc8aa",
    DAILY_SUCCESS: "#d9ead3"
  }
};
