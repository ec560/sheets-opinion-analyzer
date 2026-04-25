// global constants for sheet names, cell locations, and color mappings

const ANALYSIS_SHEET_NAME = "Tier Analysis";
const FLAG_SCAN_SHEET_NAME = "Tier Flags";
const VERSION = "v1.4.0-beta";
const TIER_CELL = "B1";
const LEVEL_CELL = "B2";
const DATA_START_ROW = 4;                      // where A:C gets populated
const OUTPUT_COL = 5;                          // column E for results 
const OUTPUT_START_ROW = 1;
const OUTPUT_WIDTH = 4;                        // E:H block

const FLAG_LOW_OPINION_WEIGHT = 4;
const FLAG_SCAN_LOW_OPINION_WEIGHT = 10;
const FLAG_SCAN_NEEDS_MORE_OPINIONS_MAX_WEIGHT = 20;
const FLAG_BOOK_LEAN_WEIGHT = 1.5;
const FLAG_LOCK_SHARE = 0.75;
const FLAG_LOCK_MIN_RAW_COUNT = 50;
const FLAG_LOCK_MIN_WEIGHT = 50;
const COUNTED_PLAYER_HIGHLIGHT = "#e6f4ea";

function isAnalyzerUtilitySheetName_(name) {
  return name === ANALYSIS_SHEET_NAME || name === FLAG_SCAN_SHEET_NAME;
}

// reliability background color = multiplier
const reliabilityFactors = {
  "#00ffff": 1.25,
  "#00ff00": 1,
  "#ffff00": 0.75,
  "#ff9900": 0.5,
  "#ff0000": 0,
  "#ff00ff": 0,
  "#000000": 0
};

// difficulty background color = tier name
const difficultyColorNames = {
  "#0000ff": "Insane Demon",
  "#0b5394": "Insane Demon", // hard/insane demon
  "#4a86e8": "Beginner",
  "#00ffff": "Easy",
  "#00ff00": "Medium",
  "#ffff00": "Hard",
  "#ff9900": "Very Hard",
  "#ff0000": "Insane",
  "#ff00ff": "Extreme",
  "#9900ff": "Remorseless",
  "#b087eb": "Relentless",
  "#f19eea": "Terrifying",
  "#ea6661": "Catastrophic",
  "#ffc183": "Inexorable",
  "#ffe599": "Excruciating",
  "#a7e58d": "Merciless",
  "#5bad96": "Monstrous",
  "#528cb1": "Apocalyptic",
  "#6d6ab0": "Demonic",
  "#9452a2": "Menacing",
  "#913869": "Unreal",
  "#832828": "Nightmare"
};

// split color = [lowerTierColor, higherTierColor]
const splitPairs = {
  "#1155cc": ["#0000ff", "#4a86e8"],
  "#31c0f0": ["#4a86e8", "#00ffff"],
  "#00ff80": ["#00ffff", "#00ff00"],
  "#80ff00": ["#00ff00", "#ffff00"],
  "#ffcc00": ["#ffff00", "#ff9900"],
  "#ff4d00": ["#ff9900", "#ff0000"],
  "#ff0080": ["#ff0000", "#ff00ff"],
  "#cc00ff": ["#ff00ff", "#9900ff"],
  "#a45cf6": ["#9900ff", "#b087eb"],
  "#d193eb": ["#b087eb", "#f19eea"],
  "#ee82a6": ["#f19eea", "#ea6661"],
  "#f59472": ["#ea6661", "#ffc183"],
  "#ffd579": ["#ffc183", "#ffe599"],
  "#d3ec84": ["#ffe599", "#a7e58d"],
  "#7bd296": ["#a7e58d", "#5bad96"],
  "#499da6": ["#5bad96", "#528cb1"],
  "#6279b5": ["#528cb1", "#6d6ab0"],
  "#805dab": ["#6d6ab0", "#9452a2"],
  "#a6478b": ["#9452a2", "#913869"],
  "#903242": ["#913869", "#832828"]
};

// Tier ordering for split logic (low -> high)
const orderedTierNames = [
  "Insane Demon",
  "Beginner",
  "Easy",
  "Medium",
  "Hard",
  "Very Hard",
  "Insane",
  "Extreme",
  "Remorseless",
  "Relentless",
  "Terrifying",
  "Catastrophic",
  "Inexorable",
  "Excruciating",
  "Merciless",
  "Monstrous",
  "Apocalyptic",
  "Demonic",
  "Menacing",
  "Unreal",
  "Nightmare"
];
