// opayc business
// Version 3
// 2026

/*************** CONFIG VALUES ***************/
const ANALYSIS_SHEET_NAME = "Tier Analysis";
const TIER_CELL = "B1";
const LEVEL_CELL = "B2";
const DATA_START_ROW = 4;                      // where A:C gets populated
const OUTPUT_COL = 5;                          // column E for results 
const OUTPUT_START_ROW = 1;
const OUTPUT_WIDTH = 4;                        // E:H block

function hex_(color) {
  if (color == null) return "";
  // If it's already a hex string like "#ff0000"
  if (typeof color === "string") {
    return color.trim().toLowerCase();
  }

  // Google Apps Script Color object (from getFontColorObjects)
  try {
    if (typeof color.asRgbColor === "function") {
      return color.asRgbColor().asHexString().toLowerCase();
    }
  } catch (e) {}

  // Fallback
  return String(color).trim().toLowerCase();
}

/*************** STRUCTS ***************/
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

function tierTextColor_(tierName) {
  if (tierName === "Insane Demon") return "#ffffff"; // white on blue
  if (tierName === "Fuck") return "#ff0000"; // red on black
  return "#000000";
}

/*************** MENU & SETUP ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tier Tools")
    .addItem("Setup", "setupTierAnalysis")
    .addItem("Refresh", "populateSelectedLevel")
    .addItem("Analyze Selected Level", "analyzeSelectedLevel")
    .addToUi();
}

function buildTierNameToColor_() {
  const out = {};
  Object.keys(difficultyColorNames).forEach(hex => {
    const name = difficultyColorNames[hex];
    if (!(name in out)) out[name] = hex_(hex); // don't overwrite
  });
  out["Fuck"] = "#000000";
  return out;
}

function tierToSparkColor_(tierName) {
  const nameToColor = buildTierNameToColor_();
  return nameToColor[tierName] || "#999999";
}

function applyTierSheetColor_(tool, r0, c0, outRowCount) {
  const labels = tool.getRange(r0, c0, outRowCount, 1).getDisplayValues().flat();
  const tierSheetRowOff = labels.findIndex(v => String(v).trim() === "Tier sheet");
  if (tierSheetRowOff < 0) return;

  const r = r0 + tierSheetRowOff;
  const tierName = String(tool.getRange(r, c0 + 1).getDisplayValue()).trim(); // value cell
  if (!tierName) return;

  const nameToColor = buildTierNameToColor_();
  const hex = nameToColor[tierName];
  if (!hex) return; // if sheet name isn't one of your tier names, skip

  tool.getRange(r, c0 + 1)
    .setBackground(hex)
    .setFontColor(tierTextColor_(tierName))
    .setFontWeight("bold");
}



function buildCumEnds_(orderedTierNames, weightsByTier) {
  const cumEnds = [];
  let cum = 0;
  for (let i = 0; i < orderedTierNames.length; i++) {
    cum += (weightsByTier[orderedTierNames[i]] || 0);
    cumEnds.push(cum);
  }
  return cumEnds; // cumEnds[i] = cumulative after tier i
}

function intervalDistance_(half, low, high) {
  if (half < low) return low - half;
  if (half > high) return half - high;
  return 0; // half inside interval
}



function setupTierAnalysis() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(ANALYSIS_SHEET_NAME);

  // Basic layout
  sh.clear();
  sh.getRange("A:H").setFontFamily("Mukta").setFontSize(10);
  sh.getRange("A1").setValue("Tier:");
  sh.getRange("A2").setValue("Level:");
  sh.getRange(TIER_CELL).setNote("Pick a tier sheet");
  sh.getRange(LEVEL_CELL).setNote("Pick a level header from that tier sheet");

  sh.getRange("A3:C3").setValues([["Player", "Opinion", "Reliability"]]);

  // Populate tier dropdown from sheet names (each tier is its own sheet)
  const tierSheetNames = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== ANALYSIS_SHEET_NAME);

  const dvTier = SpreadsheetApp.newDataValidation()
    .requireValueInList(tierSheetNames, true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(TIER_CELL).setDataValidation(dvTier);

  // Clear level validation until tier selected
  sh.getRange(LEVEL_CELL).clearDataValidations();

  sh.setFrozenRows(3);
  
  formatAnalysisSheetLayout_(sh);
}

/*************** DYNAMIC DROPDOWN (LEVELS) ***************/
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== ANALYSIS_SHEET_NAME) return;

  const a1 = range.getA1Notation();
  if (a1 === TIER_CELL) {
    refreshLevelDropdown_();
    // also clear pasted data/output whenever tier changes
    clearAnalysisArea_();
  }

  if (a1 === LEVEL_CELL) {
    clearAnalysisArea_();
    populateSelectedLevel();
  }
}

function formatAnalysisSheetLayout_(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange("A:H").setFontFamily("Mukta").setFontSize(10);

  // Column widths
  sh.setColumnWidths(1, 1, 140); // A
  sh.setColumnWidths(2, 2, 260); // B:C
  sh.setColumnWidths(4, 1, 18);  // D spacer
  sh.setColumnWidths(5, 1, 150); // E output
  sh.setColumnWidths(6, 2, 110); // F:G output
  sh.setColumnWidths(8, 1, 100); // H spark

  // Inputs
  sh.getRange("A1:A2")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  sh.getRange("B1:C2")
    .setBackground("#f3f3f3")
    .setHorizontalAlignment("left");

  // Opinions header
  sh.getRange("A3:C3")
    .setFontWeight("bold")
    .setBackground("#e6f0df")
    .setHorizontalAlignment("center");

  // Output header
  sh.getRange(1, OUTPUT_COL, 1, OUTPUT_WIDTH)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  // ONE thick divider between table and results
  sh.getRange("D1:D200")
    .setBackground("#e0e0e0");

  sh.setFrozenRows(3);

  sh.getRange("D1:D200")
  .setBorder(false, true, false, false, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID_THICK);

  //sh.getRange("F2:G2")
  //.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

}


function formatAnalysisOutput_(
  tool,
  outRowCount,
  orderedTierNames,
  weightsByTier,
  topTier,
  runnerTier,
  splitLowTier,
  splitHighTier,
  passesMajority,
  passesSplitMajority,
  verdictDiffersFromCurrent,
  isPending,
  fuckPresent,
  sd,
  fuckWeight,
  toppct,
  fuckpct,
  allWeight,
  currentTier,
  verdictTier,
  verdictTierName,
  totalWeightedOpinions
) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  // Reset styling first (important)
  tool.getRange(r0, c0, outRowCount, OUTPUT_WIDTH)
    .setBackground(null)
    .setFontWeight("normal")
    .setFontColor("#000000");

  // Box is E1:F3 (labels + values)
  //const box = tool.getRange(r0, c0 + 1, 2, 1);


  // Clear previous borders just for this box
  //box.setBorder(false, false, false, false, false, false);

  // Outer border
  //box.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);


  // Labels column
  tool.getRange(r0, c0, outRowCount, 1)
    .setFontWeight("bold");

  // Right-align numbers
  tool.getRange(r0, c0 + 2, outRowCount, 2)
    .setHorizontalAlignment("right");

  // Center header & increase font size
  tool.getRange(r0, c0, 2, 2)
    .setHorizontalAlignment("center")
    .setFontSize(11);
  // Find key rows
  const labels = tool.getRange(r0, c0, outRowCount, 1)
    .getDisplayValues()
    .flat();
  applyTierSheetColor_(tool, r0, c0, outRowCount);

  const meetsRow = labels.findIndex(v => v === "Place/Move");
  const splitRow = labels.findIndex(v => v === "Split");
  const distHeader = labels.findIndex(v => v === "Tier");

  // !!! All header formatting !!!
  const tierRow = labels.findIndex(v => String(v).trim() === "Tier sheet");
  const levelRow = labels.findIndex(v => String(v).trim() === "Level");

  // Tier Header formatting
  if (tierRow >= 0) {
    const r = r0 + tierRow;
    const badge = tool.getRange(r, c0 + 1, 1, 3); // F:H
    if (!badge.isPartOfMerge()) badge.mergeAcross();
    badge
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
    tool.getRange(r, c0).setBackground(null); // keep label cell neutral
  }

  // Header formatting
  if (levelRow >= 0) {
    const r = r0 + levelRow;
    const pill = tool.getRange(r, c0 + 1, 1, 3); // F:H
    if (!pill.isPartOfMerge()) pill.mergeAcross();
    pill
      .setBackground("#f3f3f3")
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
  }

  const idxTotal = labels.findIndex(v => String(v).trim() === "Total weighted opinions");
  const idxMost  = labels.findIndex(v => String(v).trim() === "Most votes (weighted)");
  const idxRun   = labels.findIndex(v => String(v).trim() === "Runner-up (weighted)");

  const styleMetricRow = (r, bg) => {
    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground(bg)
      .setBorder(false, false, false, false, false, false);

    tool.getRange(r, c0 + 1).setHorizontalAlignment("left");  // tier/text (F)
    tool.getRange(r, c0 + 2, 1, 2).setHorizontalAlignment("right"); // numbers (G:H)

    tool.getRange(r, c0).setFontWeight("bold"); // label emphasis
  };

  if (idxTotal >= 0) {
    const r = r0 + idxTotal;
    styleMetricRow(r, "#ffffff");
    tool.getRange(r, c0 + 1).setFontWeight("bold"); // the total value
    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBorder(false, false, true, false, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID); // thin bottom rule
  }

  if (idxMost >= 0) styleMetricRow(r0 + idxMost, "#ffffff");
  if (idxRun  >= 0) styleMetricRow(r0 + idxRun,  "#ffffff");

  tool.getRange(r0 + idxMost, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun + 1, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left")

  tool.getRange(r0 + idxRun + 2, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left")

  tool.getRange(r0 + idxRun + 2, c0, 1, OUTPUT_WIDTH)
    .setBorder(false,false,true,false,false,false,"#999999", SpreadsheetApp.BorderStyle.DOTTED);

  //tool.getRange(r0 + idxRun + 2, c0, 1, OUTPUT_WIDTH)
    //.setBorder(false,false,true,false,false,false,"#999999", SpreadsheetApp.BorderStyle.DOTTED);

  // !!! Verdict emphasis !!!
  if (meetsRow >= 0) {
    const r = r0 + meetsRow;
    const val = tool.getRange(r, c0 + 1).getDisplayValue().toUpperCase();

    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground(val === "YES" ? "#e6f4ea" : "#fce8e6");

      if (val === "YES") {
        const verdictCell = tool.getRange(r, c0 + 3);
        //const nameToColor = buildTierNameToColor_();

        verdictCell
          //.setBackground(nameToColor[verdictTierName] || "#e0e0e0")
          //.setFontColor(tierTextColor_(verdictTier))
          .setHorizontalAlignment("center")
          .setFontWeight("bold");
      }
    tool.getRange(r, c0 + 1).setFontWeight("bold");

    const msgCell = tool.getRange(r, c0 + 3);
    msgCell.setFontWeight("bold");

    if (val !== "YES") {
      msgCell.setValue("");
      // Decide whether we are in "fuck mode"
      let fuckMode = !!(fuckPresent && toppct < 0.4);

      // In fuck mode, try the fuck triggers first.
      // If neither trigger is met, fall back to regular logic.
      let ruleA = false, ruleB = false;
      if (fuckMode) {
        ruleA = (fuckpct >= 0.2);     // fuck share trigger (only valid if toppct < 0.4 already)
        ruleB = (sd >= 2);            // volatility trigger

        if (!ruleA && !ruleB) fuckMode = false;
      }

      if (fuckMode) {
        if (isPending && !passesMajority) {
          msgCell.setValue("Needs more opinions (F)");
        } else if (!verdictDiffersFromCurrent) {
          msgCell.setValue("No movement necessary (F)");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#efefef");
        } else if (!passesSplitMajority && ruleB) {
          // only relevant when using the SD trigger pathway
          msgCell.setValue("Split not decisive (F)");
        } else {
          msgCell.setValue("Rules not met (F)");
        }
      } else {
        // Regular logic
        if ((isPending && !passesMajority) || (totalWeightedOpinions < 5 || !passesMajority)) {
          msgCell.setValue("Needs more opinions");
        } else if (!passesSplitMajority) {
          msgCell.setValue("Split not decisive");
        } else if (!verdictDiffersFromCurrent) {
          msgCell.setValue("No movement necessary");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#efefef");
        } else {
          msgCell.setValue("Rules not met");
        }
      }
    }
  }

  // !!! Split emphasis !!!
  if (splitRow >= 0) {
    const r = r0 + splitRow;
    const nameToColor = buildTierNameToColor_();

    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground("#fff4cc");

    if (nameToColor[splitLowTier]) {
      tool.getRange(r, c0 + 1)
        .setBackground(nameToColor[splitLowTier])
        .setFontColor(tierTextColor_(splitLowTier))
        .setHorizontalAlignment("right");
    }
    if (nameToColor[splitHighTier]) {
      tool.getRange(r, c0 + 2)
        .setBackground(nameToColor[splitHighTier])
        .setFontColor(tierTextColor_(splitHighTier))
        .setHorizontalAlignment("left");
    }

    tool.getRange(r, c0 + 3)
      .setFontWeight("bold")
      .setNumberFormat("0.00")
      .setHorizontalAlignment("center");
  }

  // !!! Distribution table !!!
  if (distHeader >= 0) {
    const headerRow = r0 + distHeader;
    const firstData = headerRow + 1;

    // Header row subtle
    tool.getRange(headerRow, c0, 1, OUTPUT_WIDTH)
      .setBackground("#eeeeee")
      .setFontWeight("bold");

    const nameToColor = buildTierNameToColor_();

    for (let i = 0; i < orderedTierNames.length; i++) {
      const row = firstData + i;
      const tier = orderedTierNames[i];
      const w = weightsByTier[tier] || 0;

      if (w === 0) {
        tool.getRange(row, c0, 1, OUTPUT_WIDTH)
          .setFontColor("#9e9e9e");
        continue;
      }

      if (nameToColor[tier]) {
        tool.getRange(row, c0)
          .setBackground(nameToColor[tier])
          .setFontWeight("bold")
          .setFontColor(tierTextColor_(tier));
      }

      const sparkCol = c0 + 3; // sparkline H column

      const sparkCell = tool.getRange(row, sparkCol);

      const color1 = nameToColor[tier] || "#999999";

      const firstDataRow = firstData;
      const lastDataRow  = firstData + orderedTierNames.length - 1;

      sparkCell.setFormulaR1C1(
        `=SPARKLINE({RC[-1]/MAX(R${firstDataRow}C[-1]:R${lastDataRow}C[-1]),1},` +
        `{"charttype","bar";"color1","${color1}";"color2","white";"max",1})`
      );
    }

    // Emphasize top + runner only
    const tIdx = orderedTierNames.indexOf(topTier);
    const rIdx = orderedTierNames.indexOf(runnerTier);

    if (tIdx >= 0) {
      tool.getRange(firstData + tIdx, c0, 1, OUTPUT_WIDTH)
        .setFontWeight("bold")
        .setFontLine("underline");
    }
    if (rIdx >= 0) {
      tool.getRange(firstData + rIdx, c0, 1, OUTPUT_WIDTH)
        .setFontStyle("italic");
    }
  }
  if(fuckPresent) {
    tool.getRange(r0 + 10, c0, 1, OUTPUT_WIDTH).setBorder(true,false,false,false,false,false,"#999999", SpreadsheetApp.BorderStyle.DASHED)
    tool.getRange(r0 + 10, c0 +1, 1, 1).setHorizontalAlignment("left").setNumberFormat("0.00");
    tool.getRange(r0 + 11, c0 +1, 1, 1).setNumberFormat("0.0%").setHorizontalAlignment("left");
  }
}



function refreshLevelDropdown_() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  if (!tierName) return;

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) return;

  const headers = getLevelHeaders_(tierSheet); // array of {name, col}
  const headerNames = headers.map(h => h.name);

  const dvLevel = SpreadsheetApp.newDataValidation()
    .requireValueInList(headerNames, true)
    .setAllowInvalid(false)
    .build();

  tool.getRange(LEVEL_CELL).setDataValidation(dvLevel);
}

function getLevelHeaders_(tierSheet) {
  // level names are headers on row 1
  // each level occupies 3 columns: [player, opinion, reliability].
  const headerRow = 1;
  const lastCol = tierSheet.getLastColumn();
  const values = tierSheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  const headers = [];
  for (let c = 1; c <= lastCol; c++) {
    const name = (values[c - 1] || "").trim();
    if (!name) continue;

    // treat any non-empty header cell as a level name
    headers.push({name, col: c});
  }

  // Remove duplicates in case merged headers repeat
  const out = [];
  const seen = new Set();
  for (const h of headers) {
    if (seen.has(h.name)) {
      continue;
    }
    seen.add(h.name);
    out.push(h);
  }
  return out;
}

/*************** POPULATE A:C FOR SELECTED LEVEL ***************/
function clearAnalysisArea_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ANALYSIS_SHEET_NAME);
  const maxRows = sh.getMaxRows();
  const height = maxRows - (DATA_START_ROW - 1);
  if (height > 0) {
    sh.getRange(DATA_START_ROW, 1, height, 3).clearContent().clearFormat().clearNote();
  }
  sh.getRange(OUTPUT_START_ROW, OUTPUT_COL, sh.getMaxRows(), OUTPUT_WIDTH).clearContent().clearFormat().clearNote();
  formatAnalysisSheetLayout_(sh);
}

function safeCellValue_(v) {
  if (v == null) return "";
  // "Image in cell" can come through as a CellImage object
  if (typeof v === "object") {
    try {
      if (typeof v.getUrl === "function") return ""; // CellImage
    } catch (e) {}
    return ""; // any other non-primitive; we only want alphanumeric characters
  }
  return v;
}

function populateSelectedLevel() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) return;

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) return;

  const headers = getLevelHeaders_(tierSheet);
  const header = headers.find(h => h.name === levelName);
  if (!header) return;

  const startCol = header.col;
  const numCols = 3; // player/opinion/reliability
  const lastRow = tierSheet.getLastRow();
  const numRows = Math.max(0, lastRow - 1);

  if (numRows === 0) return;

  const srcRange = tierSheet.getRange(2, startCol, numRows, numCols);
  const vals = srcRange.getValues();
  const bgs = srcRange.getBackgrounds();
  const fcs = srcRange.getFontColors();

  // Filter blank rows (no player and no opinion and no reliability)
  const outVals = [];
  const outBgs = [];
  const outFcs = [];
  for (let r = 0; r < vals.length; r++) {
  const row = vals[r];

  const v0 = safeCellValue_(row[0]);
  const v1 = safeCellValue_(row[1]);
  const v2 = safeCellValue_(row[2]);

  if (String(v0).trim() === "" && String(v1).trim() === "" && String(v2).trim() === "") continue;

  outVals.push([v0, v1, v2]);
  outBgs.push(bgs[r]);
  outFcs.push(fcs[r]);
}

  if (outVals.length === 0) return;

  const dest = tool.getRange(DATA_START_ROW, 1, outVals.length, 3);
  dest.setValues(outVals);
  dest.setBackgrounds(outBgs);
  dest.setFontColors(outFcs);
}


const tierNameFromTextColor = (fontHex) => {
  const h = hex_(fontHex);
  // If the font color matches one of the tier fill colors, treat it as specifying that tier
  return difficultyColorNames[h] || null;
};

function hasInsaneTokenNotDemon_(text) {
  const words = String(text || "")
    .toLowerCase()
    // turn anything that isn't a letter into spaces
    .replace(/[^a-z]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  for (let i = 0; i < words.length; i++) {
    if (words[i] === "insane") {
      // if the next token is "demon", it refers to "insane demon"
      if (words[i + 1] === "demon") {
        return false;
      }
      return true;
    }
  }
  return false;
}

function mean_(xs) {
  if (!xs || xs.length === 0) {
    return 0;
  }
  let s = 0;
  for (let i = 0; i < xs.length; i++) {
    s += xs[i];
  }
  return s / xs.length;
}

function median_(xs) {
  if (!xs || xs.length === 0) {
    return 0;
  }
  const a = xs.slice().sort((x, y) => x - y);
  const n = a.length;
  const mid = Math.floor(n / 2);
  return (n % 2) ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

// Turns an index into tier label 
// ex: 4.0 becomes "Hard"
function tierLabelFromIndex_(orderedTierNames, idxFloat) {
  const n = orderedTierNames.length;
  if (!n) return "";

  const x = Number(idxFloat);
  if (!isFinite(x)) return "";

  const i = clamp_(Math.round(x), 0, n - 1);
  return orderedTierNames[i] || "";
}

// Outlier rule: farther than k tiers from the median
function outlierPct_(rawPoints, centerIdx, k) {
  const total = rawPoints.length;
  if (!total) {
    return { count: 0, total: 0, pct: 0 };
  }

  const cutoff = Math.abs(k);
  let outliers = 0;

  for (const p of rawPoints) {
    if (Math.abs(p - centerIdx) > cutoff) {
      outliers++;
    }
  }

  return {
    count: outliers,
    total,
    pct: outliers / total
  };
}


function pickSplitLowIdx_(orderedTierNames, weightsByTier, topIdx, runnerIdx, currentIdx, centerIdx, isPending) {
  const n = orderedTierNames.length;

  const sumRange = (a, b) => {
    let s = 0;
    for (let i = a; i <= b; i++) s += (weightsByTier[orderedTierNames[i]] || 0);
    return s;
  };

  const boundaryScore = (lowIdx) => {
    const leftTotal = sumRange(0, lowIdx);
    const rightTotal = sumRange(lowIdx + 1, n - 1);
    return { leftTotal, rightTotal, score: Math.abs(leftTotal - rightTotal) };
  };

  const areAdjacent = (a, b) => a >= 0 && b >= 0 && Math.abs(a - b) === 1;

  // Constrain split to include the current placed tier unless:
  // - is Pending, OR
  // - placed tier is > 2 tiers away from the center (median)
  const currentIsTopOrRunner = (currentIdx === topIdx || currentIdx === runnerIdx);

  const constrain =
    !isPending &&
    currentIdx >= 0 &&
    currentIsTopOrRunner &&     // only anchor if current is top/runner-up
    Math.abs(currentIdx - centerIdx) <= 2;

  if (constrain) {
    const candidates = [];

    // (current-1 | current)
    if (currentIdx > 0) {
      const lowIdx = currentIdx - 1;
      candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
    }

    // (current | current+1)
    if (currentIdx < n - 1) {
      const lowIdx = currentIdx;
      candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
    }

    candidates.sort((a, b) => a.score - b.score);
    return candidates.length ? candidates[0].lowIdx : Math.max(0, Math.min(n - 2, currentIdx));
  }

  // adjacent top/runner forced, else pick boundary adjacent to top that best balances the split
  if (areAdjacent(topIdx, runnerIdx)) return Math.min(topIdx, runnerIdx);

  const candidates = [];

  if (topIdx > 0) {
    const lowIdx = topIdx - 1;
    candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
  }
  if (topIdx < n - 1) {
    const lowIdx = topIdx;
    candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
  }

  candidates.sort((a, b) => a.score - b.score);
  return candidates.length ? candidates[0].lowIdx : Math.max(0, topIdx - 1);
}



/*************** ANALYSIS ***************/
function analyzeSelectedLevel() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) {
    SpreadsheetApp.getUi().alert("Pick a tier and a level first.");
    return;
  }

  // Read populated A:C
  const lastRow = tool.getLastRow();
  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No opinions pasted yet. Pick the level dropdown again.");
    return;
  }

  const numRows = lastRow - (DATA_START_ROW - 1);
  const rng = tool.getRange(DATA_START_ROW, 1, numRows, 3);
  const vals = rng.getValues();
  const bgs = rng.getBackgrounds();
  const fcs = rng.getFontColors();

  // Tier ordering for split logic (low -> high)
  // Build stable tier order from difficultyColorNames keys list
  const orderedTierNames = [
    "Insane Demon", "Beginner","Easy","Medium","Hard","Very Hard","Insane","Extreme","Remorseless","Relentless","Terrifying","Catastrophic","Inexorable","Excruciating","Merciless","Monstrous","Apocalyptic","Demonic","Menacing","Unreal","Nightmare"
  ];

  // Accumulate weighted votes
  const weightsByTier = {};
  for (const n of orderedTierNames) weightsByTier[n] = 0;

  let fuckPresent = false;
  let fuckWeight = 0; // track separately so we can calculate later

  let totalWeight = 0;
  let rawCount = 0;
  const rawPoints = []; // numeric tier indices per raw opinion (splits -> midpoint)

  const tierIdxOf_ = (tierName) => orderedTierNames.indexOf(tierName);


  for (let r = 0; r < vals.length; r++) {
    const player = String(vals[r][0] ?? "").trim();
    const opinionText = String(vals[r][1] ?? "").trim();
    const relText = String(vals[r][2] ?? "").trim();

    if (!player && !opinionText && !relText) continue;

    const opinionColor = hex_(bgs[r][1]);      // column B background
    const relColor = hex_(bgs[r][2]);          // column C background
    const w = reliabilityFactors[relColor] ?? 0;

    const opinionFont = hex_(fcs[r][1]);       // column B font color


    if (w <= 0) continue;

    // !!! Fuck tier opinion handling (supports 3-way splits) !!!
    if (opinionColor === "#000000") {
      fuckPresent = true;

      rawCount++;

      // Always count full Fuck weight for black opinions
      fuckWeight += w;

      // 3-way: black background & font is a split-color (ex: Insane/Extreme)
      const fontKey = hex_(opinionFont);
      if (splitPairs[fontKey]) {
        const [lowHex, highHex] = splitPairs[fontKey].map(hex_);
        const lowTier = difficultyColorNames[lowHex];
        const highTier = difficultyColorNames[highHex];

        const i1 = tierIdxOf_(lowTier);
        const i2 = tierIdxOf_(highTier);
        if (i1 >= 0 && i2 >= 0) {
          rawPoints.push((i1 + i2) / 2);
        }

        if (lowTier && highTier) {
          const halfW = w / 3;
          weightsByTier[lowTier] = (weightsByTier[lowTier] || 0) + halfW;
          weightsByTier[highTier] = (weightsByTier[highTier] || 0) + halfW;
          fuckWeight -= halfW*2;
          totalWeight += halfW*2;
        }
        continue;
      }

      // 2-way: black background & font is a real tier color (ex: Very Hard)
      const fontKey2 = hex_(opinionFont);
      const fontTier = difficultyColorNames[fontKey2]; // tier name if font matches a tier color

      if (fontTier && fontTier !== "Insane") {

        const it = tierIdxOf_(fontTier);
        if (it >= 0) {
          rawPoints.push(it);
        }

        const halfW = w / 2;

        fuckWeight -= halfW; 
        totalWeight += halfW;       

        weightsByTier[fontTier] = (weightsByTier[fontTier] || 0) + halfW;
        continue;
      }

      // 2-way special case: ONLY split Insane/Fuck if the text contains "insane"
      const hasInsane = String(opinionText).toLowerCase().includes("insane");
      if (hasInsane) {

        const it = tierIdxOf_("Insane");
        if (it >= 0) {
          rawPoints.push(it);
        }

        const halfW = w / 2;
        fuckWeight -= halfW;
        totalWeight += halfW;
        weightsByTier["Insane"] = (weightsByTier["Insane"] || 0) + halfW;
      }

      continue;
    }


    // !!! Normal single-tier opinion (using background color) !!!
    if (difficultyColorNames[opinionColor]) {

      const tName = difficultyColorNames[opinionColor];
      const it = tierIdxOf_(tName);
      if (it >= 0) {
        rawPoints.push(it);
      }
      weightsByTier[tName] = (weightsByTier[tName] || 0) + w;
      rawCount++;
      totalWeight+= w;
      continue;
    }

    // !! Normal split (non-fuck) by splitPairs !!!
    const oc = hex_(opinionColor); // normalize lookup key
    if (splitPairs[oc]) {
      const [c1Raw, c2Raw] = splitPairs[oc];
      const c1 = hex_(c1Raw);
      const c2 = hex_(c2Raw);

      const t1 = difficultyColorNames[c1];
      const t2 = difficultyColorNames[c2];

      const i1 = tierIdxOf_(t1);
      const i2 = tierIdxOf_(t2);
      if (i1 >= 0 && i2 >= 0) {
        rawPoints.push((i1 + i2) / 2);
      }

      // regular split 50/50
      if (t1 && t2) {
        const halfW = w / 2;
        weightsByTier[t1] = (weightsByTier[t1] || 0) + halfW;
        weightsByTier[t2] = (weightsByTier[t2] || 0) + halfW;
        rawCount++;
        totalWeight+= w;
        continue;
      }

    // anything else ignore here (handled elsewhere)
    }



    // Unknown opinion color -> also ignore
  }

  function outlierCutoffLabels_(orderedTierNames, centerIdx, k) {
    const n = orderedTierNames.length;

    // clamps values; does not round
    const clampF = (x, lo, hi) => Math.max(lo, Math.min(hi, x));

    const labelAtGrid = (g) => {
      const gi = clampF(g, 0, n - 1);
      const lo = Math.floor(gi);
      const frac = gi - lo;

      if (Math.abs(frac) < 1e-9) return orderedTierNames[lo] || "";

      const hi = Math.min(lo + 1, n - 1);
      return `${orderedTierNames[lo]} / ${orderedTierNames[hi]}`;
    };

    // Raw numeric bounds from MEAN
    // const lowRaw = centerIdx - k;
    // const highRaw = centerIdx + k;

    const lowBoundary  = Math.floor((centerIdx - k) * 2) / 2; // down to nearest 0.5
    const highBoundary = Math.ceil ((centerIdx + k) * 2) / 2; // up to nearest 0.5

    return {
      lowBoundary,
      highBoundary,
      lowLabel: labelAtGrid(lowBoundary),
      highLabel: labelAtGrid(highBoundary),
    };
  }

  
  const rawMeanIdx = mean_(rawPoints);
  const rawMedianIdx = median_(rawPoints);

  const rawMeanLabel = tierLabelFromIndex_(orderedTierNames, rawMeanIdx);
  const rawMedianLabel = tierLabelFromIndex_(orderedTierNames, rawMedianIdx);


  // Find majority tier by highest weighted total
  const entries = Object.entries(weightsByTier).sort((a, b) => b[1] - a[1]);
  const top = entries[0] || ["(none)", 0];
  const second = entries[1] || ["(none)", 0];

  const allWeight = totalWeight + fuckWeight;

  const topTier = top[0];
  const topWeight = top[1];
  const secondTier = second[0];
  const secondWeight = second[1];

  // Majority rule (weighted)
  const passesMajority = topWeight >= 4; // Must have at least 4 opinions (by weight)
  const topVsSecond = (topWeight - secondWeight) >= 3; // Must be leading by at least 3 points

  // Weighted SD over tier indices (orderedTierNames)
  const idx = {}; orderedTierNames.forEach((t,i)=>idx[t]=i);
  let sw=0, mu=0;
  orderedTierNames.forEach(t=>{ const wt=weightsByTier[t]||0; sw+=wt; mu+=wt*idx[t]; });
  mu = sw ? mu/sw : 0;
  let v=0;
  orderedTierNames.forEach(t=>{ const wt=weightsByTier[t]||0, d=idx[t]-mu; v+=wt*d*d; });
  const sd = sw ? Math.sqrt(v/sw) : 0;

  let OUTLIER_K = 1.5;
  if (sd > 1.5) {
    OUTLIER_K = 2.5;
  } else if (sd > 1) {
    OUTLIER_K = 2;
  }
  const outlierPctObj = outlierPct_(rawPoints, rawMeanIdx, OUTLIER_K); // >K tiers from mean
  const outlierBounds = outlierCutoffLabels_(orderedTierNames, rawMeanIdx, OUTLIER_K);

  
  // !!! SPLIT LOGIC (boundary between adjacent tiers) !!!
  // Rule:
  // - Split is ALWAYS a boundary between two adjacent tiers A | B
  // - Interpreted as: (everything <= A) versus (everything >= B)
  // - The highest-weighted tier (topTier) is always part of the split
  // - If runner-up is adjacent to topTier, split MUST be between them
  // - Otherwise, choose the boundary adjacent to topTier that best balances totals (above or below)

  const idxOf = (name) => orderedTierNames.indexOf(name);

  const sumRange = (startIdx, endIdx) => {
    // inclusive indices
    let s = 0;
    for (let i = startIdx; i <= endIdx; i++) s += (weightsByTier[orderedTierNames[i]] || 0);
    return s;
  };

  const totalW = totalWeight;

  const topIdx = idxOf(topTier);
  const runnerIdx = idxOf(secondTier);

  const areAdjacent = (a, b) => a >= 0 && b >= 0 && Math.abs(a - b) === 1;

  // Candidate boundary is represented by (lowIdx, highIdx) where highIdx = lowIdx + 1
  // and the meaning is: leftTotal = sum(0..lowIdx), rightTotal = sum(highIdx..end)
  const boundaryTotals = (lowIdx) => {
    const leftTotal = sumRange(0, lowIdx);
    const rightTotal = sumRange(lowIdx + 1, orderedTierNames.length - 1);
    return { leftTotal, rightTotal };
  };

  const currentTier = String(tierName || "").trim();
  const currentIdx = orderedTierNames.indexOf(currentTier);
  const currentTierLower = currentTier.toLowerCase();
  const isPending = (currentTierLower === "pending");

  let splitLowIdx = pickSplitLowIdx_(
  orderedTierNames,
  weightsByTier,
  topIdx,
  runnerIdx,
  currentIdx,
  rawMedianIdx,   // your center
  isPending
  );

  const splitHighIdx = splitLowIdx + 1;

  const splitLowTier = orderedTierNames[splitLowIdx];
  const splitHighTier = orderedTierNames[splitHighIdx];

  // IMPORTANT: these are SIDE TOTALS
  // everything <= splitLowTier  vs  everything >= splitHighTier
  const splitLeftTotal = sumRange(0, splitLowIdx);
  const splitRightTotal = sumRange(splitHighIdx, orderedTierNames.length - 1);

  const toppct = allWeight > 0 ? topWeight/allWeight : 0;

  let vtn = "";
  if (fuckPresent) {
    if (((allWeight > 0 ? fuckWeight/allWeight : 0) >= 0.2 && toppct < 0.4) || (sd >= 2 && passesMajority)) {
      vtn = "fuck";
    } else {
      vtn = (splitLeftTotal > splitRightTotal) ? splitLowTier : splitHighTier;
    }
  } else {
    vtn = (splitLeftTotal > splitRightTotal) ? splitLowTier : splitHighTier;
  }

  const verdictTier = String(vtn).trim().toLowerCase();

  const verdictDiffersFromCurrent = verdictTier !== "" && verdictTier !== currentTierLower;

  const splitThreshold = isPending ? 3 : 4;
  const passesSplitMajority = Math.abs(splitLeftTotal - splitRightTotal) >= splitThreshold;


  // Output
  let out = [];
  out.push([`Tier sheet`, tierName, "", ""]);
  out.push([`Level`, levelName, "", ""]);
  out.push([`Total weighted opinions`, allWeight, "", ""]);
  out.push([`Most votes (weighted)`, topTier, topWeight, ""]);
  out.push([`Runner-up (weighted)`, secondTier, secondWeight, ""]);
  out.push([`Tier Mean`, rawMeanLabel, rawMeanIdx, ""]);
  out.push([`Tier Median`, rawMedianLabel, rawMedianIdx, ""]);
  out.push([`Outliers`, `${outlierPctObj.count} / ${outlierPctObj.total} (${(outlierPctObj.pct * 100).toFixed(1)}%)`, "", ""]);
  out.push([
    "Outlier range",
    `≤ ${outlierBounds.lowLabel}` + `       ≥ ${outlierBounds.highLabel}`,
    "",
    ""
  ]);
  out.push([`Standard Deviation`, sd, "", ""]);
  if (fuckPresent) { out.push([`Fuck weight`, fuckWeight, "", ""]);
                  out.push([`Fuck % of all`, (allWeight) > 0 ? fuckWeight / (allWeight) : 0, "", ""]); }
  
  let verdictTierName = verdictTier; 
  if (fuckPresent) {
    if (((((allWeight) > 0 ? fuckWeight / (allWeight) : 0) >= 0.2) && toppct < 0.4)
            || (sd >= 2 && passesMajority)) {
      verdictTierName = "Fuck";
    }
  } else {
    const verdictIdx = (splitLeftTotal > splitRightTotal) ? splitLowIdx : splitHighIdx;
    const verdictBaseName = orderedTierNames[verdictIdx];

    if (isPending) {
      verdictTierName = verdictBaseName;
    } else {
      // Compare to current tier to choose arrow direction
      let arrow = "";
      if (currentIdx >= 0) {
        if (verdictIdx < currentIdx) {
          arrow = "↓    ";
        } else if (verdictIdx > currentIdx) {
          arrow = "↑    ";
        }
      }
      verdictTierName = arrow + verdictBaseName;
}
  }

  canMove = false;
  if (fuckPresent) {
    if (((((allWeight) > 0 ? fuckWeight / (allWeight) : 0) >= 0.2 && toppct < 0.4) && verdictDiffersFromCurrent) 
            || (sd >= 2 && ((passesMajority || (topVsSecond && !isPending)) && passesSplitMajority && verdictDiffersFromCurrent))) {
      canMove = true;
    } 
  } else {
      if ((passesMajority || (topVsSecond && !isPending)) && passesSplitMajority && verdictDiffersFromCurrent) {
        canMove = true;
      }
  }
  

   out.push([`Place/Move`, canMove ? "YES" : "NO", "", canMove ? verdictTierName : ""]);
    // Weighted median split displayed as: LowTier cumulative | HighTier cumulative
  out.push([`Split`,
    `${splitLowTier}`,
    `${splitHighTier}`,
    `${splitLeftTotal.toFixed(2)} | ${splitRightTotal.toFixed(2)}`
  ]);


  // Also dump full distribution
  out.push(["", "", "", ""]);
  out.push(["Tier", "Weighted", "%", ""]);
  for (const t of orderedTierNames) {
    const w = weightsByTier[t] || 0;
    const denom = allWeight > 0 ? allWeight : 1;
    const share = w / denom;
    out.push([t, w, share, ""]);
  }


  const startRow = 1; const startCol = OUTPUT_COL; // E
  tool.getRange(startRow, startCol, out.length, OUTPUT_WIDTH).setValues(out);

  tool.getRange(startRow, startCol + 1, out.length, 1)
  .setNumberFormat("0.###");

  const sdRow = out.findIndex(r => r[0] === "Standard Deviation");
  if (sdRow >= 0) tool.getRange(startRow + sdRow, startCol + 1).setNumberFormat("0.00").setHorizontalAlignment("left");

  const outRow = out.findIndex(r => r[0] === "Outlier %");
  if (outRow >= 0) {
    tool.getRange(startRow + outRow, startCol + 1).setNumberFormat("0.0%").setHorizontalAlignment("left");
    tool.getRange(startRow + outRow + 1, startCol + 3).setHorizontalAlignment("right");
  }


  // Format share column as percent where applicable
  tool.getRange(startRow + 9, startCol + 2, Math.max(0, out.length - 9), 1).setNumberFormat("0.0%");
  const fuckpct = allWeight > 0 ? fuckWeight / allWeight : 0;

  formatAnalysisOutput_(
    tool,
    out.length,
    orderedTierNames,
    weightsByTier,
    topTier,
    secondTier,
    splitLowTier,
    splitHighTier,
    passesMajority,
    passesSplitMajority,
    verdictDiffersFromCurrent,
    isPending,
    fuckPresent,
    sd,
    fuckWeight,
    toppct,
    fuckpct,
    allWeight,
    currentTier,
    verdictTier,
    verdictTierName,
    allWeight
  );

}
