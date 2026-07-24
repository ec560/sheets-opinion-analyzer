// Paste-first tier configuration. Tier key goes in Column A with names, order, and colors:
// tier rows alternate with "Lower/Higher" split rows.

const TIER_CONFIG_HEADER_ROW = 4;
const TIER_CONFIG_START_ROW = 5;
const TIER_CONFIG_LIST_COL = 1;
const TIER_CONFIG_BOTTOM_CELL = "E5";
const TIER_CONFIG_TOP_CELL = "E6";
let tierConfigurationLoaded_ = false;
let tierConfigurationResult_ = null;

function defaultTierConfigRows_() {
  const colorsByName = {};
  Object.keys(DEFAULT_DIFFICULTY_COLOR_NAMES).forEach(color => {
    const name = DEFAULT_DIFFICULTY_COLOR_NAMES[color];
    if (!(name in colorsByName)) colorsByName[name] = hex_(color);
  });

  const borderByLowerTier = {};
  Object.keys(DEFAULT_SPLIT_PAIRS).forEach(borderColor => {
    const lowerName = DEFAULT_DIFFICULTY_COLOR_NAMES[hex_(DEFAULT_SPLIT_PAIRS[borderColor][0])];
    if (lowerName) borderByLowerTier[lowerName] = hex_(borderColor);
  });

  return DEFAULT_ORDERED_TIER_NAMES.map((name, index) => ({
    order: index + 1,
    name,
    primaryColor: colorsByName[name] || "#ffffff",
    alternateColors: [],
    borderColor: borderByLowerTier[name] || "",
    group: name === "Insane Demon" ? "Bottom" : "Own",
    fontColor: tierConfigurationTextColor_(name),
    splitLabel: index < DEFAULT_ORDERED_TIER_NAMES.length - 1
      ? name + "/" + DEFAULT_ORDERED_TIER_NAMES[index + 1]
      : "",
    splitFontColor: index === 0 ? "#ffffff" : "#000000"
  }));
}

function tierRowsToPasteList_(rows) {
  const values = [];
  const backgrounds = [];
  const fontColors = [];
  rows.forEach((row, index) => {
    values.push([row.name]);
    backgrounds.push([row.primaryColor]);
    fontColors.push([row.fontColor || "#000000"]);
    if (index < rows.length - 1) {
      values.push([row.splitLabel || row.name + "/" + rows[index + 1].name]);
      backgrounds.push([row.borderColor]);
      fontColors.push([row.splitFontColor || "#000000"]);
    }
  });
  return { values, backgrounds, fontColors };
}

function setupTierConfiguration_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TIER_CONFIG_SHEET_NAME);
  if (sh && String(sh.getRange("A4").getDisplayValue()).trim() === "Paste tier list") {
    return sh;
  }
  if (!sh) sh = ss.insertSheet(TIER_CONFIG_SHEET_NAME);

  const rows = defaultTierConfigRows_();
  const pasted = tierRowsToPasteList_(rows);
  sh.getDataRange().breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();

  sh.getRange("A1:E1").mergeAcross()
    .setValue("Tier Configuration")
    .setFontFamily("Mukta")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground("#434343")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("left");
  sh.getRange("A2:E2").mergeAcross()
    .setValue("Paste one formatted column below. Tier rows set order and color; any row containing \"/\" marks an optional split.")
    .setFontFamily("Mukta")
    .setFontColor("#4f5b66");
  sh.getRange("A3:E3").mergeAcross()
    .setValue("Use the two optional edge lists only when several outer tiers should share one placement tier.")
    .setFontFamily("Mukta")
    .setFontColor("#6b6257")
    .setBackground("#fff8e7");

  sh.getRange("A4").setValue("Paste tier list")
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setBackground("#e9edf1");
  sh.getRange("D4:E4").setValues([["Edge grouping", "Tier names"]])
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setBackground("#e9edf1");
  sh.getRange("D5:D6").setValues([["Count toward bottom"], ["Count toward top"]])
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setBackground("#f6f7f8");
  sh.getRange(TIER_CONFIG_BOTTOM_CELL).setValue("Insane Demon")
    .setNote("Optional. Separate multiple tier names with commas or newlines.");
  sh.getRange(TIER_CONFIG_TOP_CELL)
    .setNote("Optional. Separate multiple tier names with commas or newlines.");

  sh.getRange(TIER_CONFIG_START_ROW, TIER_CONFIG_LIST_COL, pasted.values.length, 1)
    .setValues(pasted.values)
    .setBackgrounds(pasted.backgrounds)
    .setFontColors(pasted.fontColors)
    .setFontFamily("Mukta")
    .setVerticalAlignment("middle");

  sh.setFrozenRows(TIER_CONFIG_HEADER_ROW);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 245);
  sh.setColumnWidth(2, 24);
  sh.setColumnWidth(3, 24);
  sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 210);
  sh.getRange("E5:E6").setWrap(true).setVerticalAlignment("top");
  return sh;
}

function isTierHexColor_(value) {
  return /^#[0-9a-f]{6}$/.test(String(value || "").trim().toLowerCase());
}

function parseTierNameList_(value) {
  return String(value || "")
    .split(/[\n,;]+/)
    .map(name => name.trim())
    .filter(Boolean);
}

function isTierSplitLabel_(label) {
  return String(label || "").includes("/");
}

function parseTierConfigList_(values, backgrounds, bottomValue, topValue, fontColors) {
  const errors = [];
  const entries = [];
  values.forEach((row, offset) => {
    const label = String(row[0] || "").trim();
    if (!label) return;
    const color = hex_(backgrounds[offset] && backgrounds[offset][0]);
    const pastedFontColor = hex_(fontColors && fontColors[offset] && fontColors[offset][0]);
    const fontColor = isTierHexColor_(pastedFontColor) ? pastedFontColor : "#000000";
    const sheetRow = TIER_CONFIG_START_ROW + offset;
    entries.push({ label, color, fontColor, sheetRow, isSplit: isTierSplitLabel_(label) });
    if (!isTierHexColor_(color)) errors.push("Row " + sheetRow + ": paste a cell with a solid fill color.");
    if (color === "#000000") errors.push("Row " + sheetRow + ": black is reserved by the analyzer.");
  });

  const tierEntries = entries.filter(entry => !entry.isSplit);
  if (tierEntries.length < 2) errors.push("Paste at least two tier rows.");
  if (entries.length && entries[0].isSplit) errors.push("The pasted list must begin with a tier, not a split.");
  if (entries.length && entries[entries.length - 1].isSplit) errors.push("The pasted list must end with a tier, not a split.");

  const seenNames = {};
  const seenColors = {};
  entries.forEach((entry, index) => {
    if (seenColors[entry.color]) errors.push("Each pasted fill color must be unique (duplicate " + entry.color + ").");
    seenColors[entry.color] = true;
    if (entry.isSplit) {
      const previous = entries[index - 1];
      const next = entries[index + 1];
      if (!previous || !next || previous.isSplit || next.isSplit) {
        errors.push("Row " + entry.sheetRow + ": a split row must sit between two tier rows.");
      }
    } else {
      const key = entry.label.toLowerCase();
      if (key === "fuck") errors.push("Row " + entry.sheetRow + ": Fuck is reserved by the analyzer.");
      if (seenNames[key]) errors.push("Tier names must be unique (duplicate " + entry.label + ").");
      seenNames[key] = entry;
    }
  });

  const bottomNames = parseTierNameList_(bottomValue);
  const topNames = parseTierNameList_(topValue);
  const bottomKeys = {};
  const topKeys = {};
  bottomNames.forEach(name => bottomKeys[name.toLowerCase()] = true);
  topNames.forEach(name => topKeys[name.toLowerCase()] = true);
  bottomNames.concat(topNames).forEach(name => {
    if (!seenNames[name.toLowerCase()]) errors.push("Edge grouping names an unknown tier: " + name + ".");
  });
  Object.keys(bottomKeys).forEach(key => {
    if (topKeys[key]) errors.push((seenNames[key] ? seenNames[key].label : key) + " cannot count toward both bottom and top.");
  });

  const rows = tierEntries.map((entry, index) => {
    const key = entry.label.toLowerCase();
    const entryIndex = entries.indexOf(entry);
    const nextEntry = entries[entryIndex + 1];
    const borderColor = nextEntry && nextEntry.isSplit ? nextEntry.color : "";
    return {
      order: index + 1,
      name: entry.label,
      primaryColor: entry.color,
      alternateColors: [],
      borderColor,
      group: bottomKeys[key] ? "Bottom" : topKeys[key] ? "Top" : "Own",
      fontColor: entry.fontColor,
      splitLabel: nextEntry && nextEntry.isSplit ? nextEntry.label : "",
      splitFontColor: nextEntry && nextEntry.isSplit ? nextEntry.fontColor : "",
      sheetRow: entry.sheetRow
    };
  });

  const ownIndexes = rows.map((row, index) => row.group === "Own" ? index : -1).filter(index => index >= 0);
  if (!ownIndexes.length) {
    errors.push("At least one tier must remain outside the bottom/top groups.");
  } else {
    const firstOwn = ownIndexes[0];
    const lastOwn = ownIndexes[ownIndexes.length - 1];
    rows.forEach((row, index) => {
      if (row.group === "Bottom" && index > firstOwn) errors.push(row.name + " must be before every ungrouped tier to count toward bottom.");
      if (row.group === "Top" && index < lastOwn) errors.push(row.name + " must be after every ungrouped tier to count toward top.");
    });
  }

  return {
    valid: errors.length === 0,
    rows,
    splitCount: entries.filter(entry => entry.isSplit).length,
    errors
  };
}

function readTierConfiguration_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TIER_CONFIG_SHEET_NAME);
  if (!sh) return { valid: false, missing: true, rows: [], errors: [] };
  const lastRow = Math.max(sh.getLastRow(), TIER_CONFIG_START_ROW);
  const count = lastRow - TIER_CONFIG_START_ROW + 1;
  const range = sh.getRange(TIER_CONFIG_START_ROW, TIER_CONFIG_LIST_COL, count, 1);
  return parseTierConfigList_(
    range.getDisplayValues(),
    range.getBackgrounds(),
    sh.getRange(TIER_CONFIG_BOTTOM_CELL).getDisplayValue(),
    sh.getRange(TIER_CONFIG_TOP_CELL).getDisplayValue(),
    range.getFontColors()
  );
}

function applyTierConfiguration_(rows) {
  Object.keys(difficultyColorNames).forEach(key => delete difficultyColorNames[key]);
  Object.keys(splitPairs).forEach(key => delete splitPairs[key]);
  Object.keys(tierDecisionTargets).forEach(key => delete tierDecisionTargets[key]);
  Object.keys(tierFontColors).forEach(key => delete tierFontColors[key]);
  Object.keys(opinionFontColorsByFill).forEach(key => delete opinionFontColorsByFill[key]);
  orderedTierNames.splice(0, orderedTierNames.length);

  rows.forEach(row => {
    orderedTierNames.push(row.name);
    difficultyColorNames[hex_(row.primaryColor)] = row.name;
    tierFontColors[row.name] = hex_(row.fontColor) || "#000000";
    opinionFontColorsByFill[hex_(row.primaryColor)] = hex_(row.fontColor) || "#000000";
    if (row.splitLabel) tierFontColors[row.splitLabel] = hex_(row.splitFontColor) || "#000000";
    if (row.borderColor) {
      opinionFontColorsByFill[hex_(row.borderColor)] = hex_(row.splitFontColor) || "#000000";
    }
  });
  rows.forEach((row, index) => {
    if (index < rows.length - 1 && row.borderColor) {
      splitPairs[hex_(row.borderColor)] = [row.primaryColor, rows[index + 1].primaryColor];
    }
  });

  const ownRows = rows.filter(row => row.group === "Own");
  const bottomTarget = ownRows.length ? ownRows[0].name : rows[0].name;
  const topTarget = ownRows.length ? ownRows[ownRows.length - 1].name : rows[rows.length - 1].name;
  rows.forEach(row => {
    tierDecisionTargets[row.name] = row.group === "Bottom"
      ? bottomTarget
      : row.group === "Top"
        ? topTarget
        : row.name;
  });
}

function resetTierConfigurationToDefaults_() {
  applyTierConfiguration_(defaultTierConfigRows_());
}

function loadTierConfiguration_(force) {
  if (tierConfigurationLoaded_ && !force) return tierConfigurationResult_;
  const result = readTierConfiguration_();
  if (result.valid) applyTierConfiguration_(result.rows);
  else if (result.missing) resetTierConfigurationToDefaults_();
  tierConfigurationLoaded_ = true;
  tierConfigurationResult_ = result;
  return result;
}

function tierConfigurationErrorMessage_(result) {
  if (!result || result.valid || result.missing) return "";
  return "Fix Tier Configuration before analyzing:\n\n" + result.errors.join("\n");
}

function tierConfigurationTextColor_(name) {
  const normalized = String(name || "").trim().toLowerCase();
  return normalized === "insane demon" || normalized === "insane demon/beginner"
    ? "#ffffff"
    : "#000000";
}
