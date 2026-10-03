// Paste-first tier configuration. Tier key goes in Column A with names, order, and colors:
// tier rows alternate with "Lower/Higher" split rows.

const TIER_CONFIG_HEADER_ROW = 4;
const TIER_CONFIG_START_ROW = 5;
const TIER_CONFIG_LIST_COL = 1;
const TIER_CONFIG_REQUIRED_COLUMNS = 5;
const TIER_CONFIG_DECISION_HEADER_CELL = "D4";
const TIER_CONFIG_MINIMUM_OPINIONS_CELL = "E5";
const TIER_CONFIG_LOCK_THRESHOLD_CELL = "E6";
const TIER_CONFIG_LOCK_MINIMUM_RAW_CELL = "E7";
const TIER_CONFIG_LOCK_MINIMUM_WEIGHTED_CELL = "E8";
const TIER_CONFIG_SPLIT_START_ROW = 12;
const TIER_CONFIG_MAX_SPLIT_ROW_COUNT = 6;
const TIER_CONFIG_SPLIT_START_COL = 4;
const TIER_CONFIG_SECTION_GAP_ROWS = 1;
const TIER_CONFIG_EDGE_HEADER = "Edge grouping";
const TIER_CONFIG_PANEL_END_ROW = TIER_CONFIG_SPLIT_START_ROW +
  TIER_CONFIG_MAX_SPLIT_ROW_COUNT + TIER_CONFIG_SECTION_GAP_ROWS + 2;
const TIER_CONFIG_LEGACY_BOTTOM_CELL = "E5";
const TIER_CONFIG_LEGACY_TOP_CELL = "E6";
const TIER_CONFIG_CACHE_SCHEMA = 2;
const TIER_CONFIG_CACHE_KEY = "analyzer:tier-configuration:v" + TIER_CONFIG_CACHE_SCHEMA;
const TIER_CONFIG_CACHE_TTL_SECONDS = 21600;
let tierConfigurationLoaded_ = false;
let tierConfigurationResult_ = null;
let tierConfigurationLoadSource_ = "not-loaded";
let tierConfigurationCacheStatus_ = "not-checked";

function getTierConfigurationCache_() {
  if (typeof CacheService === "undefined" || !CacheService.getDocumentCache) return null;
  try {
    return CacheService.getDocumentCache();
  } catch (error) {
    return null;
  }
}

function invalidateTierConfigurationCache_() {
  tierConfigurationLoaded_ = false;
  tierConfigurationResult_ = null;
  const cache = getTierConfigurationCache_();
  if (!cache) return;
  try {
    cache.remove(TIER_CONFIG_CACHE_KEY);
  } catch (error) {}
}

function tierConfigurationNumber_(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function parseTierDecisionConfiguration_(settingsValues, splitValues) {
  const errors = [];
  const settings = Array.isArray(settingsValues) ? settingsValues : [];
  const readSetting = index => tierConfigurationNumber_(
    Array.isArray(settings[index]) ? settings[index][0] : settings[index]
  );
  const minimumOpinions = readSetting(0);
  const lockThreshold = readSetting(1);
  const lockMinimumRawOpinions = readSetting(2);
  const lockMinimumWeightedOpinions = readSetting(3);

  if (minimumOpinions === null || minimumOpinions < 0) {
    errors.push("Minimum opinions must be a number greater than or equal to 0.");
  }
  if (lockThreshold === null || lockThreshold < 0 || lockThreshold > 1) {
    errors.push("Lock threshold must be a percentage from 0% to 100%.");
  }
  if (
    lockMinimumRawOpinions === null ||
    lockMinimumRawOpinions < 0 ||
    !Number.isInteger(lockMinimumRawOpinions)
  ) {
    errors.push("Lock minimum raw opinions must be a whole number greater than or equal to 0.");
  }
  if (lockMinimumWeightedOpinions === null || lockMinimumWeightedOpinions < 0) {
    errors.push("Lock minimum weighted opinions must be a number greater than or equal to 0.");
  }

  const splitRequirements = [];
  (Array.isArray(splitValues) ? splitValues : []).forEach((row, offset) => {
    const minimumValue = row && row[0];
    const splitValue = row && row[1];
    const minimumBlank = minimumValue === null || minimumValue === undefined || String(minimumValue).trim() === "";
    const splitBlank = splitValue === null || splitValue === undefined || String(splitValue).trim() === "";
    if (minimumBlank && splitBlank) return;
    const sheetRow = TIER_CONFIG_SPLIT_START_ROW + offset;
    if (minimumBlank || splitBlank) {
      errors.push("Row " + sheetRow + ": enter both Minimum weighted opinions and Required split.");
      return;
    }
    const minimumWeightedOpinions = tierConfigurationNumber_(minimumValue);
    const requiredSplit = tierConfigurationNumber_(splitValue);
    if (minimumWeightedOpinions === null || minimumWeightedOpinions < 0) {
      errors.push("Row " + sheetRow + ": Minimum weighted opinions must be 0 or greater.");
      return;
    }
    if (requiredSplit === null || requiredSplit < 0) {
      errors.push("Row " + sheetRow + ": Required split must be 0 or greater.");
      return;
    }
    splitRequirements.push({ minimumWeightedOpinions, requiredSplit, sheetRow });
  });

  if (!splitRequirements.length) {
    errors.push("Enter at least one split requirement.");
  } else {
    if (splitRequirements[0].minimumWeightedOpinions !== 0) {
      errors.push("The first split requirement must begin at 0 weighted opinions.");
    }
    for (let index = 1; index < splitRequirements.length; index++) {
      if (
        splitRequirements[index].minimumWeightedOpinions <=
        splitRequirements[index - 1].minimumWeightedOpinions
      ) {
        errors.push(
          "Row " + splitRequirements[index].sheetRow +
          ": Minimum weighted opinions must increase from the row above."
        );
      }
    }
  }

  return {
    valid: errors.length === 0,
    configuration: {
      minimumOpinions,
      lockThreshold,
      lockMinimumRawOpinions,
      lockMinimumWeightedOpinions,
      splitRequirements: splitRequirements.map(rule => ({
        minimumWeightedOpinions: rule.minimumWeightedOpinions,
        requiredSplit: rule.requiredSplit
      }))
    },
    errors
  };
}

function isTierDecisionConfigurationValid_(configuration) {
  if (!configuration || !Array.isArray(configuration.splitRequirements)) return false;
  const parsed = parseTierDecisionConfiguration_(
    [
      [configuration.minimumOpinions],
      [configuration.lockThreshold],
      [configuration.lockMinimumRawOpinions],
      [configuration.lockMinimumWeightedOpinions]
    ],
    configuration.splitRequirements.map(rule => [
      rule.minimumWeightedOpinions,
      rule.requiredSplit
    ])
  );
  return parsed.valid;
}

function tierConfigurationHasDecisionPanel_(sheet) {
  return sheet && String(sheet.getRange(TIER_CONFIG_DECISION_HEADER_CELL).getDisplayValue() || "").trim() ===
    "Decision rules";
}

function tierConfigurationEdgeRows_(sheet) {
  const fallbackHeaderRow = TIER_CONFIG_SPLIT_START_ROW +
    TIER_CONFIG_MAX_SPLIT_ROW_COUNT + TIER_CONFIG_SECTION_GAP_ROWS;
  if (!sheet || !tierConfigurationHasDecisionPanel_(sheet)) {
    return {
      headerRow: fallbackHeaderRow,
      bottomRow: fallbackHeaderRow + 1,
      topRow: fallbackHeaderRow + 2
    };
  }

  const candidates = sheet.getRange(
    TIER_CONFIG_SPLIT_START_ROW,
    TIER_CONFIG_SPLIT_START_COL,
    TIER_CONFIG_MAX_SPLIT_ROW_COUNT + TIER_CONFIG_SECTION_GAP_ROWS + 1,
    1
  ).getValues();
  const offset = candidates.findIndex(row => String(row && row[0] || "").trim() === TIER_CONFIG_EDGE_HEADER);
  const headerRow = offset >= 0 ? TIER_CONFIG_SPLIT_START_ROW + offset : fallbackHeaderRow;
  return { headerRow, bottomRow: headerRow + 1, topRow: headerRow + 2 };
}

function tierConfigurationSplitRowCount_(sheet) {
  const edgeRows = tierConfigurationEdgeRows_(sheet);
  const rowsBeforeEdge = Math.max(1, edgeRows.headerRow - TIER_CONFIG_SPLIT_START_ROW);
  const precedingRow = edgeRows.headerRow > TIER_CONFIG_SPLIT_START_ROW
    ? sheet.getRange(edgeRows.headerRow - 1, TIER_CONFIG_SPLIT_START_COL, 1, 2).getValues()[0]
    : [];
  const hasSpacer = !precedingRow.some(value => {
    return value !== null && value !== undefined && String(value).trim() !== "";
  });
  return Math.max(
    1,
    Math.min(
      TIER_CONFIG_MAX_SPLIT_ROW_COUNT,
      rowsBeforeEdge - (hasSpacer ? TIER_CONFIG_SECTION_GAP_ROWS : 0)
    )
  );
}

function readTierDecisionConfiguration_(sheet) {
  if (!tierConfigurationHasDecisionPanel_(sheet)) {
    return {
      valid: true,
      missing: true,
      configuration: defaultAnalyzerDecisionConfiguration_(),
      errors: []
    };
  }
  const settingsValues = sheet.getRange(
    TIER_CONFIG_MINIMUM_OPINIONS_CELL + ":" + TIER_CONFIG_LOCK_MINIMUM_WEIGHTED_CELL
  ).getValues();
  const splitValues = sheet.getRange(
    TIER_CONFIG_SPLIT_START_ROW,
    TIER_CONFIG_SPLIT_START_COL,
    tierConfigurationSplitRowCount_(sheet),
    2
  ).getValues();
  const parsed = parseTierDecisionConfiguration_(settingsValues, splitValues);
  return {
    valid: parsed.valid,
    missing: false,
    configuration: parsed.configuration,
    errors: parsed.errors
  };
}

function layoutTierDecisionSections_(sheet, splitValues, bottomValue, topValue) {
  const rules = (Array.isArray(splitValues) ? splitValues : [])
    .filter(row => Array.isArray(row) && row.some(value => {
      return value !== null && value !== undefined && String(value).trim() !== "";
    }))
    .slice(0, TIER_CONFIG_MAX_SPLIT_ROW_COUNT);
  if (!rules.length) rules.push(["", ""]);

  const splitRowCount = rules.length;
  const spacerRow = TIER_CONFIG_SPLIT_START_ROW + splitRowCount;
  const edgeHeaderRow = spacerRow + TIER_CONFIG_SECTION_GAP_ROWS;
  const edgeBottomRow = edgeHeaderRow + 1;
  const edgeTopRow = edgeHeaderRow + 2;
  const panelRange = sheet.getRange(
    TIER_CONFIG_SPLIT_START_ROW,
    TIER_CONFIG_SPLIT_START_COL,
    TIER_CONFIG_PANEL_END_ROW - TIER_CONFIG_SPLIT_START_ROW + 1,
    2
  );
  panelRange
    .clearDataValidations()
    .clearContent()
    .setBackground("#fbfcfe")
    .setFontColor("#202124")
    .setFontFamily("Mukta")
    .setFontWeight("normal")
    .setWrap(false)
    .setVerticalAlignment("middle");

  const splitRange = sheet.getRange(
    TIER_CONFIG_SPLIT_START_ROW,
    TIER_CONFIG_SPLIT_START_COL,
    splitRowCount,
    2
  );
  splitRange
    .setValues(rules)
    .setBackground("#fff2cc")
    .setHorizontalAlignment("right")
    .setNumberFormat("0.###")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(false)
        .build()
    );

  sheet.getRange(edgeHeaderRow, 4, 1, 2)
    .setBackground("#e8f0fe")
    .setFontColor("#294b66")
    .setFontFamily("Mukta")
    .setFontWeight("bold");
  sheet.getRange(edgeHeaderRow, 4).setValue(TIER_CONFIG_EDGE_HEADER);
  sheet.getRange(edgeBottomRow, 4, 2, 1)
    .setValues([["Count toward bottom"], ["Count toward top"]])
    .setBackground("#f7f9fb")
    .setFontColor("#202124")
    .setFontFamily("Mukta")
    .setFontWeight("bold");
  sheet.getRange(edgeBottomRow, 5).setValue(bottomValue).setBackground("#fff2cc");
  sheet.getRange(edgeTopRow, 5).setValue(topValue).setBackground("#fff2cc");
  sheet.getRange(edgeBottomRow, 5, 2, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("middle");

  sheet.setRowHeights(TIER_CONFIG_SPLIT_START_ROW, splitRowCount, 21);
  sheet.setRowHeight(spacerRow, 22);
  sheet.setRowHeight(edgeHeaderRow, 22);
  sheet.setRowHeights(edgeBottomRow, 2, 22);

  return { splitRowCount, spacerRow, edgeHeaderRow, edgeBottomRow, edgeTopRow };
}

function tierConfigurationEditTouchesDecisionSections_(range) {
  if (
    !range ||
    typeof range.getRow !== "function" ||
    typeof range.getColumn !== "function"
  ) return true;

  const firstRow = range.getRow();
  const lastRow = firstRow + (typeof range.getNumRows === "function" ? range.getNumRows() : 1) - 1;
  const firstColumn = range.getColumn();
  const lastColumn = firstColumn +
    (typeof range.getNumColumns === "function" ? range.getNumColumns() : 1) - 1;
  return firstRow <= TIER_CONFIG_PANEL_END_ROW &&
    lastRow >= TIER_CONFIG_SPLIT_START_ROW &&
    firstColumn <= TIER_CONFIG_SPLIT_START_COL + 1 &&
    lastColumn >= TIER_CONFIG_SPLIT_START_COL;
}

function refreshTierDecisionSections_(sheet) {
  if (!sheet || !tierConfigurationHasDecisionPanel_(sheet)) return;
  const edgeRows = tierConfigurationEdgeRows_(sheet);
  const candidateRowCount = Math.max(
    1,
    Math.min(TIER_CONFIG_MAX_SPLIT_ROW_COUNT, edgeRows.headerRow - TIER_CONFIG_SPLIT_START_ROW)
  );
  const splitValues = sheet.getRange(
    TIER_CONFIG_SPLIT_START_ROW,
    TIER_CONFIG_SPLIT_START_COL,
    candidateRowCount,
    2
  ).getValues();
  const bottomValue = sheet.getRange(edgeRows.bottomRow, 5).getValue();
  const topValue = sheet.getRange(edgeRows.topRow, 5).getValue();
  layoutTierDecisionSections_(sheet, splitValues, bottomValue, topValue);
}

function isCachedTierConfigurationValid_(cached) {
  if (
    !cached ||
    cached.schema !== TIER_CONFIG_CACHE_SCHEMA ||
    !Array.isArray(cached.rows) ||
    !isTierDecisionConfigurationValid_(cached.decisionConfiguration)
  ) {
    return false;
  }
  if (cached.rows.length < 2) return false;
  return cached.rows.every((row, index) => {
    return row &&
      row.order === index + 1 &&
      typeof row.name === "string" && row.name.trim() !== "" &&
      isTierHexColor_(row.primaryColor) &&
      (row.borderColor === "" || isTierHexColor_(row.borderColor)) &&
      (row.fontColor === "" || isTierHexColor_(row.fontColor)) &&
      (row.splitFontColor === "" || isTierHexColor_(row.splitFontColor)) &&
      ["Bottom", "Top", "Own"].includes(row.group);
  });
}

function readCachedTierConfiguration_() {
  const cache = getTierConfigurationCache_();
  if (!cache) {
    tierConfigurationCacheStatus_ = "unavailable";
    return null;
  }
  try {
    const serialized = cache.get(TIER_CONFIG_CACHE_KEY);
    if (!serialized) {
      tierConfigurationCacheStatus_ = "miss";
      return null;
    }
    const cached = JSON.parse(serialized);
    if (!isCachedTierConfigurationValid_(cached)) {
      tierConfigurationCacheStatus_ = "invalid";
      cache.remove(TIER_CONFIG_CACHE_KEY);
      return null;
    }
    tierConfigurationCacheStatus_ = "hit";
    return {
      valid: true,
      missing: false,
      rows: cached.rows,
      splitCount: cached.splitCount || 0,
      decisionConfiguration: cached.decisionConfiguration,
      decisionConfigurationMissing: false,
      errors: []
    };
  } catch (error) {
    tierConfigurationCacheStatus_ = "error";
    try { cache.remove(TIER_CONFIG_CACHE_KEY); } catch (removeError) {}
    return null;
  }
}

function cacheTierConfiguration_(result) {
  if (!result || !result.valid || !isCachedTierConfigurationValid_({
    schema: TIER_CONFIG_CACHE_SCHEMA,
    rows: result.rows,
    decisionConfiguration: result.decisionConfiguration || defaultAnalyzerDecisionConfiguration_()
  })) return;
  const cache = getTierConfigurationCache_();
  if (!cache) return;
  try {
    cache.put(TIER_CONFIG_CACHE_KEY, JSON.stringify({
      schema: TIER_CONFIG_CACHE_SCHEMA,
      rows: result.rows,
      splitCount: result.splitCount || 0,
      decisionConfiguration: result.decisionConfiguration || defaultAnalyzerDecisionConfiguration_()
    }), TIER_CONFIG_CACHE_TTL_SECONDS);
  } catch (error) {}
}

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
  invalidateTierConfigurationCache_();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TIER_CONFIG_SHEET_NAME);
  const hasExistingTierList = sh &&
    String(sh.getRange("A4").getDisplayValue()).trim() === "Paste tier list";
  const hasExistingDecisionPanel = sh && tierConfigurationHasDecisionPanel_(sh);
  const existingEdgeRows = hasExistingDecisionPanel ? tierConfigurationEdgeRows_(sh) : null;
  const existingBottomValue = hasExistingTierList
    ? hasExistingDecisionPanel
      ? sh.getRange(existingEdgeRows.bottomRow, 5).getValue()
      : sh.getRange(TIER_CONFIG_LEGACY_BOTTOM_CELL).getValue()
    : "Insane Demon";
  const existingTopValue = hasExistingTierList
    ? hasExistingDecisionPanel
      ? sh.getRange(existingEdgeRows.topRow, 5).getValue()
      : sh.getRange(TIER_CONFIG_LEGACY_TOP_CELL).getValue()
    : "";
  const defaultDecisionConfiguration = defaultAnalyzerDecisionConfiguration_();
  const existingSettings = hasExistingTierList && hasExistingDecisionPanel
    ? sh.getRange(
      TIER_CONFIG_MINIMUM_OPINIONS_CELL + ":" + TIER_CONFIG_LOCK_MINIMUM_WEIGHTED_CELL
    ).getValues()
    : [
      [defaultDecisionConfiguration.minimumOpinions],
      [defaultDecisionConfiguration.lockThreshold],
      [defaultDecisionConfiguration.lockMinimumRawOpinions],
      [defaultDecisionConfiguration.lockMinimumWeightedOpinions]
    ];
  const existingSplitRequirements = hasExistingTierList && hasExistingDecisionPanel
    ? sh.getRange(
      TIER_CONFIG_SPLIT_START_ROW,
      TIER_CONFIG_SPLIT_START_COL,
      tierConfigurationSplitRowCount_(sh),
      2
    ).getValues().filter(row => row.some(value => {
      return value !== null && value !== undefined && String(value).trim() !== "";
    }))
    : defaultDecisionConfiguration.splitRequirements.map(rule => [
      rule.minimumWeightedOpinions,
      rule.requiredSplit
    ]);
  const splitRequirements = existingSplitRequirements
    .slice(0, TIER_CONFIG_MAX_SPLIT_ROW_COUNT);
  if (!splitRequirements.length) splitRequirements.push(["", ""]);

  if (!sh) {
    sh = ss.insertSheet(TIER_CONFIG_SHEET_NAME);
  }
  setManagedSheetColumnCount_(sh, TIER_CONFIG_REQUIRED_COLUMNS);

  if (!hasExistingTierList) {
    const rows = defaultTierConfigRows_();
    const pasted = tierRowsToPasteList_(rows);
    sh.getDataRange().breakApart();
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getRange(TIER_CONFIG_START_ROW, TIER_CONFIG_LIST_COL, pasted.values.length, 1)
      .setValues(pasted.values)
      .setBackgrounds(pasted.backgrounds)
      .setFontColors(pasted.fontColors)
      .setFontFamily("Mukta")
      .setVerticalAlignment("middle");
  } else {
    sh.getRange("A1:E3").breakApart().clear();
    sh.getRange(4, 4, TIER_CONFIG_PANEL_END_ROW - 3, 2)
      .breakApart()
      .clearDataValidations()
      .clear();
  }

  sh.getRange("A1:E1").mergeAcross()
    .setValue("Tier Configuration")
    .setFontFamily("Mukta")
    .setFontSize(17)
    .setFontWeight("bold")
    .setBackground("#315f86")
    .setFontColor("#f8fbff")
    .setHorizontalAlignment("left");
  sh.getRange("A2:E2").mergeAcross()
    .setValue("Paste tier colors in column A. Edit yellow cells; separate edge tiers with commas or new lines.")
    .setFontFamily("Mukta")
    .setFontSize(10)
    .setFontColor("#40566b")
    .setBackground("#f4f7fb");
  sh.getRange("A3:E3").mergeAcross()
    .setValue("Split requirements use weighted opinions to determine whether a level qualifies for placement or movement.")
    .setFontFamily("Mukta")
    .setFontSize(9)
    .setFontColor("#4a5f75")
    .setBackground("#e8f0fe");

  sh.getRange("A4").setValue("Paste tier list")
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setFontColor("#24435b")
    .setBackground("#d7e5f3");

  sh.getRange("D4:E4")
    .setBackground("#d7e5f3")
    .setFontColor("#24435b")
    .setFontFamily("Mukta")
    .setFontWeight("bold");
  sh.getRange(TIER_CONFIG_DECISION_HEADER_CELL).setValue("Decision rules");
  sh.getRange("D5:D8").setValues([
    ["Minimum opinions"],
    ["Lock threshold"],
    ["Lock minimum raw opinions"],
    ["Lock minimum weighted opinions"]
  ])
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setFontColor("#202124")
    .setBackground("#f7f9fb");
  sh.getRange("E5:E8").setValues(existingSettings)
    .setBackground("#fff2cc")
    .setHorizontalAlignment("right");
  sh.getRange("D10:E10")
    .setBackground("#e8f0fe")
    .setFontColor("#294b66")
    .setFontFamily("Mukta")
    .setFontWeight("bold");
  sh.getRange("D10").setValue("Split requirements");
  sh.getRange("D11:E11").setValues([[
    "Minimum weighted opinions",
    "Required split"
  ]])
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setFontColor("#334e63")
    .setBackground("#f1f5f9");
  sh.getRange("E11").setHorizontalAlignment("center");
  sh.getRange("D4:E11").setWrap(false);
  const sectionLayout = layoutTierDecisionSections_(
    sh,
    splitRequirements,
    existingBottomValue,
    existingTopValue
  );

  const nonnegativeValidation = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .build();
  const percentValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1)
    .setAllowInvalid(false)
    .build();
  const wholeNumberValidation = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied("=AND(ISNUMBER(E7),E7>=0,MOD(E7,1)=0)")
    .setAllowInvalid(false)
    .build();
  sh.getRange(TIER_CONFIG_MINIMUM_OPINIONS_CELL).setDataValidation(nonnegativeValidation);
  sh.getRange(TIER_CONFIG_LOCK_THRESHOLD_CELL).setDataValidation(percentValidation);
  sh.getRange(TIER_CONFIG_LOCK_MINIMUM_RAW_CELL).setDataValidation(wholeNumberValidation);
  sh.getRange(TIER_CONFIG_LOCK_MINIMUM_WEIGHTED_CELL).setDataValidation(nonnegativeValidation);

  sh.getRange("E5").setNumberFormat("0.###");
  sh.getRange("E6").setNumberFormat("0%");
  sh.getRange("E7:E8").setNumberFormat("0.###");
  sh.getRange(1, 1, sectionLayout.edgeTopRow, 5)
    .setFontFamily("Mukta")
    .setVerticalAlignment("middle");

  sh.setFrozenRows(TIER_CONFIG_HEADER_ROW);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 245);
  sh.setColumnWidth(2, 16);
  sh.setColumnWidth(3, 16);
  sh.setColumnWidth(4, 206);
  sh.setColumnWidth(5, 100);
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 21);
  sh.setRowHeight(3, 21);
  sh.setRowHeight(4, 24);
  sh.setRowHeights(5, 4, 22);
  sh.setRowHeight(9, 22);
  sh.setRowHeight(10, 22);
  sh.setRowHeight(11, 24);
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
    if (color === LEVEL_LOCK_BLACK_MARKER) {
      errors.push("Row " + sheetRow + ": " + LEVEL_LOCK_BLACK_MARKER + " is reserved for locked black cells.");
    }
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
  const hasDecisionPanel = tierConfigurationHasDecisionPanel_(sh);
  const edgeRows = hasDecisionPanel ? tierConfigurationEdgeRows_(sh) : null;
  const tierResult = parseTierConfigList_(
    range.getDisplayValues(),
    range.getBackgrounds(),
    hasDecisionPanel
      ? sh.getRange(edgeRows.bottomRow, 5).getDisplayValue()
      : sh.getRange(TIER_CONFIG_LEGACY_BOTTOM_CELL).getDisplayValue(),
    hasDecisionPanel
      ? sh.getRange(edgeRows.topRow, 5).getDisplayValue()
      : sh.getRange(TIER_CONFIG_LEGACY_TOP_CELL).getDisplayValue(),
    range.getFontColors()
  );
  const decisionResult = readTierDecisionConfiguration_(sh);
  return {
    valid: tierResult.valid && decisionResult.valid,
    missing: false,
    rows: tierResult.rows,
    splitCount: tierResult.splitCount,
    decisionConfiguration: decisionResult.configuration,
    decisionConfigurationMissing: decisionResult.missing,
    errors: tierResult.errors.concat(decisionResult.errors)
  };
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
  resetAnalyzerDecisionConfiguration_();
}

function normalizeTierConfigurationResult_(result) {
  if (!result || !result.valid) return result;
  if (!isTierDecisionConfigurationValid_(result.decisionConfiguration)) {
    result.decisionConfiguration = defaultAnalyzerDecisionConfiguration_();
    result.decisionConfigurationMissing = true;
  }
  return result;
}

function loadTierConfiguration_(force, timer) {
  if (tierConfigurationLoaded_ && !force) {
    tierConfigurationLoadSource_ = "memory";
    return tierConfigurationResult_;
  }
  if (force) invalidateTierConfigurationCache_();
  tierConfigurationCacheStatus_ = force ? "bypassed" : "not-checked";
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const cachedResult = !force ? readCachedTierConfiguration_() : null;
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "tierConfigurationCacheLookup", phaseStartedAt);
  }
  let result = cachedResult;
  if (!result) {
    phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    result = readTierConfiguration_();
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "tierConfigurationSheetRead", phaseStartedAt);
    }
  }
  result = normalizeTierConfigurationResult_(result);
  tierConfigurationLoadSource_ = cachedResult ? "document-cache" : "sheet";
  if (result.valid) {
    applyTierConfiguration_(result.rows);
    applyAnalyzerDecisionConfiguration_(result.decisionConfiguration);
  }
  else if (result.missing) resetTierConfigurationToDefaults_();
  if (result.valid && !cachedResult) {
    phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    cacheTierConfiguration_(result);
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "tierConfigurationCacheWrite", phaseStartedAt);
    }
  }
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
