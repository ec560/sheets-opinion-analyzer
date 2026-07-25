const TIER_FLAG_STYLES = {
  movement: {
    priority: 1,
    background: "#b3261e",
    text: "#ffffff",
    fontWeight: "bold",
    styleDifference: false
  },
  book: {
    priority: 2,
    background: "#f9ab00",
    text: "#3c2400",
    fontWeight: "bold",
    styleDifference: true
  },
  lock: {
    priority: 3,
    background: "#d2e3fc",
    text: "#174ea6",
    fontWeight: "bold",
    styleDifference: false
  },
  needs_more: {
    priority: 4,
    background: "#fef7e0",
    text: "#7a4f01",
    fontWeight: "bold",
    styleDifference: true
  },
  low_unsettled: {
    priority: 5,
    background: "#fce8e6",
    text: "#b3261e",
    fontWeight: "bold",
    styleDifference: true
  },
  low_solid: {
    priority: 6,
    background: "#f1f3f4",
    text: "#5f6368",
    fontWeight: "normal",
    styleDifference: true
  },
  neutral: {
    priority: 99,
    background: "#f5f5f5",
    text: "#202124",
    fontWeight: "bold",
    styleDifference: false
  }
};

function tierFlagStyle_(styleKey) {
  return TIER_FLAG_STYLES[styleKey] || TIER_FLAG_STYLES.neutral;
}

function scanSelectedTierFlags() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (typeof loadTierConfiguration_ === "function") {
    const configResult = loadTierConfiguration_();
    const configError = tierConfigurationErrorMessage_(configResult);
    if (configError) {
      ui.alert("Error with Tier Configuration", configError, ui.ButtonSet.OK);
      return;
    }
  }
  const tierName = getFlagScanTierName_(ss, tool);

  if (!tierName) {
    ui.alert("Tier Flag Scan", "Open a tier sheet or pick a tier in " + TIER_CELL + " first.", ui.ButtonSet.OK);
    return;
  }

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet || isAnalyzerUtilitySheetName_(tierName)) {
    ui.alert("Tier Flag Scan", "Could not find a tier sheet named \"" + tierName + "\".", ui.ButtonSet.OK);
    return;
  }

  const result = buildTierFlagScan_(tierName, tierSheet);
  if (result.scanned === 0) {
    ui.alert("Tier Flag Scan", "No level headers found on \"" + tierName + "\".", ui.ButtonSet.OK);
    return;
  }

  const outputSheet = renderTierFlagScan_(ss, tierName, result.rows);
  ss.setActiveSheet(outputSheet);

  ui.alert(
    "Tier Flag Scan",
    result.rows.length + " flagged level" + (result.rows.length === 1 ? "" : "s") +
      " found across " + result.scanned + " level" + (result.scanned === 1 ? "" : "s") + ".",
    ui.ButtonSet.OK
  );
}

function getFlagScanTierName_(ss, tool) {
  const active = ss.getActiveSheet();
  if (active) {
    const activeName = active.getName();
    if (!isAnalyzerUtilitySheetName_(activeName)) return activeName;
  }

  if (!tool) return "";
  return String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim();
}

function buildTierFlagScan_(tierName, tierSheet) {
  const headers = getLevelHeaders_(tierSheet);
  const lastRow = tierSheet.getLastRow();
  const lastCol = tierSheet.getLastColumn();
  const numRows = Math.max(0, lastRow - 1);

  let vals = [];
  let bgs = [];
  let fcs = [];
  if (numRows > 0 && lastCol > 0) {
    const sourceRange = tierSheet.getRange(2, 1, numRows, lastCol);
    vals = sourceRange.getValues();
    bgs = sourceRange.getBackgrounds();
    fcs = sourceRange.getFontColors();
  }

  const rows = [];
  headers.forEach(header => {
    const levelData = extractLevelFlagData_(header, vals, bgs, fcs, lastCol);
    const analysis = calculateLevelAnalysis_(
      tierName,
      header.name,
      levelData.vals,
      levelData.bgs,
      levelData.fcs
    );
    if (analysis.rawCount === 0 && analysis.allAndFuck === 0) return;

    const flagSummary = buildLevelFlagSummary_(analysis);
    if (flagSummary.flags.length > 0) rows.push(buildTierFlagRow_(analysis, flagSummary));
  });

  return {
    scanned: headers.length,
    rows
  };
}

function extractLevelFlagData_(header, vals, bgs, fcs, lastCol) {
  const outVals = [];
  const outBgs = [];
  const outFcs = [];
  const startIdx = header.col - 1;

  if (startIdx + 2 >= lastCol) {
    return { vals: outVals, bgs: outBgs, fcs: outFcs };
  }

  for (let r = 0; r < vals.length; r++) {
    const row = [
      safeCellValue_(vals[r][startIdx]),
      safeCellValue_(vals[r][startIdx + 1]),
      safeCellValue_(vals[r][startIdx + 2])
    ];

    if (isBlankFlagOpinionRow_(row)) continue;

    outVals.push(row);
    outBgs.push([
      bgs[r][startIdx],
      bgs[r][startIdx + 1],
      bgs[r][startIdx + 2]
    ]);
    outFcs.push([
      fcs[r][startIdx],
      fcs[r][startIdx + 1],
      fcs[r][startIdx + 2]
    ]);
  }

  return { vals: outVals, bgs: outBgs, fcs: outFcs };
}

function isBlankFlagOpinionRow_(row) {
  return String(row[0]).trim() === "" &&
    String(row[1]).trim() === "" &&
    String(row[2]).trim() === "";
}

function buildLevelFlagSummary_(analysis) {
  let flag = "";
  let styleKey = "";
  let differenceAlert = analysis.isLockWorthy ? null : analysis.difference;
  const needsMoreOpinionsAlert = buildNeedsMoreOpinionsAlert_(analysis);
  const lowOpinionAlert = buildLowOpinionAlert_(analysis);

  if (analysis.canMove) {
    styleKey = "movement";
    if (analysis.isPending) {
      flag = "placement alert (" + analysis.verdictTier + ")";
    } else {
      const direction = analysis.moveDirection ? ", " + analysis.moveDirection : "";
      flag = "move alert (" + analysis.verdictTier + direction + ")";
    }
  } else if (analysis.isLockWorthy) {
    styleKey = "lock";
    flag = "lock alert (" + (analysis.lockSharePct * 100).toFixed(0) + "% " + analysis.lockSideLabel + ")";
  } else if (analysis.bookAlert) {
    styleKey = "book";
    flag = formatBookAlert_(analysis.bookAlert);
    differenceAlert = analysis.bookAlert;
  } else if (needsMoreOpinionsAlert) {
    styleKey = "needs_more";
    flag = "needs more opinions";
    differenceAlert = needsMoreOpinionsAlert;
  } else if (lowOpinionAlert) {
    styleKey = lowOpinionAlert.subdued ? "low_solid" : "low_unsettled";
    flag = lowOpinionAlert.text;
  }

  return {
    flags: flag ? [flag] : [],
    differenceAlert,
    styleKey,
    priority: tierFlagStyle_(styleKey).priority,
    subduedLowOpinion: !!(lowOpinionAlert && lowOpinionAlert.subdued && flag === lowOpinionAlert.text)
  };
}

function buildLowOpinionAlert_(analysis) {
  if (!analysis) return null;

  const rawCountIsLow = analysis.rawCount <= FLAG_SCAN_LOW_OPINION_MAX_RAW_COUNT;
  const weightIsLow = analysis.allWeight < FLAG_SCAN_LOW_OPINION_WEIGHT;
  const strongCurrentTierLean = hasStrongCurrentTierLean_(analysis);

  if (!rawCountIsLow && (!weightIsLow || strongCurrentTierLean)) return null;

  return {
    text: formatLowOpinionFlag_(analysis),
    subdued: rawCountIsLow && strongCurrentTierLean
  };
}

function hasStrongCurrentTierLean_(analysis) {
  if (!analysis || analysis.decisionCurrentIdx < 0 || !analysis.difference) return false;

  const currentSide = analysis.decisionCurrentIdx <= analysis.decisionTierSplit.lowIndex ? "left" : "right";
  return analysis.difference.winningSide === currentSide &&
    analysis.difference.lean >= FLAG_SCAN_LOW_OPINION_CURRENT_LEAN;
}

function formatLowOpinionFlag_(analysis) {
  return "low opinion count (" +
    formatFlagNumber_(analysis.allWeight) +
    "/" +
    FLAG_SCAN_LOW_OPINION_WEIGHT +
    " weight)";
}

function buildNeedsMoreOpinionsAlert_(analysis) {
  if (!analysis || analysis.decisionCurrentIdx < 0 || !analysis.difference) return null;
  if (analysis.allWeight > FLAG_SCAN_NEEDS_MORE_OPINIONS_MAX_WEIGHT) return null;
  if (analysis.bookAlert) return null;
  if (analysis.difference.lean > FLAG_BOOK_LEAN_WEIGHT) return null;

  const currentSide = analysis.decisionCurrentIdx <= analysis.decisionTierSplit.lowIndex ? "left" : "right";
  const favorsCurrentTier = analysis.difference.lean === 0 || analysis.difference.winningSide === currentSide;

  return {
    lean: analysis.difference.lean,
    tier: favorsCurrentTier ? analysis.currentTier : analysis.difference.tier
  };
}

function buildTierFlagRow_(analysis, flagSummary) {
  return {
    values: [
      analysis.levelName,
      formatFlagNumber_(analysis.allWeight),
      analysis.rawCount,
      flagSummary.flags.join("; "),
      analysis.tierSplit.left.label + " / " + analysis.tierSplit.right.label,
      formatFlagNumber_(analysis.tierSplit.left.weight) + " | " + formatFlagNumber_(analysis.tierSplit.right.weight),
      formatDifference_(flagSummary.differenceAlert)
    ],
    styleKey: flagSummary.styleKey || "",
    priority: flagSummary.priority || 99,
    subduedLowOpinion: !!flagSummary.subduedLowOpinion
  };
}

function renderTierFlagScan_(ss, tierName, flagRows) {
  let sh = ss.getSheetByName(FLAG_SCAN_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(FLAG_SCAN_SHEET_NAME);

  const headers = [
    "Level",
    "Weight",
    "Raw",
    "Flags",
    "Split",
    "Split totals",
    "Difference"
  ];
  setManagedSheetColumnCount_(sh, headers.length);
  sh.clear();

  const titleRow = new Array(headers.length).fill("");
  titleRow[0] = "Tier flag scan";
  titleRow[1] = tierName;
  titleRow[2] = VERSION;
  titleRow[3] = new Date();

  const rows = [titleRow, headers];
  if (flagRows.length > 0) {
    flagRows.forEach(row => rows.push(row.values));
  } else {
    const emptyRow = new Array(headers.length).fill("");
    emptyRow[0] = "No flagged levels";
    rows.push(emptyRow);
  }

  sh.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatTierFlagScan_(sh, rows, headers.length, flagRows);
  return sh;
}

function formatTierFlagScan_(sh, rows, colCount, flagRows) {
  const rowCount = rows.length;
  if (rowCount === 0 || colCount === 0) return;

  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, rowCount, colCount)
    .setFontFamily("Mukta")
    .setFontSize(10)
    .setFontColor("#202124")
    .setBackground("#ffffff")
    .setVerticalAlignment("middle");

  sh.getRange(1, 1, 1, colCount)
    .setBackground("#e8f0fe")
    .setFontWeight("bold")
    .setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(1, 3)
    .setFontColor("#5f6368")
    .setFontWeight("normal")
    .setHorizontalAlignment("right");

  sh.getRange(1, 4)
    .setNumberFormat("m/d/yyyy h:mm")
    .setFontColor("#5f6368")
    .setFontWeight("normal")
    .setHorizontalAlignment("left");
  styleTierFlagCell_(sh, 1, 2, rows[0][1]);

  sh.getRange(2, 1, 1, colCount)
    .setBackground("#eeeeee")
    .setFontWeight("bold")
    .setBorder(true, false, true, false, false, false, "#d6d9dc", SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(4, 225);
  sh.setColumnWidth(5, 160);
  sh.setColumnWidth(7, 125);

  if (rowCount <= 2) return;

  sh.getRange(3, 2, rowCount - 2, 1).setNumberFormat("0.###");
  sh.getRange(3, 3, rowCount - 2, 1).setNumberFormat("0");
  sh.getRange(3, 2, rowCount - 2, 2).setHorizontalAlignment("right");

  for (let i = 2; i < rowCount; i++) {
    const rowNum = i + 1;
    const row = rows[i];
    const flagRow = flagRows && flagRows.length > (i - 2) ? flagRows[i - 2] : null;
    const flagStyle = tierFlagStyle_(flagRow && flagRow.styleKey);

    if (String(row[0] || "") === "No flagged levels") {
      sh.getRange(rowNum, 1, 1, colCount)
        .setBackground("#f5f5f5")
        .setFontColor("#5f6368")
        .setFontStyle("italic");
      continue;
    }

    const differenceCell = sh.getRange(rowNum, 7);
    sh.getRange(rowNum, 1, 1, colCount)
      .setBorder(false, false, true, false, false, false, "#eeeeee", SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(rowNum, 1).setFontWeight(flagStyle.fontWeight);

    const flagCell = sh.getRange(rowNum, 4);
    flagCell
      .setBackground(flagStyle.background)
      .setFontColor(flagStyle.text)
      .setFontWeight(flagStyle.fontWeight);

    if (flagStyle.styleDifference) {
      differenceCell
        .setBackground(flagStyle.background)
        .setFontColor(flagStyle.text)
        .setFontWeight(flagStyle.fontWeight);
    }
  }
}

function styleTierFlagCell_(sh, row, col, tierName) {
  const name = String(tierName || "").trim();
  if (!name) return;

  const nameToColor = buildTierNameToColor_();
  const hex = nameToColor[name];
  if (!hex) return;

  sh.getRange(row, col)
    .setBackground(hex)
    .setFontColor(tierTextColor_(name))
    .setFontWeight("bold");
}

function formatFlagNumber_(value) {
  const n = Number(value);
  if (!isFinite(n)) return "";
  return n.toFixed(3).replace(/\.?0+$/, "");
}

function formatBookAlert_(alert) {
  return alert.direction + " book alert (" + formatDifference_(alert) + ")";
}

function formatDifference_(alert) {
  if (!alert) return "";

  return "+" +
    formatFlagNumber_(alert.lean) +
    " " +
    alert.tier;
}
