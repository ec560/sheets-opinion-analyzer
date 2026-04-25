// Main analysis function that reads opinions, applies weights, calculates stats, and outputs results
function analyzeSelectedLevel() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);

  tool.getRange(OUTPUT_START_ROW, OUTPUT_COL, tool.getMaxRows(), OUTPUT_WIDTH).clearContent().breakApart();

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) {
    SpreadsheetApp.getUi().alert("Pick a tier and a level first.");
    return;
  }

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

  const analysis = calculateLevelAnalysis_(tierName, levelName, vals, bgs, fcs);
  const {
    weightsByTier,
    countedRowFlags,
    fuckPresent,
    fuckWeight,
    tierWeightSum,
    allWeight,
    allAndFuck,
    rawCount,
    rawMeanIdx,
    rawMedianIdx,
    rawMeanLabel,
    rawMedianLabel,
    outlierPctObj,
    outlierBounds,
    sd,
    topTier,
    topWeight,
    secondTier,
    secondWeight,
    splitLowTier,
    splitHighTier,
    splitLeftTotal,
    splitRightTotal,
    toppct,
    fuckpct,
    passesMajority,
    passesSplitMajority,
    verdictDiffersFromCurrent,
    isPending,
    verdictTier,
    verdictTierName,
    verdictBaseName,
    canMove,
    minimumOpinionWeight,
    lockSharePct,
    passesSplitPct,
    splitMarginPct,
    splitMargin,
    splitThreshold,
    currentTier
  } = analysis;

  const out = [];
  out.push([`Tier sheet`, tierName, "", ""]);
  out.push([`Level`, levelName, "", ""]);
  out.push([`Total weighted opinions`, allAndFuck, "", VERSION]);
  out.push([`Most votes (weighted)`, topTier, topWeight, ""]);
  out.push([`Runner-up (weighted)`, secondTier, secondWeight, ""]);
  out.push([`Tier Mean`, rawMeanLabel, rawMeanIdx, ""]);
  out.push([`Tier Median`, rawMedianLabel, rawMedianIdx, ""]);
  out.push([`Outliers`, `${outlierPctObj.count} / ${outlierPctObj.total} (${(outlierPctObj.pct * 100).toFixed(1)}%)`, "", ""]);
  out.push([
    "Outlier range",
    `<= ${outlierBounds.lowLabel}` + `       >= ${outlierBounds.highLabel}`,
    "",
    ""
  ]);
  out.push([`Standard Deviation`, sd, "", ""]);
  out.push([`Place/Move`, canMove ? "YES" : "NO", "", canMove ? verdictTierName : ""]);

  function upTo3dec_(x) {
    return Number(x).toFixed(3).replace(/\.?0+$/, "");
  }

  out.push([
    `Split`,
    `${splitLowTier}`,
    `${splitHighTier}`,
    `${upTo3dec_(splitLeftTotal)} | ${upTo3dec_(splitRightTotal)}`
  ]);

  out.push(["", "", "", ""]);
  if (fuckPresent) {
    out.push(["Fuck % of all", "", "", ""]);
    out.push(["Fuck", fuckWeight, fuckpct, ""]);
    out.push(["", "", "", ""]);
  }
  out.push(["Tier", "Weighted", "%", ""]);
  for (const t of orderedTierNames) {
    const w = weightsByTier[t] || 0;
    const denom = tierWeightSum > 0 ? tierWeightSum : 1;
    const share = w / denom;
    out.push([t, w, share, ""]);
  }

  const startRow = 1;
  const startCol = OUTPUT_COL;
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

  tool.getRange(startRow + 9, startCol + 2, Math.max(0, out.length - 9), 1).setNumberFormat("0.0%");

  formatAnalysisOutput_(
    tool,
    out.length,
    orderedTierNames,
    weightsByTier,
    topTier,
    topWeight,
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
    verdictBaseName,
    minimumOpinionWeight,
    rawCount,
    lockSharePct,
    allWeight,
    passesSplitPct,
    splitMarginPct,
    splitMargin,
    splitThreshold
  );

  applyCountedPlayerHighlights_(tool, DATA_START_ROW, bgs, countedRowFlags);
}
