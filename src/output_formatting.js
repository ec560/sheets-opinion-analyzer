function applyCountedPlayerHighlights_(tool, startRow, sourceBackgrounds, countedRowFlags, sourceValues) {
  if (!sourceBackgrounds || sourceBackgrounds.length === 0) return;

  const duplicateHighlights = buildDuplicatePlayerHighlights_(sourceValues || []);
  const fills = sourceBackgrounds.map((row, idx) => {
    return [duplicateHighlights[idx] ||
      (countedRowFlags && countedRowFlags[idx] ? COUNTED_PLAYER_HIGHLIGHT : "#ffffff")];
  });

  tool.getRange(startRow, 1, fills.length, 1).setBackgrounds(fills);
}

function buildDuplicatePlayerHighlights_(values) {
  const highlights = new Array(values.length).fill("");
  const players = new Map();

  values.forEach((row, index) => {
    const player = String(row[0] ?? "").trim().toLowerCase();
    if (!player) return;
    if (!players.has(player)) players.set(player, { current: [], preUpdate: [] });
    const reliabilityText = String(row[2] ?? "");
    const isPreUpdate = /\bpre[\s\-\u2010-\u2015]*update\b/i.test(reliabilityText);
    const group = players.get(player);
    (isPreUpdate ? group.preUpdate : group.current).push(index);
  });

  for (const group of players.values()) {
    const hasBothVersions = group.current.length > 0 && group.preUpdate.length > 0;
    for (const rows of [group.current, group.preUpdate]) {
      // Same-version duplicates take priority over a pre-update/current pair.
      const color = rows.length > 1
        ? DUPLICATE_PLAYER_HIGHLIGHT
        : hasBothVersions ? PRE_UPDATE_PLAYER_HIGHLIGHT : "";
      rows.forEach(index => highlights[index] = color);
    }
  }
  return highlights;
}

function trimFixed_(value, digits) {
  return Number(value).toFixed(digits).replace(/\.?0+$/, "");
}

function buildSplitNotDecisiveMessage_(splitMargin, verdictBaseName) {
  return "Split not decisive (+" + trimFixed_(splitMargin, 2) + " " + verdictBaseName.toLowerCase() + ")";
}

function buildLowSplitMarginMessage_(splitMarginPct, isPending) {
  const requiredPct = 0.05;
  return "Low split margin (" +
    (splitMarginPct * 100).toFixed(0) +
    "% ; " +
    (requiredPct * 100).toFixed(0) +
    "% required)";
}

function resolvePlaceMoveFailure_(ctx) {
  const {
    lockSharePct,
    verdictBaseName,
    placementComparison,
    moveFailureReason
  } = ctx;

  const comparisonLabel = placementComparison.decisionLabel || verdictBaseName;
  const splitNotDecisiveMessage = buildSplitNotDecisiveMessage_(
    placementComparison.marginWeight,
    comparisonLabel
  );
  switch (moveFailureReason) {
    case "needs_more_opinions":
      return { text: "Needs more opinions" };
    case "no_movement_f":
      return { text: "No movement necessary (F)", bg: "#efefef" };
    case "split_not_decisive_f":
      return { text: "Split not decisive (F)" };
    case "fuck_rules_not_met":
      return { text: "Rules not met (F)" };
    case "split_not_decisive":
      return { text: splitNotDecisiveMessage };
    case "fuck_placement_not_decisive":
      return { text: splitNotDecisiveMessage };
    case "low_split_margin":
      return {
        text: buildLowSplitMarginMessage_(placementComparison.marginPct, ctx.isPending),
        bg: "#ffdfcc"
      };
    case "lock_threshold_met":
      return {
        text: "Lock threshold met (" + (lockSharePct * 100).toFixed(0) + "%)",
        bg: "#e6f4ea"
      };
    case "no_movement":
      return { text: "No movement necessary", bg: "#efefef" };
    default:
      return { text: "Rules not met (please tell opayc)" };
  }
}

function createFormatMatrix_(rowCount, value) {
  return Array.from({ length: rowCount }, () => Array(OUTPUT_WIDTH).fill(value));
}

function createAnalysisStylePlan_(rowCount) {
  return {
    backgrounds: createFormatMatrix_(rowCount, null),
    fontColors: createFormatMatrix_(rowCount, "#000000"),
    fontWeights: createFormatMatrix_(rowCount, "normal"),
    fontStyles: createFormatMatrix_(rowCount, "normal"),
    fontLines: createFormatMatrix_(rowCount, "none"),
    fontSizes: createFormatMatrix_(rowCount, 10),
    horizontalAlignments: createFormatMatrix_(rowCount, null),
    verticalAlignments: createFormatMatrix_(rowCount, null)
  };
}

function fillFormatRow_(matrix, row, value) {
  if (row < 0 || row >= matrix.length) return;
  matrix[row].fill(value);
}

function fillFormatSpan_(matrix, row, startCol, width, value) {
  if (row < 0 || row >= matrix.length) return;
  for (let col = startCol; col < startCol + width; col++) {
    matrix[row][col] = value;
  }
}

function addWeightedDistributionStyles_(
  styles,
  headerIndex,
  names,
  weightsByName,
  colorsByName
) {
  fillFormatRow_(styles.backgrounds, headerIndex, "#eeeeee");
  fillFormatRow_(styles.fontWeights, headerIndex, "bold");

  names.forEach((name, index) => {
    const row = headerIndex + index + 1;
    const weight = weightsByName[name] || 0;
    const isPositive = weight > 0;

    styles.backgrounds[row][0] = isPositive ? colorsByName[name] || "#999999" : null;
    fillFormatRow_(styles.fontColors, row, isPositive ? "#000000" : "#9e9e9e");
    styles.fontColors[row][0] = isPositive ? tierTextColor_(name) : "#9e9e9e";
    styles.fontWeights[row][0] = "bold";
  });
}

function applyAnalysisStylePlan_(tool, startRow, startCol, styles) {
  tool.getRange(startRow, startCol, styles.backgrounds.length, OUTPUT_WIDTH)
    .setBackgrounds(styles.backgrounds)
    .setFontColors(styles.fontColors)
    .setFontWeights(styles.fontWeights)
    .setFontStyles(styles.fontStyles)
    .setFontLines(styles.fontLines)
    .setFontSizes(styles.fontSizes)
    .setHorizontalAlignments(styles.horizontalAlignments)
    .setVerticalAlignments(styles.verticalAlignments);
}

function setWeightedDistributionFormulas_(
  tool,
  headerRow,
  startCol,
  names,
  weightsByName,
  colorsByName,
  totalWeightedOpinions
) {
  if (names.length === 0) return;

  const firstDataRow = headerRow + 1;
  const lastDataRow = firstDataRow + names.length - 1;
  const confidenceScale = Math.min(
    0.2 + 0.8 * (totalWeightedOpinions / 45),
    1
  );
  const formulas = names.map(name => {
    const color = colorsByName[name] || "#999999";
    return [
      `=IF(RC[-2]<=0,"",SPARKLINE(` +
      `{(RC[-1]/MAX(R${firstDataRow}C[-1]:R${lastDataRow}C[-1]))*${confidenceScale},1},` +
      `{"charttype","bar";"color1","${color}";"color2","white";"max",1}))`
    ];
  });

  tool.getRange(firstDataRow, startCol + 3, names.length, 1)
    .setFormulasR1C1(formulas);
}

function formatAnalysisOutput_(
  tool,
  outRowCount,
  distributionTierNames,
  weightsByTier,
  topTier,
  runnerTier,
  placementComparison,
  fuckpct,
  currentTier,
  canMove,
  totalWeightedOpinions,
  placeMoveFailure,
  reliabilityDistribution,
  labels
) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;
  const meetsRow = labels.findIndex(v => v === "Place/Move");
  const splitRow = labels.findIndex(v => v === "Split");
  const signalHeader = labels.findIndex(v => String(v).trim() === "Fuck % of all");
  const distHeader = labels.findIndex(v => String(v).trim() === "Tier");
  const reliabilityHeader = labels.findIndex(v => String(v).trim() === "Reliability");

  const tierRow = labels.findIndex(v => String(v).trim() === "Tier sheet");
  const levelRow = labels.findIndex(v => String(v).trim() === "Level");

  const idxTotal = labels.findIndex(v => String(v).trim() === "Total weighted opinions");
  const idxMost  = labels.findIndex(v => String(v).trim() === "Most votes (weighted)");
  const idxRun   = labels.findIndex(v => String(v).trim() === "Runner-up (weighted)");
  const idxMean = labels.findIndex(v => String(v).trim() === "Tier Mean");
  const idxMedian = labels.findIndex(v => String(v).trim() === "Tier Median");
  const idxOutliers = labels.findIndex(v => String(v).trim() === "Outliers");
  const idxOutlierRange = labels.findIndex(v => String(v).trim() === "Outlier range");
  const idxSd = labels.findIndex(v => String(v).trim() === "Standard Deviation");
  const idxFuckSignal = labels.findIndex(v => String(v).trim() === "Fuck");
  const styles = createAnalysisStylePlan_(outRowCount);
  const nameToColor = buildTierNameToColor_();

  for (let row = 0; row < outRowCount; row++) {
    styles.fontWeights[row][0] = "bold";
    styles.horizontalAlignments[row][2] = "right";
    styles.horizontalAlignments[row][3] = "right";
  }
  for (let row = 0; row < Math.min(2, outRowCount); row++) {
    fillFormatSpan_(styles.horizontalAlignments, row, 0, 2, "center");
    fillFormatSpan_(styles.fontSizes, row, 0, 2, 11);
  }

  if (tierRow >= 0) {
    const tierColor = nameToColor[String(currentTier || "").trim()];
    fillFormatSpan_(styles.horizontalAlignments, tierRow, 1, 3, "center");
    fillFormatSpan_(styles.fontWeights, tierRow, 1, 3, "bold");
    if (tierColor) {
      fillFormatSpan_(styles.backgrounds, tierRow, 1, 3, tierColor);
      fillFormatSpan_(styles.fontColors, tierRow, 1, 3, tierTextColor_(currentTier));
    }

    const versionRow = tierRow + 2;
    if (versionRow < outRowCount) {
      styles.fontColors[versionRow][3] = "#9e9e9e";
      styles.fontWeights[versionRow][3] = "normal";
      styles.horizontalAlignments[versionRow][3] = "right";
      styles.verticalAlignments[versionRow][3] = "middle";
    }
  }

  if (levelRow >= 0) {
    fillFormatSpan_(styles.backgrounds, levelRow, 1, 3, "#f3f3f3");
    fillFormatSpan_(styles.horizontalAlignments, levelRow, 1, 3, "center");
    fillFormatSpan_(styles.fontWeights, levelRow, 1, 3, "bold");
  }

  [idxTotal, idxMost, idxRun].forEach(index => {
    if (index < 0) return;
    fillFormatRow_(styles.backgrounds, index, "#ffffff");
    styles.horizontalAlignments[index][1] = "left";
  });
  if (idxTotal >= 0) styles.fontWeights[idxTotal][1] = "bold";

  [idxMean, idxMedian, idxOutliers, idxOutlierRange, idxSd].forEach(index => {
    if (index < 0) return;
    fillFormatRow_(styles.backgrounds, index, "#fafafa");
    fillFormatRow_(styles.fontColors, index, "#6f6f6f");
    styles.fontWeights[index][0] = "normal";
  });
  [idxMost, idxRun, idxRun + 1, idxRun + 2].forEach(index => {
    if (index >= 0 && index < outRowCount) styles.horizontalAlignments[index][2] = "left";
  });
  if (idxSd >= 0) styles.horizontalAlignments[idxSd][1] = "left";

  if (meetsRow >= 0) {
    const failureBackground = placeMoveFailure && placeMoveFailure.bg
      ? placeMoveFailure.bg
      : "#fce8e6";
    fillFormatRow_(
      styles.backgrounds,
      meetsRow,
      canMove ? "#e6f4ea" : failureBackground
    );
    styles.fontWeights[meetsRow][1] = "bold";
    if (canMove) {
      styles.horizontalAlignments[meetsRow][3] = "center";
      styles.fontWeights[meetsRow][3] = "bold";
    } else {
      styles.fontWeights[meetsRow][2] = "bold";
    }
  }

  if (splitRow >= 0) {
    fillFormatRow_(styles.backgrounds, splitRow, "#fff4cc");
    fillFormatRow_(styles.verticalAlignments, splitRow, "middle");

    const leftName = placementComparison.left.label;
    if (nameToColor[leftName]) {
      styles.backgrounds[splitRow][1] = nameToColor[leftName];
      styles.fontColors[splitRow][1] = tierTextColor_(leftName);
      styles.horizontalAlignments[splitRow][1] = "right";
    }

    const rightName = placementComparison.right.label;
    if (nameToColor[rightName]) {
      styles.backgrounds[splitRow][2] = nameToColor[rightName];
      styles.fontColors[splitRow][2] = tierTextColor_(rightName);
      styles.horizontalAlignments[splitRow][2] = "left";
    }

    styles.fontSizes[splitRow][3] = 11;
    styles.fontWeights[splitRow][3] = "bold";
    styles.horizontalAlignments[splitRow][3] = "center";
  }

  if (signalHeader >= 0) {
    fillFormatRow_(styles.backgrounds, signalHeader, "#f1f3f4");
    fillFormatRow_(styles.fontWeights, signalHeader, "bold");
    fillFormatRow_(styles.fontColors, signalHeader, "#5f6368");
  }

  if (distHeader >= 0) {
    addWeightedDistributionStyles_(
      styles,
      distHeader,
      distributionTierNames,
      weightsByTier,
      nameToColor
    );

    const firstDataIndex = distHeader + 1;
    const topIndex = distributionTierNames.indexOf(topTier);
    const runnerIndex = distributionTierNames.indexOf(runnerTier);
    if (topIndex >= 0) {
      fillFormatRow_(styles.fontWeights, firstDataIndex + topIndex, "bold");
      fillFormatRow_(styles.fontLines, firstDataIndex + topIndex, "underline");
    }
    if (runnerIndex >= 0) {
      fillFormatRow_(styles.fontStyles, firstDataIndex + runnerIndex, "italic");
    }
  }

  const reliabilityColors = {};
  for (const level of reliabilityDistributionLevels) {
    reliabilityColors[level.name] = level.color;
  }
  if (reliabilityHeader >= 0 && reliabilityDistribution.names.length > 0) {
    addWeightedDistributionStyles_(
      styles,
      reliabilityHeader,
      reliabilityDistribution.names,
      reliabilityDistribution.counts,
      reliabilityColors
    );
  }

  if (idxFuckSignal >= 0) {
    const fuckIsBackground = fuckpct < 0.15;
    fillFormatRow_(styles.backgrounds, idxFuckSignal, "#f5f5f5");
    styles.backgrounds[idxFuckSignal][0] = fuckIsBackground ? "#d9d9d9" : "#000000";
    styles.fontColors[idxFuckSignal][0] = fuckIsBackground ? "#7a7a7a" : "#ff0000";
    styles.fontWeights[idxFuckSignal][0] = fuckIsBackground ? "normal" : "bold";
    styles.fontColors[idxFuckSignal][1] = fuckIsBackground ? "#8a8a8a" : "#000000";
    styles.fontWeights[idxFuckSignal][1] = fuckIsBackground ? "normal" : "bold";
    styles.fontColors[idxFuckSignal][2] = fuckIsBackground ? "#8a8a8a" : "#b71c1c";
    styles.fontWeights[idxFuckSignal][2] = fuckIsBackground ? "normal" : "bold";
    styles.horizontalAlignments[idxFuckSignal][1] = "right";
    styles.horizontalAlignments[idxFuckSignal][2] = "right";
  }

  if (tierRow >= 0) tool.getRange(r0 + tierRow, c0 + 1, 1, 3).breakApart();
  if (levelRow >= 0) tool.getRange(r0 + levelRow, c0 + 1, 1, 3).breakApart();
  if (meetsRow >= 0) tool.getRange(r0 + meetsRow, c0 + 2, 1, 2).breakApart();

  applyAnalysisStylePlan_(tool, r0, c0, styles);

  if (tierRow >= 0) {
    tool.getRange(r0 + tierRow, c0 + 1, 1, 3)
      .mergeAcross()
      .setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
  }
  if (levelRow >= 0) {
    tool.getRange(r0 + levelRow, c0 + 1, 1, 3)
      .mergeAcross()
      .setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
  }
  if (idxRun >= 0 && idxRun + 2 < outRowCount) {
    tool.getRange(r0 + idxRun + 2, c0, 1, OUTPUT_WIDTH)
      .setBorder(false, false, true, false, false, false, "#999999", SpreadsheetApp.BorderStyle.DOTTED);
  }
  if (meetsRow >= 0 && !canMove) {
    tool.getRange(r0 + meetsRow, c0 + 2, 1, 2).mergeAcross();
  }
  if (signalHeader >= 0) {
    tool.getRange(r0 + signalHeader, c0, 1, OUTPUT_WIDTH)
      .setBorder(true, false, true, false, false, false, "#d6d9dc", SpreadsheetApp.BorderStyle.SOLID);
  }
  if (idxFuckSignal >= 0) {
    const fuckIsBackground = fuckpct < 0.15;
    tool.getRange(r0 + idxFuckSignal, c0, 1, OUTPUT_WIDTH)
      .setBorder(true, false, true, false, false, false, "#d0d0d0", SpreadsheetApp.BorderStyle.DASHED);
    tool.getRange(r0 + idxFuckSignal, c0 + 3).setFormulaR1C1(
      `=SPARKLINE({RC[-1],1-RC[-1]},` +
      `{"charttype","bar";"color1","${fuckIsBackground ? "#9e9e9e" : "#000000"}";"color2","#f5f5f5";"max",1})`
    );
    tool.getRange(r0 + idxFuckSignal, c0 + 1).setNumberFormat("0.00");
    tool.getRange(r0 + idxFuckSignal, c0 + 2).setNumberFormat("0.0%");
  }
  if (splitRow >= 0) {
    tool.getRange(r0 + splitRow, c0 + 3).setNumberFormat("0.###");
  }
  [idxMost, idxRun, idxRun + 1, idxRun + 2].forEach(index => {
    if (index >= 0 && index < outRowCount) {
      tool.getRange(r0 + index, c0 + 2).setNumberFormat("0.###");
    }
  });

  if (distHeader >= 0) {
    setWeightedDistributionFormulas_(
      tool,
      r0 + distHeader,
      c0,
      distributionTierNames,
      weightsByTier,
      nameToColor,
      totalWeightedOpinions
    );
  }
  if (reliabilityHeader >= 0 && reliabilityDistribution.names.length > 0) {
    setWeightedDistributionFormulas_(
      tool,
      r0 + reliabilityHeader,
      c0,
      reliabilityDistribution.names,
      reliabilityDistribution.counts,
      reliabilityColors,
      reliabilityDistribution.totalCount
    );
  }
}

function setAnalysisStatusMessage_(tool, message, bg) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  clearAnalysisOutput_(tool, false);

  tool.getRange(r0, c0, 1, OUTPUT_WIDTH)
    .setValues([[message, "", "", ""]])
    .mergeAcross()
    .setBackground(bg || "#fce8e6")
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function clearAnalysisOutput_(tool, resetFormat) {
  const r0 = OUTPUT_START_ROW;
  let lastRow = r0;
  const bottomCell = tool.getRange(tool.getMaxRows(), OUTPUT_COL);

  if (
    typeof bottomCell.getNextDataCell === "function" &&
    SpreadsheetApp.Direction &&
    SpreadsheetApp.Direction.UP
  ) {
    lastRow = Math.max(r0, bottomCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow());
  } else {
    // Compatibility fallback for local mocks and older callers.
    lastRow = Math.max(r0, tool.getLastRow());
  }

  const output = tool.getRange(r0, OUTPUT_COL, lastRow - r0 + 1, OUTPUT_WIDTH);
  output.breakApart().clearContent();
  if (resetFormat) {
    if (typeof output.clearFormat === "function") output.clearFormat();
    if (typeof output.clearNote === "function") output.clearNote();
    if (typeof output.setFontFamily === "function") output.setFontFamily("Mukta");
    if (typeof output.setFontSize === "function") output.setFontSize(10);
  }
}

function renderAnalysisStatus_(tool) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  // If analysis output already exists, do nothing
  const firstLabel = String(tool.getRange(r0, c0).getDisplayValue() || "").trim();
  if (firstLabel === "Tier sheet") {
    return;
  }

  const tierName = String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim();
  const levelName = String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim();

  // Check opinions directly from A:C
  const vals = tool
    .getRange(DATA_START_ROW, 1, tool.getMaxRows() - DATA_START_ROW + 1, 3)
    .getDisplayValues();

  let hasOpinions = false;

  for (const r of vals) {
    if (String(r[0]).trim() || String(r[1]).trim() || String(r[2]).trim()) {
      hasOpinions = true;
      break;
    }
  }

  let message = "";
  let bg = "#fce8e6";

  if (!tierName) {
    message = "Select a tier sheet";
  } else if (!levelName) {
    message = "Select a level";
    bg = "#fff4cc";
  } else if (!hasOpinions) {
    message = "No opinions loaded";
  } else {
    message = "Loading..";
    bg = "#e8f0fe";
  }

  setAnalysisStatusMessage_(tool, message, bg);
}
