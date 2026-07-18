function applyTierSheetColor_(tool, r0, c0, outRowCount) {
  const labels = tool.getRange(r0, c0, outRowCount, 1).getDisplayValues().flat();
  const tierSheetRowOff = labels.findIndex(v => String(v).trim() === "Tier sheet");
  if (tierSheetRowOff < 0) return;

  const r = r0 + tierSheetRowOff;
  const tierName = String(tool.getRange(r, c0 + 1).getDisplayValue()).trim(); // value cell
  if (!tierName) return;

  const nameToColor = buildTierNameToColor_();
  const hex = nameToColor[tierName];
  if (!hex) return; // if sheet name isn't a tier name, skip

  tool.getRange(r, c0 + 1)
    .setBackground(hex)
    .setFontColor(tierTextColor_(tierName))
    .setFontWeight("bold");
}

function applyCountedPlayerHighlights_(tool, startRow, sourceBackgrounds, countedRowFlags) {
  if (!sourceBackgrounds || sourceBackgrounds.length === 0) return;

  const fills = sourceBackgrounds.map((row, idx) => {
    return [countedRowFlags && countedRowFlags[idx] ? COUNTED_PLAYER_HIGHLIGHT : "#ffffff"];
  });

  tool.getRange(startRow, 1, fills.length, 1).setBackgrounds(fills);
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

function formatAnalysisOutput_(
  tool,
  outRowCount,
  distributionTierNames,
  weightsByTier,
  topTier,
  topWeight,
  runnerTier,
  placementComparison,
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
  rawOpinionCount,
  lockSharePct,
  totalWeightedOpinions,
  passesSplitPct,
  splitThreshold,
  moveFailureReason,
  reliabilityDistribution
) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  // Reset styling first (important)
  tool.getRange(r0, c0, outRowCount, OUTPUT_WIDTH)
    .setBackground(null)
    .setFontWeight("normal")
    .setFontColor("#000000");

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
  const signalHeader = labels.findIndex(v => String(v).trim() === "Fuck % of all");
  const distHeader = labels.findIndex(v => String(v).trim() === "Tier");
  const reliabilityHeader = labels.findIndex(v => String(v).trim() === "Reliability");

  // Header rows
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

    const versionCell = tool.getRange(r + 2, c0 + 3, 1, 1); // H
    versionCell
      .setFontSize(10)
      .setFontColor("#9e9e9e")
      .setFontWeight("normal")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");
  }

  // Level Header formatting
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
  const idxMean = labels.findIndex(v => String(v).trim() === "Tier Mean");
  const idxMedian = labels.findIndex(v => String(v).trim() === "Tier Median");
  const idxOutliers = labels.findIndex(v => String(v).trim() === "Outliers");
  const idxOutlierRange = labels.findIndex(v => String(v).trim() === "Outlier range");
  const idxSd = labels.findIndex(v => String(v).trim() === "Standard Deviation");
  const idxFuckSignal = labels.findIndex(v => String(v).trim() === "Fuck");

  const styleMetricRow = (r, bg) => {
    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground(bg)

    tool.getRange(r, c0 + 1).setHorizontalAlignment("left");  // tier/text (F)
    tool.getRange(r, c0 + 2, 1, 2).setHorizontalAlignment("right"); // numbers (G:H)

    tool.getRange(r, c0).setFontWeight("bold"); // label emphasis
  };

  if (idxTotal >= 0) {
    const r = r0 + idxTotal;
    styleMetricRow(r, "#ffffff");
    tool.getRange(r, c0 + 1).setFontWeight("bold"); // the total value
  }

  if (idxMost >= 0) styleMetricRow(r0 + idxMost, "#ffffff");
  if (idxRun  >= 0) styleMetricRow(r0 + idxRun,  "#ffffff");

  [idxMean, idxMedian, idxOutliers, idxOutlierRange, idxSd].forEach(idx => {
    if (idx < 0) return;

    tool.getRange(r0 + idx, c0, 1, OUTPUT_WIDTH)
      .setBackground("#fafafa")
      .setFontColor("#6f6f6f");

    tool.getRange(r0 + idx, c0)
      .setFontWeight("normal");
  });

  tool.getRange(r0 + idxMost, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun + 1, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun + 2, c0 + 2)
    .setNumberFormat("0.###")
    .setHorizontalAlignment("left");

  tool.getRange(r0 + idxRun + 2, c0, 1, OUTPUT_WIDTH)
    .setBorder(false, false, true, false, false, false, "#999999", SpreadsheetApp.BorderStyle.DOTTED);

  // Verdict emphasis
  if (meetsRow >= 0) {
    const r = r0 + meetsRow;
    const val = tool.getRange(r, c0 + 1).getDisplayValue().toUpperCase();
    const msgRange = tool.getRange(r, c0 + 2, 1, 2);
    const msgCell = tool.getRange(r, c0 + 2);

    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground(val === "YES" ? "#e6f4ea" : "#fce8e6");

    if (val === "YES") {
      tool.getRange(r, c0 + 2, 1, 2).breakApart();
      const verdictCell = tool.getRange(r, c0 + 3);
      verdictCell
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    }

    tool.getRange(r, c0 + 1).setFontWeight("bold");

    if (val !== "YES") {
      if (!msgRange.isPartOfMerge()) msgRange.mergeAcross();
      msgCell
        .setValue("")
        .setFontWeight("bold");

      const failure = resolvePlaceMoveFailure_({
        isPending,
        passesMajority,
        passesSplitMajority,
        verdictDiffersFromCurrent,
        fuckPresent,
        sd,
        toppct,
        fuckpct,
        minimumOpinionWeight,
        rawOpinionCount,
        lockSharePct,
        totalWeightedOpinions,
        passesSplitPct,
        verdictBaseName,
        placementComparison,
        moveFailureReason
      });

      msgCell.setValue(failure.text);
      if (failure.bg) {
        tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground(failure.bg);
      }
    }
  }

  // Split emphasis
  if (splitRow >= 0) {
    const r = r0 + splitRow;
    const nameToColor = buildTierNameToColor_();

    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground("#fff4cc")
      .setVerticalAlignment("middle");

    if (nameToColor[placementComparison.left.label]) {
      tool.getRange(r, c0 + 1)
        .setBackground(nameToColor[placementComparison.left.label])
        .setFontColor(tierTextColor_(placementComparison.left.label))
        .setHorizontalAlignment("right")
        .setVerticalAlignment("middle");
    }

    if (nameToColor[placementComparison.right.label]) {
      tool.getRange(r, c0 + 2)
        .setBackground(nameToColor[placementComparison.right.label])
        .setFontColor(tierTextColor_(placementComparison.right.label))
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");
    }

    tool.getRange(r, c0 + 3)
      .setFontSize(11)
      .setFontWeight("bold")
      .setNumberFormat("0.###")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  if (signalHeader >= 0) {
    const signalHeaderRow = r0 + signalHeader;
    tool.getRange(signalHeaderRow, c0, 1, OUTPUT_WIDTH)
      .setBackground("#f1f3f4")
      .setFontWeight("bold")
      .setFontColor("#5f6368")
      .setBorder(true, false, true, false, false, false, "#d6d9dc", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Distribution table
  if (distHeader >= 0) {
    const headerRow = r0 + distHeader;
    const nameToColor = buildTierNameToColor_();
    const firstData = formatWeightedDistribution_(
      tool,
      headerRow,
      c0,
      distributionTierNames,
      weightsByTier,
      nameToColor,
      totalWeightedOpinions
    );

    if (idxFuckSignal >= 0) {
      const fuckRow = r0 + idxFuckSignal;
      const sparkCol = c0 + 3; // H
      const sparkCell = tool.getRange(fuckRow, sparkCol);
      const fuckIsBackground = fuckpct < 0.15;

      tool.getRange(fuckRow, c0, 1, OUTPUT_WIDTH)
        .setBackground("#f5f5f5")
        .setBorder(true, false, true, false, false, false, "#d0d0d0", SpreadsheetApp.BorderStyle.DASHED);

      tool.getRange(fuckRow, c0)
        .setBackground(fuckIsBackground ? "#d9d9d9" : "#000000")
        .setFontColor(fuckIsBackground ? "#7a7a7a" : "#ff0000")
        .setFontWeight(fuckIsBackground ? "normal" : "bold");

      tool.getRange(fuckRow, c0 + 1)
        .setFontWeight(fuckIsBackground ? "normal" : "bold")
        .setFontColor(fuckIsBackground ? "#8a8a8a" : "#000000")
        .setHorizontalAlignment("right");

      tool.getRange(fuckRow, c0 + 2)
        .setFontWeight(fuckIsBackground ? "normal" : "bold")
        .setFontColor(fuckIsBackground ? "#8a8a8a" : "#b71c1c")
        .setHorizontalAlignment("right");

      sparkCell
        .setBackground("#f5f5f5");

      sparkCell.setFormulaR1C1(
        `=SPARKLINE({RC[-1],1-RC[-1]},` +
        `{"charttype","bar";"color1","${fuckIsBackground ? "#9e9e9e" : "#000000"}";"color2","#f5f5f5";"max",1})`
      );
    }

    // Emphasize top + runner only
    const tIdx = distributionTierNames.indexOf(topTier);
    const rIdx = distributionTierNames.indexOf(runnerTier);

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

  if (reliabilityHeader >= 0 && reliabilityDistribution.names.length > 0) {
    const reliabilityColors = {};
    for (const level of reliabilityDistributionLevels) {
      reliabilityColors[level.name] = level.color;
    }

    formatWeightedDistribution_(
      tool,
      r0 + reliabilityHeader,
      c0,
      reliabilityDistribution.names,
      reliabilityDistribution.weights,
      reliabilityColors,
      reliabilityDistribution.totalWeight
    );
  }

  if (idxFuckSignal >= 0) {
    tool.getRange(r0 + idxFuckSignal, c0 + 1, 1, 1)
      .setHorizontalAlignment("right")
      .setNumberFormat("0.00");
    tool.getRange(r0 + idxFuckSignal, c0 + 2, 1, 1)
      .setNumberFormat("0.0%")
      .setHorizontalAlignment("right");
  }
}

function formatWeightedDistribution_(
  tool,
  headerRow,
  startCol,
  names,
  weightsByName,
  colorsByName,
  totalWeightedOpinions
) {
  const firstDataRow = headerRow + 1;
  const lastDataRow = firstDataRow + names.length - 1;

  tool.getRange(headerRow, startCol, 1, OUTPUT_WIDTH)
    .setBackground("#eeeeee")
    .setFontWeight("bold");

  for (let i = 0; i < names.length; i++) {
    const row = firstDataRow + i;
    const name = names[i];
    const weight = weightsByName[name] || 0;

    if (weight === 0) {
      tool.getRange(row, startCol, 1, OUTPUT_WIDTH)
        .setFontColor("#9e9e9e");
      continue;
    }

    const color = colorsByName[name] || "#999999";
    tool.getRange(row, startCol)
      .setBackground(color)
      .setFontWeight("bold")
      .setFontColor(tierTextColor_(name));

    const fullWidthAt = 45;
    const minScale = 0.2;
    const confidenceScale = Math.min(
      minScale + (1 - minScale) * (totalWeightedOpinions / fullWidthAt),
      1
    );

    tool.getRange(row, startCol + 3).setFormulaR1C1(
      `=SPARKLINE({(RC[-1]/MAX(R${firstDataRow}C[-1]:R${lastDataRow}C[-1]))*${confidenceScale},1},` +
      `{"charttype","bar";"color1","${color}";"color2","white";"max",1})`
    );
  }

  return firstDataRow;
}

function setAnalysisStatusMessage_(tool, message, bg) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  tool.getRange(r0, c0, 1, OUTPUT_WIDTH).breakApart();
  tool.getRange(r0, c0, tool.getMaxRows(), OUTPUT_WIDTH).clearContent();

  tool.getRange(r0, c0, 1, OUTPUT_WIDTH)
    .setValues([[message, "", "", ""]])
    .mergeAcross()
    .setBackground(bg || "#fce8e6")
    .setFontFamily("Mukta")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function renderAnalysisStatus_(tool) {
  const r0 = OUTPUT_START_ROW;
  const c0 = OUTPUT_COL;

  // If analysis output already exists, do nothing
  const firstLabel = String(tool.getRange(r0, c0).getDisplayValue() || "").trim();
  if (firstLabel === "Tier sheet") {
    return;
  }

  tool.getRange(r0, c0, tool.getMaxRows(), OUTPUT_WIDTH)
    .clearContent();

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
