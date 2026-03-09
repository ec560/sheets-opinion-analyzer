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
  totalWeightedOpinions,
  passesSplitPct,
  splitMarginPct
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
  const distHeader = labels.findIndex(v => v === "Tier");

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

    tool.getRange(r, c0, 1, OUTPUT_WIDTH)
      .setBackground(val === "YES" ? "#e6f4ea" : "#fce8e6");

    if (val === "YES") {
      const verdictCell = tool.getRange(r, c0 + 3);
      verdictCell
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    }

    tool.getRange(r, c0 + 1).setFontWeight("bold");

    const msgCell = tool.getRange(r, c0 + 3);
    msgCell.setFontWeight("bold");

    if (val !== "YES") {
      msgCell.setValue("");

      // Decide whether we are in "fuck mode"
      let fuckMode = !!(fuckPresent && (fuckpct >= 0.5 || toppct < 0.4));

      // In fuck mode, try the fuck triggers first.
      // If neither trigger is met, fall back to regular logic.
      let ruleA = false, ruleB = false;
      if (fuckMode) {
        ruleA = (fuckpct >= 0.2);     // fuck share trigger (only valid if toppct < 0.4 already)
        ruleB = (sd >= 2);            // volatility trigger

        if (!ruleA && !ruleB) fuckMode = false;
      }

      if (fuckMode) {
        if ((isPending && !passesMajority) || totalWeightedOpinions < 5) {
          msgCell.setValue("Needs more opinions (F)");
        } else if (!verdictDiffersFromCurrent || fuckpct >= 0.5) {
          msgCell.setValue("No movement necessary (F)");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#efefef");
        } else if (!passesSplitMajority && ruleB) {
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
        } else if (!passesSplitPct && verdictDiffersFromCurrent) {
          msgCell.setValue("Low split margin (" + (splitMarginPct * 100).toFixed(2).replace(/\.?0+$/) + "%)");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#ffdfcc");
        } else if (toppct >= 0.75 && totalWeightedOpinions >= 45) {
          msgCell.setValue("🔒 Lock threshold met");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#e6f4ea");
        } else if (!verdictDiffersFromCurrent) {
          msgCell.setValue("No movement necessary");
          tool.getRange(r, c0, 1, OUTPUT_WIDTH).setBackground("#efefef");
        } else {
          msgCell.setValue("Rules not met");
        }
      }
    }
  }

  // Split emphasis
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
      .setNumberFormat("0.###")
      .setHorizontalAlignment("center");
  }

  // Distribution table
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

      const sparkCol = c0 + 3; // H
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

  if (fuckPresent) {
    tool.getRange(r0 + 10, c0, 1, OUTPUT_WIDTH)
      .setBorder(true, false, false, false, false, false, "#999999", SpreadsheetApp.BorderStyle.DASHED);
    tool.getRange(r0 + 10, c0 + 1, 1, 1)
      .setHorizontalAlignment("left")
      .setNumberFormat("0.00");
    tool.getRange(r0 + 11, c0 + 1, 1, 1)
      .setNumberFormat("0.0%")
      .setHorizontalAlignment("left");
  }
}