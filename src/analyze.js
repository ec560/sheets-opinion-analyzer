// Main analysis function that reads opinions, applies weights, calculates stats, and outputs results
function analyzeSelectedLevel(options) {
  const ownsTimer = !(options && options.timing);
  const timer = options && options.timing ||
    (typeof createAnalyzerPhaseTimer_ === "function"
      ? createAnalyzerPhaseTimer_({ operation: "analyze-loaded-opinions" })
      : null);
  let timingOutcome = "not-rendered";
  try {
    const result = analyzeSelectedLevelWithTiming_(options, timer);
    timingOutcome = result === true ? "rendered" : "not-rendered";
    return result;
  } catch (error) {
    timingOutcome = "error";
    throw error;
  } finally {
    if (ownsTimer && timer && typeof timer.finish === "function") timer.finish(timingOutcome);
  }
}

function analyzeSelectedLevelWithTiming_(options, timer) {
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "analysisSheetLookup", phaseStartedAt);
  }
  if (!tool) {
    SpreadsheetApp.getUi().alert("Run Tier Tools > Setup before analyzing opinions.");
    return false;
  }

  const preparedValues = options && options.values;
  const preparedBackgrounds = options && options.backgrounds;
  const preparedFontColors = options && options.fontColors;
  const hasPreparedData =
    Array.isArray(preparedValues) &&
    Array.isArray(preparedBackgrounds) &&
    Array.isArray(preparedFontColors) &&
    preparedValues.length === preparedBackgrounds.length &&
    preparedValues.length === preparedFontColors.length;
  const hasValidatedPreparedContext = hasPreparedData &&
    options && options.tierName != null && options.levelName != null;
  let selected;
  if (hasValidatedPreparedContext) {
    selected = { tierName: String(options.tierName).trim(), levelName: String(options.levelName).trim() };
  } else {
    phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    selected = typeof readAnalyzerSelection_ === "function"
      ? readAnalyzerSelection_(tool)
      : {
        tierName: String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim(),
        levelName: String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim()
      };
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "initialAnalysisSelectorRead", phaseStartedAt);
    }
  }
  const selectedTierName = selected.tierName;
  const selectedLevelName = selected.levelName;
  const expectedTierName = options && options.tierName != null
    ? String(options.tierName).trim()
    : selectedTierName;
  const expectedLevelName = options && options.levelName != null
    ? String(options.levelName).trim()
    : selectedLevelName;

  if (expectedTierName !== selectedTierName || expectedLevelName !== selectedLevelName) {
    return false;
  }

  if (timer && typeof timer.setContext === "function") {
    timer.setContext({ tierName: expectedTierName, levelName: expectedLevelName });
  }

  if (typeof loadTierConfiguration_ === "function" &&
      !(hasValidatedPreparedContext && options.configurationValidated)) {
    const configStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    const configResult = loadTierConfiguration_(false, timer);
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "tierConfigurationLoading", configStartedAt);
    }
    if (timer && typeof timer.setContext === "function") {
      timer.setContext({
        tierConfigurationSource: typeof tierConfigurationLoadSource_ === "string"
          ? tierConfigurationLoadSource_
          : "unknown",
        tierConfigurationCacheStatus: typeof tierConfigurationCacheStatus_ === "string"
          ? tierConfigurationCacheStatus_
          : "unknown"
      });
    }
    const configError = tierConfigurationErrorMessage_(configResult);
    if (configError) {
      setAnalysisStatusMessage_(tool, "Error with Tier Configuration", "#fce8e6");
      SpreadsheetApp.getUi().alert(configError);
      return;
    }
  }

  if (!options || !options.outputAlreadyCleared) {
    if (typeof clearAnalysisOutput_ === "function") {
      clearAnalysisOutput_(tool, true);
    } else {
      const lastRow = Math.max(OUTPUT_START_ROW, tool.getLastRow());
      tool.getRange(
        OUTPUT_START_ROW,
        OUTPUT_COL,
        lastRow - OUTPUT_START_ROW + 1,
        OUTPUT_WIDTH
      ).clearContent().breakApart();
    }
  }

  const tierName = expectedTierName;
  const levelName = expectedLevelName;
  if (!tierName || !levelName) {
    SpreadsheetApp.getUi().alert("Pick a tier and a level first.");
    return;
  }
  if (!isTierSheetName_(tierName)) {
    setAnalysisStatusMessage_(tool, "Select a tier sheet", "#fce8e6");
    return false;
  }

  let vals;
  let bgs;
  let fcs;
  if (hasPreparedData) {
    vals = preparedValues;
    bgs = preparedBackgrounds;
    fcs = preparedFontColors;
  } else {
    const lastRow = tool.getLastRow();
    if (lastRow < DATA_START_ROW) {
      SpreadsheetApp.getUi().alert("No opinions pasted yet. Pick the level dropdown again.");
      return;
    }

    const numRows = lastRow - (DATA_START_ROW - 1);
    const rng = tool.getRange(DATA_START_ROW, 1, numRows, 3);
    vals = rng.getValues();
    bgs = rng.getBackgrounds();
    fcs = rng.getFontColors();
  }

  if (timer && typeof timer.setContext === "function") timer.setContext({ opinionCount: vals.length });
  const analysisStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const analysis = calculateLevelAnalysis_(tierName, levelName, vals, bgs, fcs);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "inMemoryAnalysis", analysisStartedAt);
  }
  if (analysis.rawCount === 0) {
    if (!analysisSelectionStillCurrent_(
      tool,
      tierName,
      levelName,
      timer,
      "finalSelectorValidation"
    )) return false;
    applyCountedPlayerHighlights_(tool, DATA_START_ROW, bgs, analysis.countedRowFlags, vals);
    setAnalysisStatusMessage_(tool, "No usable opinions", "#fce8e6");
    return false;
  }

  const {
    weightsByTier,
    countedRowFlags,
    fuckPresent,
    fuckWeight,
    tierWeightSum,
    allWeight,
    totalWeight,
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
    placementComparison,
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
    splitThreshold,
    currentTier,
    moveFailureReason
  } = analysis;
  const assemblyStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const reliabilityDistribution = buildReliabilityDistribution_(bgs, countedRowFlags);
  const placeMoveFailure = canMove ? null : resolvePlaceMoveFailure_({
    isPending,
    passesMajority,
    passesSplitMajority,
    verdictDiffersFromCurrent,
    fuckPresent,
    sd,
    toppct,
    fuckpct,
    minimumOpinionWeight,
    rawOpinionCount: rawCount,
    lockSharePct,
    totalWeightedOpinions: allWeight,
    passesSplitPct,
    verdictBaseName,
    placementComparison,
    moveFailureReason
  });

  const out = [];
  out.push([`Tier sheet`, tierName, "", ""]);
  out.push([`Level`, levelName, "", ""]);
  out.push([`Total weighted opinions`, totalWeight, "", VERSION]);
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
  out.push([
    `Place/Move`,
    canMove ? "YES" : "NO",
    placeMoveFailure ? placeMoveFailure.text : "",
    canMove ? verdictTierName : ""
  ]);

  function upTo3dec_(x) {
    return Number(x).toFixed(3).replace(/\.?0+$/, "");
  }

  out.push([
    `Split`,
    `${placementComparison.left.label}`,
    `${placementComparison.right.label}`,
    `${upTo3dec_(placementComparison.left.weight)} | ${upTo3dec_(placementComparison.right.weight)}`
  ]);

  out.push(["", "", "", ""]);
  if (fuckPresent) {
    out.push(["Fuck % of all", "", "", ""]);
    out.push(["Fuck", fuckWeight, fuckpct, ""]);
    out.push(["", "", "", ""]);
  }
  out.push(["Tier", "Weighted", "%", ""]);
  const distributionTierNames = getVisibleDistributionNames_(orderedTierNames, weightsByTier);
  for (const t of distributionTierNames) {
    const w = weightsByTier[t] || 0;
    const denom = tierWeightSum > 0 ? tierWeightSum : 1;
    const share = w / denom;
    out.push([t, w, share, ""]);
  }

  if (reliabilityDistribution.names.length > 0) {
    out.push(["", "", "", ""]);
    out.push(["Reliability", "Count", "%", ""]);
    for (const name of reliabilityDistribution.names) {
      const count = reliabilityDistribution.counts[name] || 0;
      const denominator = reliabilityDistribution.totalCount || 1;
      out.push([name, count, count / denominator, ""]);
    }
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "analysisResultAssembly", assemblyStartedAt);
  }

  if (!analysisSelectionStillCurrent_(
    tool,
    tierName,
    levelName,
    timer,
    "finalSelectorValidation"
  )) {
    return false;
  }

  const startRow = 1;
  const startCol = OUTPUT_COL;
  const valuesStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  tool.getRange(startRow, startCol, out.length, OUTPUT_WIDTH).setValues(out);
  if (typeof applyAnalysisNumberFormatPlan_ === "function") {
    applyAnalysisNumberFormatPlan_(tool, startRow, startCol, out.length, out.map(row => row[0]));
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "outputValuesAndNumberFormatsWrites", valuesStartedAt);
  }

  const formattingStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  formatAnalysisOutput_(
    tool,
    out.length,
    distributionTierNames,
    weightsByTier,
    topTier,
    secondTier,
    placementComparison,
    fuckpct,
    currentTier,
    canMove,
    allWeight,
    placeMoveFailure,
    reliabilityDistribution,
    out.map(row => row[0])
  );

  applyCountedPlayerHighlights_(tool, DATA_START_ROW, bgs, countedRowFlags, vals);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "outputFormattingAndSparklineWrites", formattingStartedAt);
  }
  return true;
}

function analysisSelectionStillCurrent_(tool, tierName, levelName, timer, phaseName) {
  const phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const selection = typeof readAnalyzerSelection_ === "function"
    ? readAnalyzerSelection_(tool)
    : {
      tierName: String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim(),
      levelName: String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim()
    };
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, phaseName || "analysisSelectorValidation", phaseStartedAt);
  }
  return selection.tierName === tierName && selection.levelName === levelName;
}

function getVisibleDistributionNames_(orderedNames, weightsByName) {
  let firstPositiveIdx = -1;
  let lastPositiveIdx = -1;

  for (let i = 0; i < orderedNames.length; i++) {
    if ((weightsByName[orderedNames[i]] || 0) <= 0) continue;
    if (firstPositiveIdx < 0) firstPositiveIdx = i;
    lastPositiveIdx = i;
  }

  if (firstPositiveIdx < 0) return [];
  return orderedNames.slice(firstPositiveIdx, lastPositiveIdx + 1);
}

function buildReliabilityDistribution_(backgrounds, countedRowFlags) {
  const positiveLevels = reliabilityDistributionLevels.filter(level => {
    return (reliabilityFactors[hex_(level.color)] || 0) > 0;
  });
  const orderedNames = positiveLevels.map(level => level.name);
  const counts = {};
  const nameByColor = {};

  for (const level of positiveLevels) {
    counts[level.name] = 0;
    nameByColor[hex_(level.color)] = level.name;
  }

  for (let row = 0; row < countedRowFlags.length; row++) {
    if (!countedRowFlags[row]) continue;

    const reliabilityColor = hex_(backgrounds[row] && backgrounds[row][2]);
    const reliabilityName = nameByColor[reliabilityColor];
    if (!reliabilityName) continue;

    counts[reliabilityName] += 1;
  }

  const totalCount = orderedNames.reduce((sum, name) => sum + counts[name], 0);
  return {
    names: orderedNames,
    counts,
    totalCount
  };
}
