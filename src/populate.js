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

const LEVEL_SOURCE_DIRECT_READ_MAX_ROWS = 2000;

function getLevelLastRow_(tierSheet, startCol, numCols, knownSheetLastRow) {
  const sheetLastRow = knownSheetLastRow == null
    ? tierSheet.getLastRow()
    : knownSheetLastRow;
  if (sheetLastRow <= 1) return sheetLastRow;
  if (
    typeof tierSheet.getMaxRows !== "function" ||
    !SpreadsheetApp.Direction ||
    !SpreadsheetApp.Direction.UP
  ) {
    return sheetLastRow;
  }

  const maxRows = tierSheet.getMaxRows();
  if (sheetLastRow >= maxRows) return sheetLastRow;

  // The row below the sheet-wide last value is guaranteed to be empty. Starting
  // there lets getNextDataCell find each selected column's actual final value.
  const searchRow = sheetLastRow + 1;
  let levelLastRow = 1;
  for (let offset = 0; offset < numCols; offset++) {
    const searchCell = tierSheet.getRange(searchRow, startCol + offset);
    if (typeof searchCell.getNextDataCell !== "function") return sheetLastRow;
    const lastCell = searchCell.getNextDataCell(SpreadsheetApp.Direction.UP);
    if (!lastCell || typeof lastCell.getRow !== "function") return sheetLastRow;
    levelLastRow = Math.max(levelLastRow, lastCell.getRow());
  }
  return Math.min(levelLastRow, sheetLastRow);
}

function lastNonblankSourceRowIndex_(values) {
  for (let row = values.length - 1; row >= 0; row--) {
    for (let col = 0; col < values[row].length; col++) {
      if (String(safeCellValue_(values[row][col])).trim() !== "") return row;
    }
  }
  return -1;
}

function readSelectedLevelSource_(tierSheet, startCol, numCols, timer) {
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const sheetLastRow = tierSheet.getLastRow();
  const candidateRowCount = Math.max(0, sheetLastRow - 1);
  const useDirectRead = candidateRowCount <= LEVEL_SOURCE_DIRECT_READ_MAX_ROWS;
  const levelLastRow = useDirectRead
    ? sheetLastRow
    : getLevelLastRow_(tierSheet, startCol, numCols, sheetLastRow);
  const numRows = Math.max(0, levelLastRow - 1);
  if (timer && typeof timer.setContext === "function") {
    timer.setContext({
      sourceBoundaryStrategy: useDirectRead ? "trim-values-in-memory" : "column-boundary-probe",
      sourceCandidateRowCount: candidateRowCount
    });
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "levelBoundaryDiscovery", phaseStartedAt);
  }
  if (numRows === 0) return { values: [], backgrounds: [], fontColors: [] };

  const sourceRange = tierSheet.getRange(2, startCol, numRows, numCols);
  const sourceReadsStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  let values = sourceRange.getValues();
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "sourceValuesRead", phaseStartedAt);
  }

  let finalRowCount = values.length;
  if (useDirectRead) finalRowCount = lastNonblankSourceRowIndex_(values) + 1;
  if (finalRowCount === 0) {
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "sourceValuesBackgroundsFontColorsReads", sourceReadsStartedAt);
    }
    return { values: [], backgrounds: [], fontColors: [] };
  }

  values = values.slice(0, finalRowCount);
  const formatRange = finalRowCount === numRows
    ? sourceRange
    : tierSheet.getRange(2, startCol, finalRowCount, numCols);

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const backgrounds = formatRange.getBackgrounds();
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "sourceBackgroundsRead", phaseStartedAt);
  }

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const fontColors = formatRange.getFontColors();
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "sourceFontColorsRead", phaseStartedAt);
    endAnalyzerPhase_(timer, "sourceValuesBackgroundsFontColorsReads", sourceReadsStartedAt);
  }
  return { values, backgrounds, fontColors };
}

function populateSelectedLevel(options) {
  const timer = typeof createAnalyzerPhaseTimer_ === "function"
    ? createAnalyzerPhaseTimer_({ operation: "populate-selected-level" })
    : null;
  const lock = typeof LockService !== "undefined" && LockService.getDocumentLock
    ? LockService.getDocumentLock()
    : null;
  const lockStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  let result = false;
  let failedWithError = false;
  let lockAcquired = false;
  let lockPhaseEnded = false;
  try {
    if (lock) {
      lock.waitLock(30000);
      lockAcquired = true;
    }
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "documentLockAcquisition", lockStartedAt);
    }
    lockPhaseEnded = true;
    const preparedOptions = Object.assign({}, options || {}, { timing: timer });
    result = populateSelectedLevelUnlocked_(preparedOptions);
    return result;
  } catch (error) {
    failedWithError = true;
    if (!lockPhaseEnded && typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "documentLockAcquisition", lockStartedAt);
    }
    throw error;
  } finally {
    try {
      if (lockAcquired) {
        const releaseStartedAt = typeof startAnalyzerPhase_ === "function"
          ? startAnalyzerPhase_(timer)
          : Date.now();
        lock.releaseLock();
        if (typeof endAnalyzerPhase_ === "function") {
          endAnalyzerPhase_(timer, "documentLockRelease", releaseStartedAt);
        }
      }
    } finally {
      if (timer && typeof timer.finish === "function") {
        timer.finish(failedWithError ? "error" : result ? "rendered" : "not-rendered");
      }
    }
  }
}

function populateSelectedLevelUnlocked_(options) {
  const timer = options && options.timing;
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "analyzerSheetLookup", phaseStartedAt);
  }
  if (!tool) {
    SpreadsheetApp.getUi().alert("Run Tier Tools > Setup before loading a level.");
    return false;
  }
  if (!options || !options.skipTierDropdownRefresh) refreshTierDropdown_();

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const selection = typeof readAnalyzerSelection_ === "function"
    ? readAnalyzerSelection_(tool)
    : {
      tierName: String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim(),
      levelName: String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim()
    };
  const tierName = selection.tierName;
  const levelName = selection.levelName;
  if (timer && typeof timer.setContext === "function") timer.setContext({ tierName, levelName });

  if (typeof loadTierConfiguration_ === "function") {
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "selectorRead", phaseStartedAt);
    }
    phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    const configResult = loadTierConfiguration_(false, timer);
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "tierConfigurationLoading", phaseStartedAt);
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
    if (!configResult.valid && !configResult.missing) {
      return failSelectedLevelLoad_(
        tool,
        tierName,
        levelName,
        "Error with Tier Configuration",
        "#fce8e6"
      );
    }
  } else if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "selectorRead", phaseStartedAt);
  }

  if (!tierName || !levelName) {
    clearAnalysisArea_(tool);
    renderAnalysisStatus_(tool);
    return false;
  }

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) {
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "headerAndSelectedLevelLookup", phaseStartedAt);
    }
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed to load opinions", "#fce8e6");
  }

  let headers = typeof getLevelHeadersForSelection_ === "function"
    ? getLevelHeadersForSelection_(tierSheet, timer)
    : getLevelHeaders_(tierSheet);
  let header = headers.find(h => h.name === levelName);
  let headerVerificationStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const cachedHeaderMatches = header && selectedLevelHeaderMatches_(tierSheet, header, levelName);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "selectedHeaderVerification", headerVerificationStartedAt);
  }
  if (header && !cachedHeaderMatches) {
    if (typeof invalidateLevelHeaderCache_ === "function") invalidateLevelHeaderCache_(tierSheet);
    headers = typeof getLevelHeadersForSelection_ === "function"
      ? getLevelHeadersForSelection_(tierSheet, timer)
      : getLevelHeaders_(tierSheet);
    header = headers.find(h => h.name === levelName);
    if (typeof levelHeaderLoadSource_ !== "undefined") {
      levelHeaderLoadSource_ = "cache-mismatch-rescan";
    }
  }
  if (timer && typeof timer.setContext === "function") {
    timer.setContext({
      levelHeaderSource: typeof levelHeaderLoadSource_ === "string"
        ? levelHeaderLoadSource_
        : "unknown",
      levelHeaderCacheStatus: typeof levelHeaderCacheStatus_ === "string"
        ? levelHeaderCacheStatus_
        : "unknown",
      levelHeaderLastColumnSource: typeof levelHeaderLastColumnSource_ === "string"
        ? levelHeaderLastColumnSource_
        : "unknown",
      levelHeaderColumnIndexStatus: typeof levelHeaderColumnIndexStatus_ === "string"
        ? levelHeaderColumnIndexStatus_
        : "unknown"
    });
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerAndSelectedLevelLookup", phaseStartedAt);
  }
  if (!header) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed to load opinions", "#fce8e6");
  }

  const startCol = header.col;
  const numCols = 3; // player/opinion/reliability
  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const levelIsLocked = typeof isLevelLocked_ === "function" &&
    isLevelLocked_(tierSheet, startCol);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "selectedLevelLockCheck", phaseStartedAt);
  }
  if (levelIsLocked) {
    return failSelectedLevelLoad_(
      tool,
      tierName,
      levelName,
      "Locked levels cannot be analyzed",
      "#fff4cc"
    );
  }
  const source = readSelectedLevelSource_(tierSheet, startCol, numCols, timer);
  const vals = source.values;
  const bgs = source.backgrounds;
  const fcs = source.fontColors;
  if (vals.length === 0) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed: no opinions found", "#fce8e6");
  }

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
    const displayFontColors = fcs[r].slice();
    displayFontColors[1] = configuredOpinionFontColor_(bgs[r][1], fcs[r][1]);
    outFcs.push(displayFontColors);
  }

  if (outVals.length === 0) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed: no opinions found", "#fce8e6");
  }
  if (timer && typeof timer.setContext === "function") {
    timer.setContext({ opinionCount: outVals.length });
  }

  if (!selectedLevelStillCurrent_(tool, tierName, levelName, timer, "preWriteSelectorValidation")) {
    return false;
  }

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  clearAnalysisArea_(tool, timer);

  const dest = tool.getRange(DATA_START_ROW, 1, outVals.length, 3);
  let writeStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  dest.setValues(outVals);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "loadedOpinionValuesWrite", writeStartedAt);
  }
  writeStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  dest.setBackgrounds(outBgs);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "loadedOpinionBackgroundsWrite", writeStartedAt);
  }
  writeStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  dest.setFontColors(outFcs);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "loadedOpinionFontColorsWrite", writeStartedAt);
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "loadedOpinionsClearAndWrite", phaseStartedAt);
  }

  try {
    return analyzeSelectedLevel({
      tierName,
      levelName,
      configurationValidated: true,
      outputAlreadyCleared: true,
      values: outVals,
      backgrounds: outBgs,
      fontColors: outFcs,
      timing: timer
    }) !== false;
  } catch (e) {
    setAnalysisStatusMessage_(tool, "Failed to analyze opinions", "#fce8e6");
    return false;
  }
}

function selectedLevelHeaderMatches_(tierSheet, header, expectedName) {
  if (!tierSheet || !header) return false;
  try {
    const actualName = String(tierSheet.getRange(1, header.col).getDisplayValue() || "").trim();
    return actualName === expectedName;
  } catch (error) {
    // Lightweight callers that cannot expose a single-cell header read still
    // receive the same discovery behavior as before caching was introduced.
    return true;
  }
}

function selectedLevelStillCurrent_(tool, tierName, levelName, timer, phaseName) {
  if (!tool) return false;
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
    endAnalyzerPhase_(timer, phaseName || "selectorValidation", phaseStartedAt);
  }
  return selection.tierName === tierName && selection.levelName === levelName;
}

function failSelectedLevelLoad_(tool, tierName, levelName, message, background) {
  if (!selectedLevelStillCurrent_(tool, tierName, levelName)) return false;
  clearAnalysisArea_(tool);
  setAnalysisStatusMessage_(tool, message, background);
  return false;
}
