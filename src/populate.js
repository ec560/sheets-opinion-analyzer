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

function getLevelLastRow_(tierSheet, startCol, numCols) {
  const sheetLastRow = tierSheet.getLastRow();
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

function populateSelectedLevel(options) {
  const lock = typeof LockService !== "undefined" && LockService.getDocumentLock
    ? LockService.getDocumentLock()
    : null;
  if (lock) lock.waitLock(30000);

  try {
    return populateSelectedLevelUnlocked_(options);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function populateSelectedLevelUnlocked_(options) {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!tool) {
    SpreadsheetApp.getUi().alert("Run Tier Tools > Setup before loading a level.");
    return false;
  }
  if (!options || !options.skipTierDropdownRefresh) refreshTierDropdown_();

  const tierName = String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim();
  const levelName = String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim();

  if (typeof loadTierConfiguration_ === "function") {
    const configResult = loadTierConfiguration_();
    if (!configResult.valid && !configResult.missing) {
      return failSelectedLevelLoad_(
        tool,
        tierName,
        levelName,
        "Error with Tier Configuration",
        "#fce8e6"
      );
    }
  }

  if (!tierName || !levelName) {
    clearAnalysisArea_(tool);
    renderAnalysisStatus_(tool);
    return false;
  }

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed to load opinions", "#fce8e6");
  }

  const headers = getLevelHeaders_(tierSheet);
  const header = headers.find(h => h.name === levelName);
  if (!header) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed to load opinions", "#fce8e6");
  }

  const startCol = header.col;
  const numCols = 3; // player/opinion/reliability
  if (typeof isLevelLocked_ === "function" && isLevelLocked_(tierSheet, startCol)) {
    return failSelectedLevelLoad_(
      tool,
      tierName,
      levelName,
      "Locked levels cannot be analyzed",
      "#fff4cc"
    );
  }
  const lastRow = getLevelLastRow_(tierSheet, startCol, numCols);
  const numRows = Math.max(0, lastRow - 1);

  if (numRows === 0) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed: no opinions found", "#fce8e6");
  }

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
    const displayFontColors = fcs[r].slice();
    displayFontColors[1] = configuredOpinionFontColor_(bgs[r][1], fcs[r][1]);
    outFcs.push(displayFontColors);
  }

  if (outVals.length === 0) {
    return failSelectedLevelLoad_(tool, tierName, levelName, "Failed: no opinions found", "#fce8e6");
  }

  if (!selectedLevelStillCurrent_(tool, tierName, levelName)) {
    return false;
  }

  clearAnalysisArea_(tool);

  const dest = tool.getRange(DATA_START_ROW, 1, outVals.length, 3);
  dest.setValues(outVals);
  dest.setBackgrounds(outBgs);
  dest.setFontColors(outFcs);

  try {
    return analyzeSelectedLevel({
      tierName,
      levelName,
      outputAlreadyCleared: true,
      values: outVals,
      backgrounds: outBgs,
      fontColors: outFcs
    }) !== false;
  } catch (e) {
    setAnalysisStatusMessage_(tool, "Failed to analyze opinions", "#fce8e6");
    return false;
  }
}

function selectedLevelStillCurrent_(tool, tierName, levelName) {
  if (!tool) return false;
  const currentTierName = String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim();
  const currentLevelName = String(tool.getRange(LEVEL_CELL).getDisplayValue() || "").trim();
  return currentTierName === tierName && currentLevelName === levelName;
}

function failSelectedLevelLoad_(tool, tierName, levelName, message, background) {
  if (!selectedLevelStillCurrent_(tool, tierName, levelName)) return false;
  clearAnalysisArea_(tool);
  setAnalysisStatusMessage_(tool, message, background);
  return false;
}
