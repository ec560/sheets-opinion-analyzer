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

function populateSelectedLevel() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);

  if (typeof loadTierConfiguration_ === "function") {
    const configResult = loadTierConfiguration_();
    if (!configResult.valid && !configResult.missing) {
      setAnalysisStatusMessage_(tool, "Error with Tier Configuration", "#fce8e6");
      return false;
    }
  }

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) {
    renderAnalysisStatus_(tool);
    return false;
  }

  clearAnalysisArea_();
  setAnalysisStatusMessage_(tool, "Loading..", "#e8f0fe");

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) {
    setAnalysisStatusMessage_(tool, "Failed to load opinions", "#fce8e6");
    return false;
  }

  const headers = getLevelHeaders_(tierSheet);
  const header = headers.find(h => h.name === levelName);
  if (!header) {
    setAnalysisStatusMessage_(tool, "Failed to load opinions", "#fce8e6");
    return false;
  }

  const startCol = header.col;
  const numCols = 3; // player/opinion/reliability
  const lastRow = tierSheet.getLastRow();
  const numRows = Math.max(0, lastRow - 1);

  if (numRows === 0) {
    setAnalysisStatusMessage_(tool, "Failed: no opinions found", "#fce8e6");
    return false;
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
    setAnalysisStatusMessage_(tool, "Failed: no opinions found", "#fce8e6");
    return false;
  }

  const dest = tool.getRange(DATA_START_ROW, 1, outVals.length, 3);
  dest.setValues(outVals);
  dest.setBackgrounds(outBgs);
  dest.setFontColors(outFcs);

  try {
    analyzeSelectedLevel();
    return true;
  } catch (e) {
    setAnalysisStatusMessage_(tool, "Failed to analyze opinions", "#fce8e6");
    return false;
  }
}
