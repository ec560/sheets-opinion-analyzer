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

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) return;

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) return;

  const headers = getLevelHeaders_(tierSheet);
  const header = headers.find(h => h.name === levelName);
  if (!header) return;

  const startCol = header.col;
  const numCols = 3; // player/opinion/reliability
  const lastRow = tierSheet.getLastRow();
  const numRows = Math.max(0, lastRow - 1);

  if (numRows === 0) return;

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
    outFcs.push(fcs[r]);
  }

  if (outVals.length === 0) return;

  const dest = tool.getRange(DATA_START_ROW, 1, outVals.length, 3);
  dest.setValues(outVals);
  dest.setBackgrounds(outBgs);
  dest.setFontColors(outFcs);

  renderAnalysisStatus_(tool);
}