// dynamic dropdowns and analysis area population based on selected tier/level
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== ANALYSIS_SHEET_NAME) return;

  const a1 = range.getA1Notation();
  if (a1 === TIER_CELL) {
    refreshLevelDropdown_();

    // tier change invalidates the previously selected level
    sh.getRange(LEVEL_CELL).clearContent();

    // also clear pasted data/output whenever tier changes
    clearAnalysisArea_();
    renderAnalysisStatus_(sh);
  }

  if (a1 === LEVEL_CELL) {
    clearAnalysisArea_();
    populateSelectedLevel();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tier Tools")
    .addItem("Setup", "setupTierAnalysis")
    .addItem("Refresh", "populateSelectedLevel")
    .addItem("Analyze Selected Level", "analyzeSelectedLevel")
    .addSeparator()
    .addItem("Scan Tier Flags", "scanSelectedTierFlags")
    .addToUi();
}

// when a level is selected, populate A:C with player/opinion/reliability for that level
function refreshLevelDropdown_() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  if (!tierName) return;

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) return;

  const headers = getLevelHeaders_(tierSheet); // array of {name, col}
  const headerNames = headers.map(h => h.name);

  const dvLevel = SpreadsheetApp.newDataValidation()
    .requireValueInList(headerNames, true)
    .setAllowInvalid(false)
    .build();

  tool.getRange(LEVEL_CELL).setDataValidation(dvLevel);
}

// get level headers from the tier sheet (each level is a header on row 1, occupying 3 columns)
function getLevelHeaders_(tierSheet) {
  // level names are headers on row 1
  // each level occupies 3 columns: [player, opinion, reliability].
  const headerRow = 1;
  const lastCol = tierSheet.getLastColumn();
  const values = tierSheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  const headers = [];
  for (let c = 1; c <= lastCol; c++) {
    const name = (values[c - 1] || "").trim();
    if (!name) continue;

    // treat any non-empty header cell as a level name
    headers.push({ name, col: c });
  }

  // Remove duplicates in case merged headers repeat
  const out = [];
  const seen = new Set();
  for (const h of headers) {
    if (seen.has(h.name)) continue;
    seen.add(h.name);
    out.push(h);
  }
  return out;
}

// clear A:C and output area
function clearAnalysisArea_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ANALYSIS_SHEET_NAME);
  const maxRows = sh.getMaxRows();
  const height = maxRows - (DATA_START_ROW - 1);
  if (height > 0) {
    sh.getRange(DATA_START_ROW, 1, height, 3).clearContent().clearFormat().clearNote();
  }
  sh.getRange(OUTPUT_START_ROW, OUTPUT_COL, sh.getMaxRows(), OUTPUT_WIDTH).clearContent().clearFormat().clearNote();
  formatAnalysisSheetLayout_(sh);
}
