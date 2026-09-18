// dynamic dropdowns and analysis area population based on selected tier/level
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() === TIER_CONFIG_SHEET_NAME) {
    tierConfigurationLoaded_ = false;
    tierConfigurationResult_ = null;
    refreshTierDropdown_();
    return;
  }
  if (sh.getName() !== ANALYSIS_SHEET_NAME) return;

  const a1 = range.getA1Notation();
  if (a1 === TIER_CELL) {
    refreshTierDropdown_();
    refreshLevelDropdown_();

    // tier change invalidates the previously selected level
    sh.getRange(LEVEL_CELL).clearContent();

    // also clear pasted data/output whenever tier changes
    clearAnalysisArea_();
    renderAnalysisStatus_(sh);
  }

  if (a1 === LEVEL_CELL) {
    populateSelectedLevel({ skipTierDropdownRefresh: true });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tier Tools")
    .addItem("Setup", "setupTierAnalysis")
    .addItem("Refresh", "populateSelectedLevel")
    .addItem("Analyze Selected Level", "analyzeSelectedLevel")
    .addSeparator()
    .addItem("Lock/Unlock Selected Level", "toggleSelectedLevelLock")
    .addSeparator()
    .addItem("Scan Tier Flags", "scanSelectedTierFlags")
    .addItem("Toggle Opinion Validation", "toggleTierOpinionValidation")
    .addToUi();
  refreshTierDropdown_();
}

function refreshTierDropdown_() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!tool) return;

  const names = getTierSheets_(ss).map(sheet => sheet.getName());
  const cell = tool.getRange(TIER_CELL);
  const selectedName = String(cell.getDisplayValue() || "").trim();
  const validation = SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  if (names.length) {
    validation.requireValueInList(names, true);
  } else {
    // Keep an empty selector closed to arbitrary sheet names.
    validation.requireFormulaSatisfied("=FALSE");
  }
  cell.setDataValidation(validation.build());
  cell.setNote(names.length
    ? "Pick a configured tier sheet."
    : "No eligible tier sheets. Check Tier Configuration and sheet names.");

  if (selectedName && !names.includes(selectedName)) {
    cell.clearContent();
    tool.getRange(LEVEL_CELL).clearContent().clearDataValidations();
    clearAnalysisArea_();
    renderAnalysisStatus_(tool);
  }
}

// when a level is selected, populate A:C with player/opinion/reliability for that level
function refreshLevelDropdown_() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  if (!isTierSheetName_(tierName)) return;

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
function getLevelHeaders_(tierSheet, bodyValues) {
  // level names are headers on row 1
  // each level occupies 3 columns: [player, opinion, reliability].
  const headerRow = 1;
  const lastCol = tierSheet.getLastColumn();
  const headerRange = tierSheet.getRange(headerRow, 1, 1, lastCol);
  const values = headerRange.getDisplayValues()[0];
  const canReadMergedRanges = typeof headerRange.getMergedRanges === "function";
  const mergedHeaderWidths = {};

  if (canReadMergedRanges) {
    const mergedRanges = headerRange.getMergedRanges();
    for (const mergedRange of mergedRanges) {
      if (mergedRange.getRow() !== headerRow || mergedRange.getNumRows() !== 1) continue;
      mergedHeaderWidths[mergedRange.getColumn()] = mergedRange.getNumColumns();
    }
  }

  const headers = [];
  for (let c = 1; c <= lastCol; c++) {
    const name = (values[c - 1] || "").trim();
    if (!name) continue;

    // Real level headers span their three data columns. Wider merged headings are
    // section labels, regardless of their text. Keep the fallback for simple
    // mocks or legacy callers that cannot provide merged-range metadata.
    if (canReadMergedRanges && mergedHeaderWidths[c] !== 3) continue;

    headers.push({ name, col: c });
  }

  let levelHeaders = headers;
  if (canReadMergedRanges && typeof tierSheet.getLastRow === "function") {
    const sectionHeaderIndexes = new Set();
    headers.forEach((header, index) => {
      if (isAlphabeticalSectionHeader_(headers, index)) sectionHeaderIndexes.add(index);
    });

    if (sectionHeaderIndexes.size > 0) {
      const bodyRowCount = bodyValues
        ? bodyValues.length
        : Math.max(0, tierSheet.getLastRow() - headerRow);
      levelHeaders = headers.filter((header, index) => {
        if (!sectionHeaderIndexes.has(index)) return true;
        if (bodyValues) return headerGroupHasData_(bodyValues, header.col);
        return headerRangeHasData_(tierSheet, bodyRowCount, header.col);
      });
    }
  }

  // Remove duplicates in case merged headers repeat
  const out = [];
  const seen = new Set();
  for (const h of levelHeaders) {
    if (seen.has(h.name)) continue;
    seen.add(h.name);
    out.push(h);
  }
  return out;
}

function isAlphabeticalSectionHeader_(headers, index) {
  const current = headers[index];
  const next = headers[index + 1];
  const afterNext = headers[index + 2];
  if (!current || !next || !afterNext) return false;

  const currentName = current.name.toLocaleLowerCase();
  const nextName = next.name.toLocaleLowerCase();
  const afterNextName = afterNext.name.toLocaleLowerCase();
  const resetsAfterCurrent = nextName.localeCompare(currentName) < 0;
  const ascendingRunResumes = afterNextName.localeCompare(nextName) >= 0;

  return resetsAfterCurrent && ascendingRunResumes;
}

function headerRangeHasData_(tierSheet, bodyRowCount, startCol) {
  if (bodyRowCount <= 0) return false;
  const values = tierSheet
    .getRange(2, startCol, bodyRowCount, 3)
    .getDisplayValues();
  return headerGroupHasData_(values, 1);
}

function headerGroupHasData_(bodyValues, startCol) {
  const startIndex = startCol - 1;
  for (const row of bodyValues) {
    for (let offset = 0; offset < 3; offset++) {
      if (String(row[startIndex + offset] || "").trim() !== "") return true;
    }
  }
  return false;
}

// clear A:C and output area
function clearAnalysisArea_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ANALYSIS_SHEET_NAME);
  const lastRow = sh.getLastRow();
  const height = Math.max(0, lastRow - DATA_START_ROW + 1);
  if (height > 0) {
    sh.getRange(DATA_START_ROW, 1, height, 3)
      .clearContent()
      .clearFormat()
      .clearNote()
      .setFontFamily("Mukta")
      .setFontSize(10);
  }
  clearAnalysisOutput_(sh, true);
}
