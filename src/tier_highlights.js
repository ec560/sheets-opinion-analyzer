// A constant true formula tags our overlay without helper cells or saved fills.
const OPINION_VALIDATION_RULE_FORMULA = '=N("TIER_TOOLS_OPINION_VALIDATION_V1")=0';
const OPINION_VALIDATION_RANGES_PER_RULE = 250;
const OPINION_VALIDATION_HIGHLIGHT = "#6fcf97";
const LEGACY_COUNTED_OPINIONS_RULE_FORMULA = '=N("TIER_TOOLS_COUNTED_OPINIONS_V1")=0';

// Keep previously assigned buttons working after the menu action is renamed.
function toggleTierCountedOpinions() {
  return toggleTierOpinionValidation();
}

function toggleTierOpinionValidation() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const title = "Opinion Validation";
  if (!sheet || isAnalyzerUtilitySheetName_(sheet.getName())) {
    ui.alert(title, "Open a tier sheet, then run Toggle Opinion Validation.", ui.ButtonSet.OK);
    return false;
  }

  // Serialize menu clicks by different editors.
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) {
    ss.toast("Another tier highlight update is running. Try again shortly.", title, 5);
    return false;
  }

  let result;
  try {
    result = toggleTierOpinionValidation_(sheet);
    SpreadsheetApp.flush();
  } catch (error) {
    console.error("Could not toggle opinion validation: " + error);
    result = { error: "Could not update opinion validation highlights: " + error.message };
  } finally {
    lock.releaseLock();
  }

  // Alerts suspend Apps Script execution, so show them only after releasing the lock.
  if (result.error) {
    ui.alert(title, result.error, ui.ButtonSet.OK);
    return false;
  }
  ss.toast(result.message, title, 6);
  return true;
}

function isOpinionValidationRule_(rule) {
  const condition = rule.getBooleanCondition();
  return !!condition &&
    condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
    [OPINION_VALIDATION_RULE_FORMULA, LEGACY_COUNTED_OPINIONS_RULE_FORMULA]
      .includes(String(condition.getCriteriaValues()[0]).trim());
}

function toggleTierOpinionValidation_(sheet) {
  const rules = sheet.getConditionalFormatRules();
  const otherRules = rules.filter(rule => !isOpinionValidationRule_(rule));
  const tierName = sheet.getName();
  if (otherRules.length !== rules.length) {
    // Removal requires no analysis. It remains possible after a tier is renamed
    // or its configuration becomes invalid.
    sheet.setConditionalFormatRules(otherRules);
    return { message: 'Opinion validation hidden on "' + tierName + '".' };
  }

  const configResult = loadTierConfiguration_();
  const configError = tierConfigurationErrorMessage_(configResult);
  if (configError) return { error: configError };
  if (!isTierSheetName_(tierName)) {
    return { error: "Open a configured tier sheet, then run Toggle Opinion Validation." };
  }

  const scan = buildTierOpinionValidationRanges_(sheet);
  if (scan.scanned === 0) return { message: 'No level headers found on "' + tierName + '".' };
  const skippedMessage = scan.skippedMerged
    ? " " + scan.skippedMerged + " opinion row(s) skipped because the username cell is merged."
    : "";
  const invalidMessage = " " + scan.unrecognized + " unrecognized opinion color(s).";
  const duplicateMessage = scan.duplicates
    ? " " + scan.duplicates + " duplicate opinion(s) highlighted."
    : "";
  const allRangeGroups = [
    { ranges: scan.duplicateRanges, background: DUPLICATE_PLAYER_HIGHLIGHT },
    { ranges: scan.preUpdateDuplicateRanges, background: PRE_UPDATE_PLAYER_HIGHLIGHT },
    { ranges: scan.ranges, background: OPINION_VALIDATION_HIGHLIGHT }
  ];
  if (!allRangeGroups.some(group => group.ranges.length)) {
    return { message: 'No recognized opinion colors to highlight on "' + tierName + '".' + invalidMessage + skippedMessage };
  }

  const overlays = [];
  // Keep sparse, wide tier sheets to a small number of rules, with bounded lists.
  for (const group of allRangeGroups) {
    for (let i = 0; i < group.ranges.length; i += OPINION_VALIDATION_RANGES_PER_RULE) {
      overlays.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(OPINION_VALIDATION_RULE_FORMULA)
        .setBackground(group.background)
        .setFontColor("#000000")
        .setRanges(group.ranges.slice(i, i + OPINION_VALIDATION_RANGES_PER_RULE))
        .build());
    }
  }
  // Install only after the full scan succeeds. Read existing rules again so edits
  // made to other formatting during the scan are retained, in their original order.
  sheet.setConditionalFormatRules(overlays.concat(sheet.getConditionalFormatRules()));
  return {
    message: scan.highlighted + ' recognized opinion color(s) highlighted on "' + tierName +
      '" across ' + scan.scanned + " level(s)." + duplicateMessage + invalidMessage + skippedMessage
  };
}

function buildTierOpinionValidationRanges_(sheet) {
  const result = {
    scanned: 0,
    highlighted: 0,
    duplicates: 0,
    unrecognized: 0,
    skippedMerged: 0,
    ranges: [],
    duplicateRanges: [],
    preUpdateDuplicateRanges: []
  };
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return result;
  const headers = getLevelHeaders_(sheet);
  result.scanned = headers.length;
  const numRows = Math.max(0, sheet.getLastRow() - 1);
  if (!headers.length || !numRows) return result;

  const source = sheet.getRange(2, 1, numRows, lastCol);
  const vals = source.getValues();
  const bgs = source.getBackgrounds();
  const fcs = source.getFontColors();
  const merged = source.getMergedRanges().map(range => ({
    firstRow: range.getRow(),
    lastRow: range.getRow() + range.getNumRows() - 1,
    firstCol: range.getColumn(),
    lastCol: range.getColumn() + range.getNumColumns() - 1
  }));

  for (const header of headers) {
    const data = extractLevelFlagData_(header, vals, bgs, fcs, lastCol);
    const duplicateHighlights = buildDuplicatePlayerHighlights_(data.vals);
    const playerMerges = merged.filter(merge => header.col >= merge.firstCol && header.col <= merge.lastCol);
    const rows = [];
    const duplicateRows = [];
    const preUpdateDuplicateRows = [];
    data.vals.forEach((value, index) => {
      // Extraction excludes fully blank rows, all recognized reliability types included.
      const row = data.sourceRows[index];
      // skip merged cells in the player column
      if (playerMerges.some(merge => row >= merge.firstRow && row <= merge.lastRow)) {
        result.skippedMerged++;
        return;
      }
      const recognized = isRecognizedOpinionColor_(data.bgs[index][1]);
      if (recognized) result.highlighted++;
      else result.unrecognized++;

      if (duplicateHighlights[index] === DUPLICATE_PLAYER_HIGHLIGHT) {
        duplicateRows.push(row);
        result.duplicates++;
      } else if (duplicateHighlights[index] === PRE_UPDATE_PLAYER_HIGHLIGHT) {
        preUpdateDuplicateRows.push(row);
        result.duplicates++;
      } else if (recognized) {
        rows.push(row);
      }
    });
    for (const run of opinionValidationRowRuns_(rows)) {
      result.ranges.push(sheet.getRange(run.start, header.col, run.length, 1));
    }
    for (const run of opinionValidationRowRuns_(duplicateRows)) {
      result.duplicateRanges.push(sheet.getRange(run.start, header.col, run.length, 1));
    }
    for (const run of opinionValidationRowRuns_(preUpdateDuplicateRows)) {
      result.preUpdateDuplicateRanges.push(sheet.getRange(run.start, header.col, run.length, 1));
    }
  }
  return result;
}

function opinionValidationRowRuns_(rows) {
  const runs = [];
  for (const row of rows) {
    const last = runs[runs.length - 1];
    if (last && row === last.start + last.length) last.length++;
    else runs.push({ start: row, length: 1 });
  }
  return runs;
}
