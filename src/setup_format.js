// setup
function setupTierAnalysis() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(ANALYSIS_SHEET_NAME);

  // Basic layout
  sh.clear();
  sh.getRange("A:H").setFontFamily("Mukta").setFontSize(10);
  sh.getRange("A1").setValue("Tier:");
  sh.getRange("A2").setValue("Level:");
  sh.getRange(TIER_CELL).setNote("Pick a tier sheet");
  sh.getRange(LEVEL_CELL).setNote("Pick a level header from that tier sheet");

  sh.getRange("A3:C3").setValues([["Player", "Opinion", "Reliability"]]);

  // Populate tier dropdown from sheet names (each tier is its own sheet)
  const tierSheetNames = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== ANALYSIS_SHEET_NAME);

  const dvTier = SpreadsheetApp.newDataValidation()
    .requireValueInList(tierSheetNames, true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(TIER_CELL).setDataValidation(dvTier);

  // Clear level validation until tier selected
  sh.getRange(LEVEL_CELL).clearDataValidations();

  sh.setFrozenRows(3);

  formatAnalysisSheetLayout_(sh);
  renderAnalysisStatus_(sh);
}

// format the analysis sheet with column widths, fonts, colors, etc
function formatAnalysisSheetLayout_(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange("A:H").setFontFamily("Mukta").setFontSize(10);

  // Column widths
  sh.setColumnWidths(1, 1, 140); // A
  sh.setColumnWidths(2, 2, 260); // B:C
  sh.setColumnWidths(4, 1, 18);  // D spacer
  sh.setColumnWidths(5, 1, 150); // E output
  sh.setColumnWidths(6, 2, 110); // F:G output
  sh.setColumnWidths(8, 1, 125); // H spark

  // Inputs
  sh.getRange("A1:A2")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  sh.getRange("B1:C2")
    .setBackground("#f3f3f3")
    .setHorizontalAlignment("left");

  // Opinions header
  sh.getRange("A3:C3")
    .setFontWeight("bold")
    .setBackground("#e6f0df")
    .setHorizontalAlignment("center");

  // Output header
  sh.getRange(1, OUTPUT_COL, 1, OUTPUT_WIDTH)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  // ONE thick divider between table and results
  sh.getRange("D1:D200")
    .setBackground("#e0e0e0");

  sh.setFrozenRows(3);

  sh.getRange("D1:D200")
    .setBorder(false, true, false, false, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID_THICK);
}