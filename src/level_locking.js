const LEVEL_LOCK_BACKGROUND = "#000000";
const LEVEL_UNLOCK_BACKGROUND = "#ffffff";
const LEVEL_UNLOCK_FONT_COLOR = "#000000";

function toggleSelectedLevelLock() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const target = getSelectedLevelForLocking_(ss);

  if (!target) {
    ui.alert(
      "Level Lock",
      "Open a tier sheet and select any cell within the level's three-column block.",
      ui.ButtonSet.OK
    );
    return false;
  }

  return setSelectedLevelLockState_(!isLevelLocked_(target.sheet, target.startCol));
}

function lockSelectedLevel() {
  return setSelectedLevelLockState_(true);
}

function unlockSelectedLevel() {
  return setSelectedLevelLockState_(false);
}

function setSelectedLevelLockState_(shouldLock) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const target = getSelectedLevelForLocking_(ss);

  if (!target) {
    ui.alert(
      "Level Lock",
      "Open a tier sheet and select any cell within the level's three-column block.",
      ui.ButtonSet.OK
    );
    return false;
  }

  const currentlyLocked = isLevelLocked_(target.sheet, target.startCol);
  if (currentlyLocked === shouldLock) {
    ss.toast(
      "\"" + target.name + "\" is already " + (shouldLock ? "locked." : "unlocked."),
      "Tier Tools",
      3
    );
    return false;
  }

  const segments = getLevelLockSegments_(target.sheet, target.startCol, target.sheet.getMaxRows());
  const failedSegments = applyLevelLockSegments_(segments, shouldLock);
  SpreadsheetApp.flush();

  if (failedSegments.length > 0) {
    ui.alert(
      "Level Lock",
      (shouldLock ? "Locked" : "Unlocked") + " the level, but could not restyle merged row(s) " +
        failedSegments.join(", ") + ". The rows above and below them were still updated.",
      ui.ButtonSet.OK
    );
    return false;
  }

  ss.toast(
    (shouldLock ? "Locked" : "Unlocked") + " \"" + target.name + "\".",
    "Tier Tools",
    3
  );
  return true;
}

function getSelectedLevelForLocking_(ss) {
  const sheet = ss.getActiveSheet();
  if (!sheet || isAnalyzerUtilitySheetName_(sheet.getName())) return null;

  const activeCell = sheet.getActiveCell();
  if (!activeCell) return null;

  const selectedCol = activeCell.getColumn();
  const header = getLevelHeaders_(sheet).find(item => {
    return selectedCol >= item.col && selectedCol <= item.col + 2;
  });

  if (!header) return null;
  return {
    sheet,
    name: header.name,
    startCol: header.col
  };
}

function isLevelLocked_(sheet, startCol) {
  return hex_(sheet.getRange(1, startCol).getBackground()) === LEVEL_LOCK_BACKGROUND;
}

function getLevelLockSegments_(sheet, startCol, maxRows) {
  const fullRange = sheet.getRange(1, startCol, maxRows, 3);
  if (typeof fullRange.getMergedRanges !== "function") {
    return [{ range: fullRange, isMerged: false }];
  }

  const mergedRows = fullRange.getMergedRanges()
    .filter(range => range.getRow() >= 1 && range.getRow() <= maxRows)
    .sort((a, b) => a.getRow() - b.getRow());

  if (mergedRows.length === 0) {
    return [{ range: fullRange, isMerged: false }];
  }

  const segments = [];
  let nextRow = 1;
  for (const mergedRange of mergedRows) {
    const mergeStart = mergedRange.getRow();
    const mergeEnd = Math.min(maxRows, mergeStart + mergedRange.getNumRows() - 1);
    if (mergeStart < nextRow) continue;

    if (mergeStart > nextRow) {
      segments.push({
        range: sheet.getRange(nextRow, startCol, mergeStart - nextRow, 3),
        isMerged: false
      });
    }
    const mergeStartsAtLevel = mergedRange.getColumn() === startCol;
    const mergeEndsAtLevel = mergedRange.getColumn() + mergedRange.getNumColumns() === startCol + 3;
    segments.push({
      range: mergeStartsAtLevel && mergeEndsAtLevel
        ? mergedRange
        : sheet.getRange(mergeStart, startCol, mergeEnd - mergeStart + 1, 3),
      isMerged: true,
      row: mergeStart
    });
    nextRow = mergeEnd + 1;
  }

  if (nextRow <= maxRows) {
    segments.push({
      range: sheet.getRange(nextRow, startCol, maxRows - nextRow + 1, 3),
      isMerged: false
    });
  }
  return segments;
}

function applyLevelLockSegments_(segments, shouldLock) {
  const failedRows = [];
  const orderedSegments = segments.slice().sort((a, b) => Number(a.isMerged) - Number(b.isMerged));
  for (const segment of orderedSegments) {
    try {
      if (shouldLock) {
        segment.isMerged ? lockMergedLevelRange_(segment.range) : lockLevelRange_(segment.range);
      } else {
        segment.isMerged ? unlockMergedLevelRange_(segment.range) : unlockLevelRange_(segment.range);
      }
    } catch (error) {
      const row = segment.row || (typeof segment.range.getRow === "function" ? segment.range.getRow() : "unknown");
      failedRows.push(row);
      console.error("Could not update merged level row " + row + ": " + error);
    }
  }
  return failedRows;
}

function lockLevelRange_(range) {
  const originalBackgrounds = normalizeColorMatrix_(range.getBackgrounds());
  range.setFontColors(originalBackgrounds);
  range.setBackground(LEVEL_LOCK_BACKGROUND);
}

function unlockLevelRange_(range) {
  const storedBackgrounds = normalizeColorMatrix_(range.getFontColors());
  const displayValues = range.getDisplayValues();
  const restoredBackgrounds = buildUnlockedLevelBackgrounds_(storedBackgrounds, displayValues);
  range.setBackgrounds(restoredBackgrounds);
  range.setFontColors(buildUnlockedLevelFontColors_(restoredBackgrounds));
}

function lockMergedLevelRange_(range) {
  const originalBackground = hex_(range.getBackground());
  range.setFontColor(originalBackground);
  range.setBackground(LEVEL_LOCK_BACKGROUND);
}

function unlockMergedLevelRange_(range) {
  const hasVisibleValue = String(range.getDisplayValue() || "").trim() !== "";
  const restoredBackground = hasVisibleValue
    ? hex_(range.getFontColor())
    : LEVEL_UNLOCK_BACKGROUND;
  range.setBackground(restoredBackground);
  range.setFontColor(configuredOpinionFontColor_(restoredBackground, LEVEL_UNLOCK_FONT_COLOR));
}

function normalizeColorMatrix_(colors) {
  return colors.map(row => row.map(hex_));
}

function buildUnlockedLevelBackgrounds_(storedBackgrounds, displayValues) {
  return storedBackgrounds.map((row, rowIndex) => {
    return row.map((background, colIndex) => {
      const value = displayValues[rowIndex] && displayValues[rowIndex][colIndex];
      return String(value || "").trim() === "" ? LEVEL_UNLOCK_BACKGROUND : background;
    });
  });
}

function buildUnlockedLevelFontColors_(backgrounds) {
  return backgrounds.map((row, rowIndex) => {
    return row.map((background, colIndex) => {
      if (rowIndex === 0 || colIndex === 1) {
        return configuredOpinionFontColor_(background, LEVEL_UNLOCK_FONT_COLOR);
      }
      return LEVEL_UNLOCK_FONT_COLOR;
    });
  });
}
