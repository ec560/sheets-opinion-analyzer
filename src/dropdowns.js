// dynamic dropdowns and analysis area population based on selected tier/level
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sh = range.getSheet();
  const sheetName = sh.getName();
  const editedFirstRow = typeof range.getRow === "function" &&
    range.getRow() <= 1 &&
    range.getRow() + (typeof range.getNumRows === "function" ? range.getNumRows() : 1) - 1 >= 1;
  const sheetCanContainLevelHeaders = typeof isAnalyzerUtilitySheetName_ !== "function" ||
    !isAnalyzerUtilitySheetName_(sheetName);
  if (editedFirstRow && sheetCanContainLevelHeaders &&
      typeof invalidateLevelHeaderCache_ === "function") {
    invalidateLevelHeaderCache_(sh);
  }
  if (sheetName === TIER_CONFIG_SHEET_NAME) {
    if (typeof invalidateTierConfigurationCache_ === "function") {
      invalidateTierConfigurationCache_();
    } else {
      tierConfigurationLoaded_ = false;
      tierConfigurationResult_ = null;
    }
    refreshTierDropdown_();
    return;
  }
  if (sheetName !== ANALYSIS_SHEET_NAME) return;

  const a1 = range.getA1Notation();
  if (a1 === TIER_CELL) {
    // tier change invalidates the previously selected level
    sh.getRange(LEVEL_CELL).clearContent().clearDataValidations();

    // Clear stale data before any potentially slow dropdown discovery. This
    // also guarantees the previous tier's output is removed if refresh fails.
    clearAnalysisArea_(sh);

    const selectedTierName = e.value != null
      ? String(e.value).trim()
      : String(sh.getRange(TIER_CELL).getDisplayValue() || "").trim();
    refreshLevelDropdown_(selectedTierName);
    renderAnalysisStatus_(sh);
    return;
  }

  if (a1 === LEVEL_CELL) {
    populateSelectedLevel({ skipTierDropdownRefresh: true });
    return;
  }

  if (rangeIntersectsLoadedOpinions_(range)) {
    setAnalysisStatusMessage_(
      sh,
      "Opinions changed. Run Analyze Loaded Opinions",
      "#fff4cc"
    );
  }
}

function rangeIntersectsLoadedOpinions_(range) {
  if (!range || typeof range.getRow !== "function" || typeof range.getColumn !== "function") {
    return false;
  }

  const row = range.getRow();
  const col = range.getColumn();
  const numRows = typeof range.getNumRows === "function" ? range.getNumRows() : 1;
  const numCols = typeof range.getNumColumns === "function" ? range.getNumColumns() : 1;
  const lastRow = row + numRows - 1;
  const lastCol = col + numCols - 1;

  return lastRow >= DATA_START_ROW && col <= 3 && lastCol >= 1;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tier Tools")
    .addItem("Setup", "setupTierAnalysis")
    .addItem("Reload Selected Level", "populateSelectedLevel")
    .addItem("Analyze Loaded Opinions", "analyzeSelectedLevel")
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
function refreshLevelDropdown_(selectedTierName) {
  const timer = typeof createAnalyzerPhaseTimer_ === "function"
    ? createAnalyzerPhaseTimer_({ operation: "refresh-level-dropdown" })
    : null;
  let outcome = "not-rendered";
  try {
    const refreshed = refreshLevelDropdownWithTiming_(selectedTierName, timer);
    outcome = refreshed ? "rendered" : "not-rendered";
    return refreshed;
  } catch (error) {
    outcome = "error";
    throw error;
  } finally {
    if (timer && typeof timer.finish === "function") timer.finish(outcome);
  }
}

function refreshLevelDropdownWithTiming_(selectedTierName, timer) {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!tool) return false;
  const levelCell = tool.getRange(LEVEL_CELL);
  levelCell.clearDataValidations();
  const tierName = selectedTierName == null
    ? String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim()
    : String(selectedTierName).trim();
  if (timer && typeof timer.setContext === "function") timer.setContext({ tierName });
  if (!isTierSheetName_(tierName)) return false;

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet) return false;

  const discoveredHeaders = getLevelHeadersForSelection_(tierSheet, timer);
  const lockReadStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const headers = getUnlockedLevelHeadersForDropdown_(tierSheet, discoveredHeaders);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "levelLockBackgroundsRead", lockReadStartedAt);
  }
  const headerNames = headers.map(h => h.name);
  if (timer && typeof timer.setContext === "function") {
    timer.setContext({
      discoveredLevelCount: discoveredHeaders.length,
      unlockedLevelCount: headers.length,
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

  const validationStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const dvLevel = SpreadsheetApp.newDataValidation()
    .requireValueInList(headerNames, true)
    .setAllowInvalid(false)
    .build();

  levelCell.setDataValidation(dvLevel);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "levelValidationInstall", validationStartedAt);
  }
  return true;
}

function getUnlockedLevelHeadersForDropdown_(tierSheet, headers) {
  if (!headers.length) return [];
  const firstColumn = Math.min(...headers.map(header => header.col));
  const lastColumn = Math.max(...headers.map(header => header.col));
  const backgrounds = tierSheet
    .getRange(1, firstColumn, 1, lastColumn - firstColumn + 1)
    .getBackgrounds()[0];
  const isFuckTier = String(tierSheet.getName() || "").toLowerCase() === "fuck";

  return headers.filter(header => {
    const background = typeof hex_ === "function"
      ? hex_(backgrounds[header.col - firstColumn])
      : String(backgrounds[header.col - firstColumn] || "").toLowerCase();
    if (background === "#010000") return false;
    if (background !== "#000000") return true;
    return isFuckTier;
  });
}

// get level headers from the tier sheet (each level is a header on row 1, occupying 3 columns)
function getLevelHeaders_(tierSheet, bodyValues, timer, knownLastColumn) {
  // level names are headers on row 1
  // each level occupies 3 columns: [player, opinion, reliability].
  const headerRow = 1;
  const lastCol = knownLastColumn == null ? tierSheet.getLastColumn() : knownLastColumn;
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const headerRange = tierSheet.getRange(headerRow, 1, 1, lastCol);
  const values = headerRange.getDisplayValues()[0];
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerRowValuesRead", phaseStartedAt);
  }
  const canReadMergedRanges = typeof headerRange.getMergedRanges === "function";
  const mergedHeaderWidths = {};

  if (canReadMergedRanges) {
    phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    const mergedRanges = headerRange.getMergedRanges();
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "headerMergedRangesRead", phaseStartedAt);
    }
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
      let bodyRowCount = bodyValues ? bodyValues.length : 0;
      if (!bodyValues) {
        phaseStartedAt = typeof startAnalyzerPhase_ === "function"
          ? startAnalyzerPhase_(timer)
          : Date.now();
        bodyRowCount = Math.max(0, tierSheet.getLastRow() - headerRow);
        if (typeof endAnalyzerPhase_ === "function") {
          endAnalyzerPhase_(timer, "headerSectionLastRowRead", phaseStartedAt);
        }
      }
      let sectionBodyValues = bodyValues;
      let sectionBodyStartCol = 1;
      if (!sectionBodyValues && bodyRowCount > 0) {
        const sectionHeaders = headers.filter((header, index) => sectionHeaderIndexes.has(index));
        sectionBodyStartCol = Math.min(...sectionHeaders.map(header => header.col));
        const sectionBodyEndCol = Math.max(...sectionHeaders.map(header => header.col + 2));
        phaseStartedAt = typeof startAnalyzerPhase_ === "function"
          ? startAnalyzerPhase_(timer)
          : Date.now();
        sectionBodyValues = tierSheet
          .getRange(
            headerRow + 1,
            sectionBodyStartCol,
            bodyRowCount,
            sectionBodyEndCol - sectionBodyStartCol + 1
          )
          .getDisplayValues();
        if (typeof endAnalyzerPhase_ === "function") {
          endAnalyzerPhase_(timer, "headerSectionBodyRead", phaseStartedAt);
        }
      }
      levelHeaders = headers.filter((header, index) => {
        if (!sectionHeaderIndexes.has(index)) return true;
        if (!sectionBodyValues) return false;
        return headerGroupHasData_(
          sectionBodyValues,
          header.col - sectionBodyStartCol + 1
        );
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

const LEVEL_HEADER_CACHE_SCHEMA = 1;
const LEVEL_HEADER_CACHE_TTL_SECONDS = 21600;
const LEVEL_HEADER_COLUMN_INDEX_TTL_SECONDS = 300;
let levelHeaderLoadSource_ = "not-loaded";
let levelHeaderCacheStatus_ = "not-checked";
let levelHeaderLastColumnSource_ = "not-loaded";
let levelHeaderColumnIndexStatus_ = "not-checked";

function getLevelHeaderCache_() {
  if (typeof CacheService === "undefined" || !CacheService.getDocumentCache) return null;
  try {
    return CacheService.getDocumentCache();
  } catch (error) {
    return null;
  }
}

function levelHeaderCacheIdentity_(tierSheet) {
  let spreadsheetId = "active";
  const sheetId = typeof tierSheet.getSheetId === "function" ? tierSheet.getSheetId() : tierSheet.getName();
  try {
    const parent = typeof tierSheet.getParent === "function" ? tierSheet.getParent() : null;
    if (parent && typeof parent.getId === "function") spreadsheetId = parent.getId();
    else if (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getActive) {
      const active = SpreadsheetApp.getActive();
      if (active && typeof active.getId === "function") spreadsheetId = active.getId();
    }
  } catch (error) {}
  return spreadsheetId + ":" + sheetId;
}

function levelHeaderCacheKey_(tierSheet, lastColumn) {
  return "analyzer:level-headers:v" + LEVEL_HEADER_CACHE_SCHEMA + ":" +
    levelHeaderCacheIdentity_(tierSheet) + ":" + lastColumn;
}

function levelHeaderColumnIndexKey_(tierSheet) {
  return "analyzer:level-header-column:v" + LEVEL_HEADER_CACHE_SCHEMA + ":" +
    levelHeaderCacheIdentity_(tierSheet);
}

function parseCachedLevelHeaderColumn_(serialized) {
  if (!serialized) return null;
  try {
    const cached = JSON.parse(serialized);
    return cached.schema === LEVEL_HEADER_CACHE_SCHEMA &&
      Number.isInteger(cached.lastColumn) && cached.lastColumn > 0
      ? cached.lastColumn
      : null;
  } catch (error) {
    return null;
  }
}

function cacheLevelHeaderColumnIndex_(cache, tierSheet, lastColumn) {
  if (!cache) return;
  cache.put(levelHeaderColumnIndexKey_(tierSheet), JSON.stringify({
    schema: LEVEL_HEADER_CACHE_SCHEMA,
    lastColumn
  }), LEVEL_HEADER_COLUMN_INDEX_TTL_SECONDS);
}

function parseCachedLevelHeaders_(serialized, lastColumn) {
  if (!serialized) return null;
  try {
    const cached = JSON.parse(serialized);
    if (cached.schema !== LEVEL_HEADER_CACHE_SCHEMA || cached.lastColumn !== lastColumn) return null;
    if (!Array.isArray(cached.headers)) return null;
    const seenNames = {};
    for (const header of cached.headers) {
      if (!header || typeof header.name !== "string" || !header.name.trim()) return null;
      if (!Number.isInteger(header.col) || header.col < 1 || header.col > lastColumn) return null;
      if (seenNames[header.name]) return null;
      seenNames[header.name] = true;
    }
    return cached.headers;
  } catch (error) {
    return null;
  }
}

function cacheLevelHeaders_(tierSheet, lastColumn, headers) {
  const cache = getLevelHeaderCache_();
  if (!cache) return;
  try {
    cache.put(levelHeaderCacheKey_(tierSheet, lastColumn), JSON.stringify({
      schema: LEVEL_HEADER_CACHE_SCHEMA,
      lastColumn,
      headers
    }), LEVEL_HEADER_CACHE_TTL_SECONDS);
    cacheLevelHeaderColumnIndex_(cache, tierSheet, lastColumn);
  } catch (error) {}
}

function readCachedLevelHeaders_(cache, tierSheet, lastColumn, timer) {
  const phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  let headers = null;
  try {
    const key = levelHeaderCacheKey_(tierSheet, lastColumn);
    const serialized = cache.get(key);
    headers = parseCachedLevelHeaders_(serialized, lastColumn);
    levelHeaderCacheStatus_ = headers ? "hit" : serialized ? "invalid" : "miss";
    if (!headers && serialized) cache.remove(key);
  } catch (error) {
    levelHeaderCacheStatus_ = "error";
  }
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerCacheLookup", phaseStartedAt);
  }
  return headers;
}

function getLevelHeadersForSelection_(tierSheet, timer) {
  const cache = getLevelHeaderCache_();
  levelHeaderCacheStatus_ = cache ? "not-checked" : "unavailable";
  levelHeaderColumnIndexStatus_ = cache ? "not-checked" : "unavailable";
  let lastColumn = null;

  if (cache) {
    let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
      ? startAnalyzerPhase_(timer)
      : Date.now();
    try {
      const serialized = cache.get(levelHeaderColumnIndexKey_(tierSheet));
      lastColumn = parseCachedLevelHeaderColumn_(serialized);
      levelHeaderColumnIndexStatus_ = lastColumn
        ? "hit"
        : serialized ? "invalid" : "miss";
    } catch (error) {
      levelHeaderColumnIndexStatus_ = "error";
    }
    if (typeof endAnalyzerPhase_ === "function") {
      endAnalyzerPhase_(timer, "headerColumnIndexLookup", phaseStartedAt);
    }
    if (lastColumn) {
      levelHeaderLastColumnSource_ = "cache-index";
      const indexedHeaders = readCachedLevelHeaders_(cache, tierSheet, lastColumn, timer);
      if (indexedHeaders) {
        levelHeaderLoadSource_ = "document-cache";
        return indexedHeaders;
      }
      try { cache.remove(levelHeaderColumnIndexKey_(tierSheet)); } catch (error) {}
      lastColumn = null;
    }
  }

  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  lastColumn = tierSheet.getLastColumn();
  levelHeaderLastColumnSource_ = "sheet";
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerLastColumnRead", phaseStartedAt);
  }

  if (cache) {
    const headers = readCachedLevelHeaders_(cache, tierSheet, lastColumn, timer);
    if (headers) {
      phaseStartedAt = typeof startAnalyzerPhase_ === "function"
        ? startAnalyzerPhase_(timer)
        : Date.now();
      try { cacheLevelHeaderColumnIndex_(cache, tierSheet, lastColumn); } catch (error) {}
      if (typeof endAnalyzerPhase_ === "function") {
        endAnalyzerPhase_(timer, "headerColumnIndexWrite", phaseStartedAt);
      }
      levelHeaderLoadSource_ = "document-cache";
      return headers;
    }
  }

  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  const headers = getLevelHeaders_(tierSheet, null, timer, lastColumn);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerSheetDiscovery", phaseStartedAt);
  }
  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  cacheLevelHeaders_(tierSheet, lastColumn, headers);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "headerCacheWrite", phaseStartedAt);
  }
  levelHeaderLoadSource_ = "sheet";
  return headers;
}

function invalidateLevelHeaderCache_(tierSheet) {
  if (!tierSheet) return;
  const cache = getLevelHeaderCache_();
  if (!cache) return;
  try {
    const indexKey = levelHeaderColumnIndexKey_(tierSheet);
    const indexedLastColumn = parseCachedLevelHeaderColumn_(cache.get(indexKey));
    if (indexedLastColumn) {
      cache.remove(levelHeaderCacheKey_(tierSheet, indexedLastColumn));
    } else if (typeof tierSheet.getLastColumn === "function") {
      cache.remove(levelHeaderCacheKey_(tierSheet, tierSheet.getLastColumn()));
    }
    cache.remove(indexKey);
  } catch (error) {}
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

function clearLoadedOpinionData_(sh) {
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const height = Math.max(0, lastRow - DATA_START_ROW + 1);
  if (height > 0) {
    const range = sh.getRange(DATA_START_ROW, 1, height, 3);
    if (typeof range.clear === "function") {
      range.clear();
    } else {
      range.clearContent().clearFormat().clearNote();
    }
    range.setFontFamily("Mukta").setFontSize(10);
  }
}

// clear A:C and output area
function clearAnalysisArea_(tool, timer) {
  const sh = tool || SpreadsheetApp.getActive().getSheetByName(ANALYSIS_SHEET_NAME);
  if (!sh) return;
  let phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  clearLoadedOpinionData_(sh);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "loadedOpinionsClear", phaseStartedAt);
  }
  phaseStartedAt = typeof startAnalyzerPhase_ === "function"
    ? startAnalyzerPhase_(timer)
    : Date.now();
  clearAnalysisOutput_(sh, true);
  if (typeof endAnalyzerPhase_ === "function") {
    endAnalyzerPhase_(timer, "analysisOutputClear", phaseStartedAt);
  }
}
