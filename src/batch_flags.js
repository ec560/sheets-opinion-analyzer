const TIER_FLAG_STYLES = {
  movement: {
    priority: 1,
    background: "#b3261e",
    text: "#ffffff",
    fontWeight: "bold",
    styleDifference: false
  },
  book: {
    priority: 3,
    background: "#f9ab00",
    text: "#3c2400",
    fontWeight: "bold",
    styleDifference: true
  },
  experienced_book: {
    priority: 6,
    background: "#d9f7fa",
    text: "#0b5963",
    fontWeight: "normal",
    styleDifference: false
  },
  experienced_book_extreme: {
    priority: 4,
    background: "#00ffff",
    text: "#00363a",
    fontWeight: "bold",
    styleDifference: false
  },
  lock: {
    priority: 2,
    background: "#d2e3fc",
    text: "#174ea6",
    fontWeight: "bold",
    styleDifference: false
  },
  needs_more: {
    priority: 5,
    background: "#fef7e0",
    text: "#7a4f01",
    fontWeight: "bold",
    styleDifference: true
  },
  low_unsettled: {
    priority: 7,
    background: "#fce8e6",
    text: "#b3261e",
    fontWeight: "bold",
    styleDifference: true
  },
  low_solid: {
    priority: 8,
    background: "#f1f3f4",
    text: "#5f6368",
    fontWeight: "normal",
    styleDifference: true
  },
  pending_placement: {
    priority: 1,
    background: "#800b3b",
    text: "#fff7fb",
    fontWeight: "bold",
    styleDifference: false
  },
  pending_natural: {
    priority: 2,
    background: "#956308",
    text: "#fff9ec",
    fontWeight: "bold",
    styleDifference: true
  },
  pending_no_opinions: {
    priority: 3,
    background: "#e2e8f0",
    text: "#46536b",
    fontWeight: "bold",
    styleDifference: false
  },
  pending_high_count: {
    priority: 4,
    background: "#dceafb",
    text: "#153d70",
    fontWeight: "bold",
    styleDifference: false
  },
  pending_split: {
    priority: 5,
    background: "#e4d0f8",
    text: "#4c1d78",
    fontWeight: "bold",
    styleDifference: true
  },
  pending_strong_split: {
    priority: 6,
    background: "#eee2fa",
    text: "#512578",
    fontWeight: "bold",
    styleDifference: true
  },
  pending_low_count: {
    priority: 8,
    background: "#e2e8f0",
    text: "#46536b",
    fontWeight: "bold",
    styleDifference: false
  },
  pending_low_reliability: {
    priority: 7,
    background: "#e2e8f0",
    text: "#46536b",
    fontWeight: "bold",
    styleDifference: false
  },
  pending_unflagged: {
    priority: 99,
    background: "#f1f3f7",
    text: "#586579",
    fontWeight: "normal",
    styleDifference: false
  },
  neutral: {
    priority: 99,
    background: "#f5f5f5",
    text: "#202124",
    fontWeight: "bold",
    styleDifference: false
  }
};

function tierFlagStyle_(styleKey) {
  return TIER_FLAG_STYLES[styleKey] || TIER_FLAG_STYLES.neutral;
}

function normalizePlayerSheetName_(value) {
  return String(value ?? "").trim().toLowerCase();
}

function experiencedPlayerRosterError_(message) {
  const error = new Error("Experienced player roster: " + message);
  error.name = "ExperiencedPlayerRosterError";
  return error;
}

function loadExperiencedPlayerRosters_(configuration) {
  const config = configuration || readExperiencedPlayerConfiguration_();
  const spreadsheetId = String(config && config.spreadsheetId || "").trim();
  if (!spreadsheetId) return null;

  let rosterSpreadsheet;
  try {
    rosterSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    return null;
  }
  if (!rosterSpreadsheet) return null;

  const rosterTabName = String(config.rosterTabName || "").trim() ||
    EXPERIENCED_PLAYER_DEFAULT_ROSTER_TAB;
  let rosterSheet;
  try {
    rosterSheet = rosterSpreadsheet.getSheetByName(rosterTabName);
  } catch (error) {
    return null;
  }
  if (!rosterSheet) return null;

  let classicHeader;
  let platformerHeader;
  let classicValues;
  let platformerValues;
  try {
    classicHeader = rosterSheet.getRange("A1").getDisplayValue();
    platformerHeader = rosterSheet.getRange("B1").getDisplayValue();
    const lastRow = rosterSheet.getLastRow();
    classicValues = lastRow > 1
      ? rosterSheet.getRange(2, 1, lastRow - 1, 1).getValues()
      : [];
    platformerValues = lastRow > 1
      ? rosterSheet.getRange(2, 2, lastRow - 1, 1).getValues()
      : [];
  } catch (error) {
    return null;
  }

  if (String(classicHeader || "").trim() !== EXPERIENCED_PLAYER_CLASSIC_HEADER) {
    throw experiencedPlayerRosterError_(
      "cell A1 on \"" + rosterTabName + "\" must contain \"" +
        EXPERIENCED_PLAYER_CLASSIC_HEADER + "\"."
    );
  }

  const normalizedPlatformerHeader = String(platformerHeader || "").trim();
  const hasPlatformerNames = platformerValues.some(row => normalizePlayerSheetName_(row && row[0]));
  if (normalizedPlatformerHeader && normalizedPlatformerHeader !== EXPERIENCED_PLAYER_PLATFORMER_HEADER) {
    throw experiencedPlayerRosterError_(
      "cell B1 on \"" + rosterTabName + "\" must contain \"" +
        EXPERIENCED_PLAYER_PLATFORMER_HEADER + "\"."
    );
  }
  if (!normalizedPlatformerHeader && hasPlatformerNames) {
    throw experiencedPlayerRosterError_(
      "cell B1 on \"" + rosterTabName + "\" must contain \"" +
        EXPERIENCED_PLAYER_PLATFORMER_HEADER + "\" when column B contains names."
    );
  }

  return {
    classic: buildExperiencedPlayerNameSet_(classicValues, "classic"),
    platformer: normalizedPlatformerHeader
      ? buildExperiencedPlayerNameSet_(platformerValues, "platformer")
      : null
  };
}

function buildExperiencedPlayerNameSet_(values, rosterLabel) {
  const experiencedPlayers = new Set();
  values.forEach(row => {
    const player = normalizePlayerSheetName_(row && row[0]);
    if (!player) return;
    if (experiencedPlayers.has(player)) {
      throw experiencedPlayerRosterError_(
        "duplicate " + rosterLabel + " sheet name \"" + player + "\"."
      );
    }
    experiencedPlayers.add(player);
  });
  return experiencedPlayers;
}

// Retained for callers that only need the original column-A roster.
function loadExperiencedPlayerNames_(configuration) {
  const rosters = loadExperiencedPlayerRosters_(configuration);
  return rosters ? rosters.classic : null;
}

function countExperiencedPlayersForLevel_(values, experiencedPlayers) {
  const matched = new Set();
  (values || []).forEach(row => {
    const player = normalizePlayerSheetName_(row && row[0]);
    if (player && experiencedPlayers.has(player)) matched.add(player);
  });
  return matched.size;
}

function countUniquePlayersForLevel_(values) {
  const players = new Set();
  (values || []).forEach(row => {
    const player = normalizePlayerSheetName_(row && row[0]);
    if (player) players.add(player);
  });
  return players.size;
}

function buildExperiencedPlayerBookshelfFlag_(experiencedCount, sampleSize, totalOpinions) {
  const total = Number(sampleSize);
  const experienced = Number(experiencedCount);
  const opinions = Number(totalOpinions);
  if (!Number.isFinite(total) || total <= 0 || !Number.isFinite(experienced) || experienced < 0) {
    return null;
  }
  if (!Number.isFinite(opinions) || opinions < 0 ||
    opinions >= EXPERIENCED_PLAYER_OPINION_EXEMPT_THRESHOLD) return null;

  const share = experienced / total;
  const percentage = (share * 100).toFixed(1).replace(/\.0$/, "");
  const detail = experienced + "/" + total + ", " + percentage + "%";
  if (share > EXPERIENCED_PLAYER_HIGH_SHARE_THRESHOLD) {
    return {
      text: "many experienced (" + detail + ")",
      emphasized: share > EXPERIENCED_PLAYER_EMPHASIS_HIGH_SHARE_THRESHOLD
    };
  }
  if (share < EXPERIENCED_PLAYER_LOW_SHARE_THRESHOLD &&
    experienced < EXPERIENCED_PLAYER_LOW_COUNT_THRESHOLD) {
    return {
      text: "few experienced (" + detail + ")",
      emphasized: share < EXPERIENCED_PLAYER_EMPHASIS_LOW_SHARE_THRESHOLD
    };
  }
  return null;
}

function buildExperiencedPlayerBookshelfTag_(experiencedCount, sampleSize, totalOpinions) {
  const flag = buildExperiencedPlayerBookshelfFlag_(experiencedCount, sampleSize, totalOpinions);
  return flag ? flag.text : "";
}

function scanSelectedTierFlags() {
  return withAnalyzerDocumentLock_(scanSelectedTierFlagsUnlocked_);
}

function scanSelectedTierFlagsUnlocked_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (typeof loadTierConfiguration_ === "function") {
    const configResult = loadTierConfiguration_();
    const configError = tierConfigurationErrorMessage_(configResult);
    if (configError) {
      showAnalyzerAlert_(ui, "Error with Tier Configuration", configError, ui.ButtonSet.OK);
      return;
    }
  }
  const tierName = getFlagScanTierName_(ss, tool);

  if (!tierName) {
    showAnalyzerAlert_(ui, "Tier Flag Scan", "Open a tier sheet or pick a tier in " + TIER_CELL + " first.", ui.ButtonSet.OK);
    return;
  }

  const tierSheet = ss.getSheetByName(tierName);
  if (!tierSheet || isAnalyzerUtilitySheetName_(tierName)) {
    showAnalyzerAlert_(ui, "Tier Flag Scan", "Could not find a tier sheet named \"" + tierName + "\".", ui.ButtonSet.OK);
    return;
  }

  let result;
  try {
    result = buildTierFlagScan_(tierName, tierSheet);
  } catch (error) {
    if (error && error.name === "ExperiencedPlayerRosterError") {
      showAnalyzerAlert_(ui, "Tier Flag Scan", error.message, ui.ButtonSet.OK);
      return;
    }
    throw error;
  }
  if (result.scanned === 0) {
    const headerType = tierName.toLowerCase() === "pending" ? "level" : "unlocked level";
    showAnalyzerAlert_(ui, "Tier Flag Scan", "No " + headerType + " headers found on \"" + tierName + "\".", ui.ButtonSet.OK);
    return;
  }

  const outputSheet = renderTierFlagScan_(ss, tierName, result.rows);
  ss.setActiveSheet(outputSheet);

  showAnalyzerAlert_(ui,
    "Tier Flag Scan",
    result.rows.length + (tierName.toLowerCase() === "pending" ? " pending level" : " flagged level") +
      (result.rows.length === 1 ? "" : "s") +
      " found across " + result.scanned + " level" + (result.scanned === 1 ? "" : "s") + ".",
    ui.ButtonSet.OK
  );
}

function getFlagScanTierName_(ss, tool) {
  const active = ss.getActiveSheet();
  if (active) {
    const activeName = active.getName();
    if (!isAnalyzerUtilitySheetName_(activeName)) return activeName;
  }

  if (!tool) return "";
  return String(tool.getRange(TIER_CELL).getDisplayValue() || "").trim();
}

function buildTierFlagScan_(tierName, tierSheet) {
  const isPendingTier = tierName.toLowerCase() === "pending";
  const lastRow = tierSheet.getLastRow();
  const lastCol = tierSheet.getLastColumn();
  const numRows = Math.max(0, lastRow - 1);

  let vals = [];
  let bgs = [];
  let fcs = [];
  if (numRows > 0 && lastCol > 0) {
    const sourceRange = tierSheet.getRange(2, 1, numRows, lastCol);
    vals = sourceRange.getValues();
    bgs = sourceRange.getBackgrounds();
    fcs = sourceRange.getFontColors();
  }
  const headers = getLevelHeaders_(tierSheet, vals);
  const experiencedPlayerConfig = readExperiencedPlayerConfiguration_();
  const experiencedPlayerRosters = loadExperiencedPlayerRosters_(experiencedPlayerConfig);

  const rows = [];
  let scanned = 0;
  const platformerStartIndex = findPlatformerSectionStartIndex_(tierSheet, headers, vals);
  const lockedLevelStates = typeof getLevelLockStates_ === "function"
    ? getLevelLockStates_(tierSheet, headers.map(header => header.col))
    : null;
  headers.forEach((header, headerIndex) => {
    const sectionIndex = platformerStartIndex >= 0 && headerIndex >= platformerStartIndex ? 1 : 0;

    const isLocked = lockedLevelStates
      ? lockedLevelStates[headerIndex]
      : (typeof isLevelLocked_ === "function" && isLevelLocked_(tierSheet, header.col));
    if (isLocked && !isPendingTier) return;
    scanned++;
    const levelData = extractLevelFlagData_(header, vals, bgs, fcs, lastCol);
    let bookshelfFlag = null;
    const experiencedPlayers = experiencedPlayerRosters === null
      ? null
      : (sectionIndex === 1
        ? experiencedPlayerRosters.platformer
        : experiencedPlayerRosters.classic);
    if (experiencedPlayers !== null) {
      const experiencedCount = countExperiencedPlayersForLevel_(levelData.vals, experiencedPlayers);
      const sampleSize = countUniquePlayersForLevel_(levelData.vals);
      bookshelfFlag = buildExperiencedPlayerBookshelfFlag_(
        experiencedCount,
        sampleSize,
        levelData.vals.length
      );
    }
    const analysis = calculateLevelAnalysis_(
      tierName,
      header.name,
      levelData.vals,
      levelData.bgs,
      levelData.fcs
    );
    if (!analysis.isPending && analysis.rawCount === 0 && analysis.allAndFuck === 0) return;

    const flagContext = analysis.isPending
      ? buildPendingFlagContext_(levelData, analysis.countedRowFlags)
      : {};
    flagContext.experiencedBookshelfTag = bookshelfFlag ? bookshelfFlag.text : "";
    flagContext.experiencedBookshelfEmphasized = !!(bookshelfFlag && bookshelfFlag.emphasized);
    const flagSummary = buildLevelFlagSummary_(analysis, flagContext);
    if (flagSummary.flags.length > 0 || analysis.isPending) {
      const flagRow = buildTierFlagRow_(analysis, flagSummary);
      flagRow.sectionIndex = sectionIndex;
      rows.push(flagRow);
    }
  });

  if (isPendingTier) {
    rows.sort((a, b) => a.sectionIndex - b.sectionIndex ||
      b.splitMargin - a.splitMargin);
  }

  return {
    scanned,
    rows
  };
}

function findPlatformerSectionStartIndex_(tierSheet, headers, bodyValues) {
  if (!tierSheet || !headers || headers.length === 0) return -1;

  try {
    const lastColumn = tierSheet.getLastColumn();
    const headerRange = tierSheet.getRange(1, 1, 1, lastColumn);
    const headerValues = headerRange.getDisplayValues()[0];
    const headerMerges = headerRange.getMergedRanges();
    const bodyRowCount = Math.max(0, tierSheet.getLastRow() - 1);

    for (const headerMerge of headerMerges) {
      if (headerMerge.getRow() !== 1 || headerMerge.getNumRows() !== 1) continue;
      const platformerColumn = headerMerge.getColumn();
      const label = String(headerValues[platformerColumn - 1] || "").trim().toLocaleLowerCase();
      if (label !== "platformer") continue;

      const markerWidth = headerMerge.getNumColumns();
      const markerEndColumn = platformerColumn + markerWidth - 1;
      let markerLabelCount = 0;
      const markerBodyHasNoOpinions = (bodyValues || []).every(row => {
        for (let offset = 0; offset < markerWidth; offset++) {
          const value = String(row[platformerColumn - 1 + offset] || "").trim();
          if (!value) continue;
          if (offset !== 0) return false;
          markerLabelCount++;
          if (markerLabelCount > 1) return false;
        }
        return true;
      });
      if (!markerBodyHasNoOpinions || bodyRowCount === 0) continue;

      const bodyRange = tierSheet.getRange(2, platformerColumn, bodyRowCount, markerWidth);
      const hasMergedBody = bodyRange.getMergedRanges().some(bodyMerge => {
        return bodyMerge.getRow() === 2 &&
          bodyMerge.getColumn() === platformerColumn &&
          bodyMerge.getNumColumns() === markerWidth;
      });
      if (!hasMergedBody) continue;

      return headers.findIndex(header => header.col > markerEndColumn);
    }
  } catch (error) {
    // Lightweight test doubles may not expose merged-range metadata.
  }

  return -1;
}

function extractLevelFlagData_(header, vals, bgs, fcs, lastCol) {
  const outVals = [];
  const outBgs = [];
  const outFcs = [];
  const sourceRows = [];
  const startIdx = header.col - 1;

  if (startIdx + 2 >= lastCol) {
    return { vals: outVals, bgs: outBgs, fcs: outFcs, sourceRows };
  }

  for (let r = 0; r < vals.length; r++) {
    const row = [
      safeCellValue_(vals[r][startIdx]),
      safeCellValue_(vals[r][startIdx + 1]),
      safeCellValue_(vals[r][startIdx + 2])
    ];

    if (isBlankFlagOpinionRow_(row)) continue;

    // The source matrices start at sheet row 2; retain gaps for source highlighting.
    sourceRows.push(r + 2);
    outVals.push(row);
    outBgs.push([
      bgs[r][startIdx],
      bgs[r][startIdx + 1],
      bgs[r][startIdx + 2]
    ]);
    outFcs.push([
      fcs[r][startIdx],
      fcs[r][startIdx + 1],
      fcs[r][startIdx + 2]
    ]);
  }

  return { vals: outVals, bgs: outBgs, fcs: outFcs, sourceRows };
}

function isBlankFlagOpinionRow_(row) {
  return String(row[0]).trim() === "" &&
    String(row[1]).trim() === "" &&
    String(row[2]).trim() === "";
}

function buildLevelFlagSummary_(analysis, flagContext) {
  if (analysis && analysis.isPending) {
    return buildPendingLevelFlagSummary_(analysis, flagContext || {});
  }

  let flag = "";
  let styleKey = "";
  let differenceAlert = analysis.isLockWorthy ? null : analysis.difference;
  const needsMoreOpinionsAlert = buildNeedsMoreOpinionsAlert_(analysis);
  const lowOpinionAlert = buildLowOpinionAlert_(analysis);

  if (analysis.canMove) {
    styleKey = "movement";
    if (analysis.isPending) {
      flag = "placement alert (" + analysis.verdictTier + ")";
    } else {
      const direction = analysis.moveDirection ? ", " + analysis.moveDirection : "";
      flag = "move alert (" + analysis.verdictTier + direction + ")";
    }
  } else if (analysis.isLockWorthy) {
    styleKey = "lock";
    flag = "lock alert (" + (analysis.lockSharePct * 100).toFixed(0) + "% " + analysis.lockSideLabel + ")";
  } else if (analysis.bookAlert) {
    styleKey = "book";
    flag = formatBookAlert_(analysis.bookAlert);
    differenceAlert = analysis.bookAlert;
  } else if (flagContext && flagContext.experiencedBookshelfTag &&
    flagContext.experiencedBookshelfEmphasized) {
    styleKey = "experienced_book_extreme";
    flag = flagContext.experiencedBookshelfTag;
    differenceAlert = null;
  } else if (needsMoreOpinionsAlert) {
    styleKey = "needs_more";
    flag = "needs more opinions";
    differenceAlert = needsMoreOpinionsAlert;
  } else if (flagContext && flagContext.experiencedBookshelfTag) {
    styleKey = "experienced_book";
    flag = flagContext.experiencedBookshelfTag;
    differenceAlert = null;
  } else if (lowOpinionAlert) {
    styleKey = lowOpinionAlert.subdued ? "low_solid" : "low_unsettled";
    flag = lowOpinionAlert.text;
  }

  return {
    flags: flag ? [flag] : [],
    differenceAlert,
    styleKey,
    priority: tierFlagStyle_(styleKey).priority,
    subduedLowOpinion: !!(lowOpinionAlert && lowOpinionAlert.subdued && flag === lowOpinionAlert.text)
  };
}

function buildPendingLevelFlagSummary_(analysis, flagContext) {
  let flag = "";
  let styleKey = "";
  let differenceAlert = null;
  const recentExternalSource = flagContext.mostRecentExternalSource || "";
  const splitMargin = analysis.decisionTierSplit
    ? Number(analysis.decisionTierSplit.marginWeight) || 0
    : 0;

  if (analysis.canMove && !recentExternalSource) {
    styleKey = "movement";
    flag = "placement alert (" + analysis.verdictTier + ")";
  } else if (analysis.canMove && recentExternalSource) {
    styleKey = "pending_natural";
    flag = "requires natural opinion to place";
    differenceAlert = analysis.difference;
  } else if (analysis.rawCount === 0 && analysis.allAndFuck === 0) {
    styleKey = "pending_no_opinions";
    flag = "no opinions";
  } else if (analysis.rawCount >= PENDING_SCAN_HIGH_OPINION_MIN_RAW_COUNT) {
    styleKey = "pending_high_count";
    flag = "high opinion count (" + analysis.rawCount + ")";
  } else if (splitMargin === PENDING_SCAN_NEAR_PLACEMENT_MARGIN) {
    styleKey = "pending_split";
    flag = "+2.75 split (+" + formatFlagNumber_(splitMargin) + " " + analysis.verdictBaseName + ")";
    differenceAlert = analysis.difference;
  } else if (splitMargin >= PENDING_SCAN_STRONG_SPLIT_MARGIN) {
    styleKey = "pending_strong_split";
    flag = "close to placement";
    differenceAlert = analysis.difference;
  } else if (!flagContext.hasGreenOrBlueReliability) {
    styleKey = "pending_low_reliability";
    flag = "low reliabilities (no green/blue)";
  } else if (analysis.rawCount <= PENDING_SCAN_LOW_OPINION_MAX_RAW_COUNT) {
    styleKey = "pending_low_count";
    flag = "low opinion count (" + analysis.rawCount + "/3)";
  } else {
    styleKey = "pending_unflagged";
  }

  return {
    flags: flag ? [flag] : [],
    differenceAlert,
    styleKey,
    priority: tierFlagStyle_(styleKey).priority,
    subduedLowOpinion: false
  };
}

function buildPendingFlagContext_(levelData, countedRowFlags) {
  const vals = levelData && levelData.vals ? levelData.vals : [];
  const bgs = levelData && levelData.bgs ? levelData.bgs : [];
  const counted = countedRowFlags || [];
  let mostRecentExternalSource = "";
  let hasGreenOrBlueReliability = false;

  for (let row = 0; row < counted.length; row++) {
    if (!counted[row]) continue;
    const reliabilityColor = hex_(bgs[row] && bgs[row][2]);
    if (reliabilityColor === "#00ff00" || reliabilityColor === "#00ffff") {
      hasGreenOrBlueReliability = true;
    }
  }

  for (let row = counted.length - 1; row >= 0; row--) {
    if (!counted[row]) continue;
    mostRecentExternalSource = externalOpinionSource_(vals[row] && vals[row][2]);
    break;
  }

  return {
    mostRecentExternalSource,
    hasGreenOrBlueReliability
  };
}

function externalOpinionSource_(reliabilityText) {
  const value = String(reliabilityText || "").trim();
  const sources = [
    { pattern: /\baredl\b/i, label: "AREDL opinions" },
    { pattern: /\budl\b/i, label: "UDL opinions" },
    { pattern: /\bexternal\b/i, label: "External opinions" },
    { pattern: /\bextrapolated\b/i, label: "Extrapolated opinions" },
    { pattern: /\bgddl\b/i, label: "GDDL opinions" }
  ];

  for (const source of sources) {
    if (source.pattern.test(value)) return source.label;
  }
  return "";
}

function buildLowOpinionAlert_(analysis) {
  if (!analysis) return null;

  const rawCountIsLow = analysis.rawCount <= FLAG_SCAN_LOW_OPINION_MAX_RAW_COUNT;
  const weightIsLow = analysis.allWeight < FLAG_SCAN_LOW_OPINION_WEIGHT;
  const strongCurrentTierLean = hasStrongCurrentTierLean_(analysis);

  if (!rawCountIsLow && (!weightIsLow || strongCurrentTierLean)) return null;

  return {
    text: formatLowOpinionFlag_(analysis),
    subdued: rawCountIsLow && strongCurrentTierLean
  };
}

function hasStrongCurrentTierLean_(analysis) {
  if (!analysis || analysis.decisionCurrentIdx < 0 || !analysis.difference) return false;

  const currentSide = analysis.decisionCurrentIdx <= analysis.decisionTierSplit.lowIndex ? "left" : "right";
  return analysis.difference.winningSide === currentSide &&
    analysis.difference.lean >= FLAG_SCAN_LOW_OPINION_CURRENT_LEAN;
}

function formatLowOpinionFlag_(analysis) {
  return "low opinion count (" +
    formatFlagNumber_(analysis.allWeight) +
    "/" +
    FLAG_SCAN_LOW_OPINION_WEIGHT +
    " weight)";
}

function buildNeedsMoreOpinionsAlert_(analysis) {
  if (!analysis || analysis.decisionCurrentIdx < 0 || !analysis.difference) return null;
  if (analysis.allWeight > FLAG_SCAN_NEEDS_MORE_OPINIONS_MAX_WEIGHT) return null;
  if (analysis.bookAlert) return null;
  if (analysis.difference.lean > FLAG_BOOK_LEAN_WEIGHT) return null;

  const currentSide = analysis.decisionCurrentIdx <= analysis.decisionTierSplit.lowIndex ? "left" : "right";
  const favorsCurrentTier = analysis.difference.lean === 0 || analysis.difference.winningSide === currentSide;

  return {
    lean: analysis.difference.lean,
    tier: favorsCurrentTier ? analysis.currentTier : analysis.difference.tier
  };
}

function buildTierFlagRow_(analysis, flagSummary) {
  const comparison = analysis.placementComparison || analysis.decisionTierSplit || analysis.tierSplit;
  return {
    values: [
      analysis.levelName,
      formatFlagNumber_(analysis.allWeight),
      analysis.rawCount,
      flagSummary.styleKey === "pending_unflagged" ? "no flag" : flagSummary.flags.join("; "),
      comparison.left.label + " / " + comparison.right.label,
      formatFlagNumber_(comparison.left.weight) + " | " + formatFlagNumber_(comparison.right.weight),
      analysis.isPending || flagSummary.differenceAlert
        ? formatComparisonDifference_(comparison) : ""
    ],
    splitMargin: Math.abs(Number(comparison.left.weight) - Number(comparison.right.weight)),
    styleKey: flagSummary.styleKey || "",
    priority: flagSummary.priority || 99,
    subduedLowOpinion: !!flagSummary.subduedLowOpinion
  };
}

function renderTierFlagScan_(ss, tierName, flagRows) {
  let sh = ss.getSheetByName(FLAG_SCAN_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(FLAG_SCAN_SHEET_NAME);

  const headers = [
    "Level",
    "Weight",
    "Raw",
    "Flags",
    "Split",
    "Split totals",
    "Difference"
  ];
  setManagedSheetColumnCount_(sh, headers.length);
  sh.clear();

  const titleRow = new Array(headers.length).fill("");
  titleRow[0] = "Tier flag scan";
  titleRow[1] = tierName;
  titleRow[2] = VERSION;
  titleRow[3] = new Date();

  const rows = [titleRow, headers];
  if (flagRows.length > 0) {
    flagRows.forEach(row => rows.push(row.values));
  } else {
    const emptyRow = new Array(headers.length).fill("");
    emptyRow[0] = "No flagged levels";
    rows.push(emptyRow);
  }

  sh.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatTierFlagScan_(sh, rows, headers.length, flagRows);
  return sh;
}

function formatTierFlagScan_(sh, rows, colCount, flagRows) {
  const rowCount = rows.length;
  if (rowCount === 0 || colCount === 0) return;
  const isPendingScan = String(rows[0][1] || "").toLowerCase() === "pending";

  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, rowCount, colCount)
    .setFontFamily("Mukta")
    .setFontSize(10)
    .setFontColor("#202124")
    .setBackground("#ffffff")
    .setVerticalAlignment("middle");

  sh.getRange(1, 1, 1, colCount)
    .setBackground(isPendingScan ? "#edf0f4" : "#e8f0fe")
    .setFontWeight("bold");

  sh.getRange(1, 3)
    .setFontColor("#5f6368")
    .setFontWeight("normal")
    .setHorizontalAlignment("right");

  sh.getRange(1, 4)
    .setNumberFormat("m/d/yyyy h:mm")
    .setFontColor("#5f6368")
    .setFontWeight("normal")
    .setHorizontalAlignment("left");
  styleTierFlagCell_(sh, 1, 2, rows[0][1]);

  sh.getRange(2, 1, 1, colCount)
    .setBackground(isPendingScan ? "#e2e8f0" : "#eeeeee")
    .setFontColor(isPendingScan ? "#46536b" : "#202124")
    .setFontWeight("bold");
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(4, 225);
  sh.setColumnWidth(5, 160);
  sh.setColumnWidth(7, 125);

  if (rowCount <= 2) return;

  sh.getRange(3, 2, rowCount - 2, 1).setNumberFormat("0.###");
  sh.getRange(3, 3, rowCount - 2, 1).setNumberFormat("0");
  sh.getRange(3, 2, rowCount - 2, 2).setHorizontalAlignment("right");

  const backgrounds = [];
  const fontColors = [];
  const fontWeights = [];
  const fontStyles = [];

  for (let i = 2; i < rowCount; i++) {
    const row = rows[i];
    const flagRow = flagRows && flagRows.length > (i - 2) ? flagRows[i - 2] : null;
    const styleKey = flagRow && flagRow.styleKey;
    const flagStyle = tierFlagStyle_(isPendingScan && styleKey === "movement" ? "pending_placement" : styleKey);
    const rowBackgrounds = new Array(colCount).fill(isPendingScan ? "#fbfcfe" : "#ffffff");
    const rowFontColors = new Array(colCount).fill(isPendingScan ? "#20242e" : "#202124");
    const rowFontWeights = new Array(colCount).fill("normal");
    const rowFontStyles = new Array(colCount).fill("normal");
    if (isPendingScan) rowFontColors[6] = "#3c4043";

    if (String(row[0] || "") === "No flagged levels") {
      rowBackgrounds.fill("#f5f5f5");
      rowFontColors.fill("#5f6368");
      rowFontStyles.fill("italic");
    } else if (flagRow && flagRow.styleKey === "pending_unflagged") {
      if (!isPendingScan) rowBackgrounds.fill(flagStyle.background);
      rowBackgrounds[3] = flagStyle.background;
      // De-emphasize status without making the split evidence look disabled.
      rowFontColors.fill("#3c4043");
      rowFontColors[0] = flagStyle.text;
      rowFontColors[3] = flagStyle.text;
    } else {
      rowFontWeights[0] = flagStyle.fontWeight;
      rowBackgrounds[3] = flagStyle.background;
      rowFontColors[3] = flagStyle.text;
      rowFontWeights[3] = flagStyle.fontWeight;

      if (isPendingScan) {
        rowFontWeights[0] = "bold";
        rowFontWeights[3] = "bold";
      } else if (flagStyle.styleDifference) {
        rowBackgrounds[6] = flagStyle.background;
        rowFontColors[6] = flagStyle.text;
        rowFontWeights[6] = flagStyle.fontWeight;
      }
    }

    backgrounds.push(rowBackgrounds);
    fontColors.push(rowFontColors);
    fontWeights.push(rowFontWeights);
    fontStyles.push(rowFontStyles);
  }

  const resultRange = sh.getRange(3, 1, rowCount - 2, colCount);
  resultRange
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors)
    .setFontWeights(fontWeights)
    .setFontStyles(fontStyles);

  if (isPendingScan) {
    sh.getRange(3, 4, rowCount - 2, 1).setWrap(true);
    sh.getRange(3, 7, rowCount - 2, 1).setHorizontalAlignment("right");
    sh.autoResizeRows(3, rowCount - 2);
  }

  applyTierFlagSectionBorders_(sh, flagRows, colCount);
}

function applyTierFlagSectionBorders_(sh, flagRows, colCount) {
  if (!flagRows || flagRows.length === 0) return;

  if ((Number(flagRows[0].sectionIndex) || 0) === 1) {
    sh.getRange(3, 1, 1, colCount).setBorder(
      true,
      null,
      null,
      null,
      null,
      null,
      "#9aa0a6",
      SpreadsheetApp.BorderStyle.DOTTED
    );
    return;
  }

  for (let index = 1; index < flagRows.length; index++) {
    const previousSection = Number(flagRows[index - 1].sectionIndex) || 0;
    const currentSection = Number(flagRows[index].sectionIndex) || 0;
    if (currentSection === previousSection) continue;

    // Output begins on row 3; index points at the first row in the new section,
    // so index + 2 is the preceding section's final displayed row.
    sh.getRange(index + 2, 1, 1, colCount).setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "#9aa0a6",
      SpreadsheetApp.BorderStyle.DOTTED
    );
    return;
  }

}

function styleTierFlagCell_(sh, row, col, tierName) {
  const name = String(tierName || "").trim();
  if (!name) return;

  const nameToColor = buildTierNameToColor_();
  const hex = nameToColor[name];
  if (!hex) return;

  sh.getRange(row, col)
    .setBackground(hex)
    .setFontColor(tierTextColor_(name))
    .setFontWeight("bold");
}

function formatFlagNumber_(value) {
  const n = Number(value);
  if (!isFinite(n)) return "";
  return n.toFixed(3).replace(/\.?0+$/, "");
}

function formatBookAlert_(alert) {
  return alert.direction + " book alert (" + formatDifference_(alert) + ")";
}

function formatDifference_(alert) {
  if (!alert) return "";

  return "+" +
    formatFlagNumber_(alert.lean) +
    " " +
    alert.tier;
}

function formatComparisonDifference_(comparison) {
  if (!comparison || !comparison.left || !comparison.right) return "";

  const leftWeight = Number(comparison.left.weight) || 0;
  const rightWeight = Number(comparison.right.weight) || 0;
  const winner = leftWeight > rightWeight ? comparison.left : comparison.right;

  return "+" +
    formatFlagNumber_(Math.abs(leftWeight - rightWeight)) +
    " " +
    winner.label;
}
