// Main analysis function that reads opinions, applies weights, calculates stats, and outputs results
function analyzeSelectedLevel() {
  const ss = SpreadsheetApp.getActive();
  const tool = ss.getSheetByName(ANALYSIS_SHEET_NAME);

  const tierName = tool.getRange(TIER_CELL).getDisplayValue().trim();
  const levelName = tool.getRange(LEVEL_CELL).getDisplayValue().trim();
  if (!tierName || !levelName) {
    SpreadsheetApp.getUi().alert("Pick a tier and a level first.");
    return;
  }

  // Read populated A:C
  const lastRow = tool.getLastRow();
  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No opinions pasted yet. Pick the level dropdown again.");
    return;
  }

  const numRows = lastRow - (DATA_START_ROW - 1);
  const rng = tool.getRange(DATA_START_ROW, 1, numRows, 3);
  const vals = rng.getValues();
  const bgs = rng.getBackgrounds();
  const fcs = rng.getFontColors();

  // Accumulate weighted votes
  const weightsByTier = {};
  for (const n of orderedTierNames) weightsByTier[n] = 0;

  let fuckPresent = false;
  let fuckWeight = 0; // track separately so we can calculate later

  let totalWeight = 0;
  let rawCount = 0;
  const rawPoints = []; // numeric tier indices per raw opinion (splits -> midpoint)

  const tierIdxOf_ = (tierName) => orderedTierNames.indexOf(tierName);

  for (let r = 0; r < vals.length; r++) {
    const player = String(vals[r][0] ?? "").trim();
    const opinionText = String(vals[r][1] ?? "").trim();
    const relText = String(vals[r][2] ?? "").trim();

    if (!player && !opinionText && !relText) continue;

    const opinionColor = hex_(bgs[r][1]);      // column B background
    const relColor = hex_(bgs[r][2]);          // column C background
    const w = reliabilityFactors[relColor] ?? 0;

    const opinionFont = hex_(fcs[r][1]);       // column B font color

    if (w <= 0) continue;

    // !!! Fuck tier opinion handling (supports 3-way splits) !!!
    // - fuck opinions count normally toward tier math, but "fuck" itself is excluded from the calculation
    // - Always: fuckPresent=true and fuckWeight += w
    // - Allocate full weight to the tier/tiers implied by the text font colors
    // - If no tier is implied treat as a fuck tier vote (full weight).
    if (opinionColor === "#000000") {
      fuckPresent = true;
      rawCount++;

      // Always count full Fuck weight for black opinions
      fuckWeight += w;

      // Helper: add full-weight tier contributions and ensure totalWeight reflects normal values
      const addTierWeight_ = (tier, wt) => {
        if (!tier || wt <= 0) return;
        weightsByTier[tier] = (weightsByTier[tier] || 0) + wt;
      };

      // 3-way: black background & font is a split-color (ex: Insane/Extreme)
      const fontKey = hex_(opinionFont);
      if (splitPairs[fontKey]) {
        const [lowHex, highHex] = splitPairs[fontKey].map(hex_);
        const lowTier = difficultyColorNames[lowHex];
        const highTier = difficultyColorNames[highHex];

        const i1 = tierIdxOf_(lowTier);
        const i2 = tierIdxOf_(highTier);
        if (i1 >= 0 && i2 >= 0) {
          rawPoints.push((i1 + i2) / 2);
        }

        // Count normally: total tier weight == w (50/50 split)
        if (lowTier && highTier) {
          const halfW = w / 2;
          addTierWeight_(lowTier, halfW);
          addTierWeight_(highTier, halfW);
          totalWeight += w;
        }
        continue;
      }

      // 2-way: black background & font is a real tier color (ex: Very Hard)
      const fontKey2 = hex_(opinionFont);
      const fontTier = difficultyColorNames[fontKey2]; // tier name if font matches a tier color
      // If font indicates a tier, count FULL w for that tier (fuck is just a flag on top)
      if (fontTier && fontTier !== "Insane") {
        const it = tierIdxOf_(fontTier);
        if (it >= 0) rawPoints.push(it);

        addTierWeight_(fontTier, w);
        totalWeight += w;
        continue;
      }

      // 2-way special case: ONLY treat as Insane if the text contains "insane"
      const hasInsane = String(opinionText).toLowerCase().includes("insane");
      if (hasInsane) {
        const it = tierIdxOf_("Insane");
        if (it >= 0) rawPoints.push(it);

        addTierWeight_("Insane", w);
        totalWeight += w;
        continue;
      }

      // Otherwise: do nothing

      totalWeight += w;

      continue;
    }

    // !!! Normal single-tier opinion (using background color) !!!
    if (difficultyColorNames[opinionColor]) {
      const tName = difficultyColorNames[opinionColor];
      const it = tierIdxOf_(tName);
      if (it >= 0) rawPoints.push(it);

      weightsByTier[tName] = (weightsByTier[tName] || 0) + w;
      rawCount++;
      totalWeight += w;
      continue;
    }

    // !! Normal split (non-fuck) by splitPairs !!
    const oc = hex_(opinionColor); // normalize lookup key
    if (splitPairs[oc]) {
      const [c1Raw, c2Raw] = splitPairs[oc];
      const c1 = hex_(c1Raw);
      const c2 = hex_(c2Raw);

      const t1 = difficultyColorNames[c1];
      const t2 = difficultyColorNames[c2];

      const i1 = tierIdxOf_(t1);
      const i2 = tierIdxOf_(t2);
      if (i1 >= 0 && i2 >= 0) rawPoints.push((i1 + i2) / 2);

      // regular split 50/50
      if (t1 && t2) {
        const halfW = w / 2;
        weightsByTier[t1] = (weightsByTier[t1] || 0) + halfW;
        weightsByTier[t2] = (weightsByTier[t2] || 0) + halfW;
        rawCount++;
        totalWeight += w;
        continue;
      }
    }

    // Unknown opinion color -> ignore
  }

  const rawMeanIdx = mean_(rawPoints);
  const rawMedianIdx = median_(rawPoints);

  const rawMeanLabel = tierLabelFromIndex_(orderedTierNames, rawMeanIdx);
  const rawMedianLabel = tierLabelFromIndex_(orderedTierNames, rawMedianIdx);

  // Find majority tier by highest weighted total
  const entries = Object.entries(weightsByTier).sort((a, b) => b[1] - a[1]);
  const top = entries[0] || ["(none)", 0];
  const second = entries[1] || ["(none)", 0];

  const tierWeightSum = orderedTierNames.reduce((s, t) => s + (weightsByTier[t] || 0), 0);
  const allWeight = tierWeightSum;
  const allAndFuck = allWeight + fuckWeight;

  const topTier = top[0];
  const topWeight = top[1];
  const secondTier = second[0];
  const secondWeight = second[1];

  // Majority rule (weighted)
  const passesMajority = topWeight >= 4; // Must have at least 4 opinions (by weight)
  const topVsSecond = (topWeight - secondWeight) >= 3; // Must be leading by at least 3 points

  // Weighted SD over tier indices (orderedTierNames)
  const idx = {}; orderedTierNames.forEach((t, i) => idx[t] = i);
  let sw = 0, mu = 0;
  orderedTierNames.forEach(t => { const wt = weightsByTier[t] || 0; sw += wt; mu += wt * idx[t]; });
  mu = sw ? mu / sw : 0;
  let v = 0;
  orderedTierNames.forEach(t => { const wt = weightsByTier[t] || 0, d = idx[t] - mu; v += wt * d * d; });
  const sd = sw ? Math.sqrt(v / sw) : 0;

  let OUTLIER_K = 1.5;
  if (sd > 1.5) OUTLIER_K = 2.5;
  else if (sd > 1) OUTLIER_K = 2;

  const outlierPctObj = outlierPct_(rawPoints, rawMeanIdx, OUTLIER_K); // >K tiers from mean
  const outlierBounds = outlierCutoffLabels_(orderedTierNames, rawMeanIdx, OUTLIER_K);

  // Split logic
  const idxOf = (name) => orderedTierNames.indexOf(name);

  const sumRange = (startIdx, endIdx) => {
    let s = 0;
    for (let i = startIdx; i <= endIdx; i++) s += (weightsByTier[orderedTierNames[i]] || 0);
    return s;
  };

  const topIdx = idxOf(topTier);
  const runnerIdx = idxOf(secondTier);

  const currentTier = String(tierName || "").trim();
  const currentIdx = orderedTierNames.indexOf(currentTier);
  const currentTierLower = currentTier.toLowerCase();
  const isPending = (currentTierLower === "pending");

  let splitLowIdx = pickSplitLowIdx_(
    orderedTierNames,
    weightsByTier,
    topIdx,
    runnerIdx,
    currentIdx,
    rawMedianIdx,
    isPending
  );

  const splitHighIdx = splitLowIdx + 1;

  const splitLowTier = orderedTierNames[splitLowIdx];
  const splitHighTier = orderedTierNames[splitHighIdx];

  // Side totals: <= low vs >= high
  const splitLeftTotal = sumRange(0, splitLowIdx);
  const splitRightTotal = sumRange(splitHighIdx, orderedTierNames.length - 1);

  const toppct = allWeight > 0 ? topWeight / allWeight : 0;
  const fuckpct = allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0;

  let vtn = "";
  if (fuckPresent) {
    if ((((allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0) >= 0.2 && toppct < 0.35) || fuckpct >= 0.5) || (sd >= 2 && passesMajority)) {
      vtn = "Fuck";
    } else {
      vtn = (splitLeftTotal > splitRightTotal) ? splitLowTier : splitHighTier;
    }
  } else {
    vtn = (splitLeftTotal > splitRightTotal) ? splitLowTier : splitHighTier;
  }

  const verdictTier = String(vtn).trim();
  const verdictDiffersFromCurrent = verdictTier !== "" && verdictTier !== currentTier;

  let splitValue = "";
  if (allWeight >= 50) {
    splitValue = 4;
  } else if (allWeight >= 100) {
    splitValue = 5;
  } else {
    splitValue = 3;
  }

  const splitThreshold = splitValue; // minimum points difference needed to consider split as a majority
  const splitMargin = Math.abs(splitLeftTotal - splitRightTotal);
  const splitMarginPct = totalWeight > 0 ? (splitMargin / totalWeight) : 0;
  const SPLIT_PCT_MIN_PENDING = 0.20;   // 60% side majority for pending
  const SPLIT_PCT_MIN_MOVES   = 0.05;   // at least 50/50 for non-pending levels

  const passesSplitMajority = Math.abs(splitLeftTotal - splitRightTotal) >= splitThreshold;
  const passesSplitPct = isPending ? (splitMarginPct >= SPLIT_PCT_MIN_PENDING): (splitMarginPct >= SPLIT_PCT_MIN_MOVES);

  // Output
  let out = [];
  out.push([`Tier sheet`, tierName, "", ""]);
  out.push([`Level`, levelName, "", ""]);
  out.push([`Total weighted opinions`, allAndFuck, "", ""]);
  out.push([`Most votes (weighted)`, topTier, topWeight, ""]);
  out.push([`Runner-up (weighted)`, secondTier, secondWeight, ""]);
  out.push([`Tier Mean`, rawMeanLabel, rawMeanIdx, ""]);
  out.push([`Tier Median`, rawMedianLabel, rawMedianIdx, ""]);
  out.push([`Outliers`, `${outlierPctObj.count} / ${outlierPctObj.total} (${(outlierPctObj.pct * 100).toFixed(1)}%)`, "", ""]);
  out.push([
    "Outlier range",
    `≤ ${outlierBounds.lowLabel}` + `       ≥ ${outlierBounds.highLabel}`,
    "",
    ""
  ]);
  out.push([`Standard Deviation`, sd, "", ""]);

  if (fuckPresent) {
    out.push([`Fuck weight`, fuckWeight, "", ""]);
    out.push([`Fuck % of all`, (allAndFuck) > 0 ? fuckWeight / (totalWeight || 1) : 0, "", ""]);
  }

  let verdictTierName = verdictTier;
  if (fuckPresent) {
    if (((fuckpct >= 0.2) && toppct < 0.4)
      || (sd >= 2 && passesMajority) || fuckpct >= 0.5) {
      verdictTierName = "Fuck";
    } else if (currentTier != "Fuck") {
      const verdictIdx = (splitLeftTotal > splitRightTotal) ? splitLowIdx : splitHighIdx;
      const verdictBaseName = orderedTierNames[verdictIdx];
      let arrow = "";
      if (currentIdx >= 0) {
        if (verdictIdx < currentIdx) arrow = "↓    ";
        else if (verdictIdx > currentIdx) arrow = "↑    ";
      }
      verdictTierName = arrow + verdictBaseName;
    }
  } else {
    const verdictIdx = (splitLeftTotal > splitRightTotal) ? splitLowIdx : splitHighIdx;
    const verdictBaseName = orderedTierNames[verdictIdx];

    if (isPending) {
      verdictTierName = verdictBaseName;
    } else {
      let arrow = "";
      if (currentIdx >= 0) {
        if (verdictIdx < currentIdx) arrow = "↓    ";
        else if (verdictIdx > currentIdx) arrow = "↑    ";
      }
      verdictTierName = arrow + verdictBaseName;
    }
  }

  // FIX: prevent accidental global
  let canMove = false;

  if (fuckPresent) {
    const fuckShare = (allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0);
    if ((((fuckShare >= 0.2 && toppct < 0.4) || fuckpct >= 0.5) && verdictDiffersFromCurrent)
      || ((passesMajority || (topVsSecond && !isPending)) &&
        passesSplitMajority &&
        passesSplitPct &&
        verdictDiffersFromCurrent)) {
      canMove = true;
    }
  } else {
    if ((passesMajority || (topVsSecond && !isPending)) &&
        passesSplitMajority &&
        passesSplitPct &&
        verdictDiffersFromCurrent) {
      canMove = true;
    }
  }

  out.push([`Place/Move`, canMove ? "YES" : "NO", "", canMove ? verdictTierName : ""]);

  function upTo3dec_(x) {
    return Number(x).toFixed(3).replace(/\.?0+$/, "");
  }

  out.push([`Split`,
    `${splitLowTier}`,
    `${splitHighTier}`,
    `${upTo3dec_(splitLeftTotal)} | ${upTo3dec_(splitRightTotal)}`
  ]);

  // Distribution dump
  out.push(["", "", "", ""]);
  out.push(["Tier", "Weighted", "%", ""]);
  for (const t of orderedTierNames) {
    const w = weightsByTier[t] || 0;
    const denom = tierWeightSum > 0 ? tierWeightSum : 1;
    const share = w / denom;
    out.push([t, w, share, ""]);
  }

  const startRow = 1;
  const startCol = OUTPUT_COL; // E
  tool.getRange(startRow, startCol, out.length, OUTPUT_WIDTH).setValues(out);

  tool.getRange(startRow, startCol + 1, out.length, 1)
    .setNumberFormat("0.###");

  const sdRow = out.findIndex(r => r[0] === "Standard Deviation");
  if (sdRow >= 0) tool.getRange(startRow + sdRow, startCol + 1).setNumberFormat("0.00").setHorizontalAlignment("left");

  const outRow = out.findIndex(r => r[0] === "Outlier %");
  if (outRow >= 0) {
    tool.getRange(startRow + outRow, startCol + 1).setNumberFormat("0.0%").setHorizontalAlignment("left");
    tool.getRange(startRow + outRow + 1, startCol + 3).setHorizontalAlignment("right");
  }

  // Format share column as percent where applicable
  tool.getRange(startRow + 9, startCol + 2, Math.max(0, out.length - 9), 1).setNumberFormat("0.0%");

  formatAnalysisOutput_(
    tool,
    out.length,
    orderedTierNames,
    weightsByTier,
    topTier,
    secondTier,
    splitLowTier,
    splitHighTier,
    passesMajority,
    passesSplitMajority,
    verdictDiffersFromCurrent,
    isPending,
    fuckPresent,
    sd,
    fuckWeight,
    toppct,
    fuckpct,
    allWeight,
    currentTier,
    verdictTier,
    verdictTierName,
    allWeight,
    passesSplitPct,
    splitMarginPct
  );
}