function calculateLevelAnalysis_(tierName, levelName, vals, bgs, fcs) {
  const weightsByTier = {};
  for (const n of orderedTierNames) weightsByTier[n] = 0;

  let fuckPresent = false;
  let fuckWeight = 0;
  let totalWeight = 0;
  let rawCount = 0;
  const rawPoints = [];
  const countedRowFlags = new Array(vals.length).fill(false);

  const tierIdxOf = (name) => orderedTierNames.indexOf(name);
  const addTierWeight = (tier, wt) => {
    if (!tier || wt <= 0) return;
    weightsByTier[tier] = (weightsByTier[tier] || 0) + wt;
  };

  for (let r = 0; r < vals.length; r++) {
    const player = String(vals[r][0] ?? "").trim();
    const opinionText = String(vals[r][1] ?? "").trim();
    const relText = String(vals[r][2] ?? "").trim();

    if (!player && !opinionText && !relText) continue;

    const opinionColor = hex_(bgs[r][1]);
    const relColor = hex_(bgs[r][2]);
    const w = reliabilityFactors[relColor] ?? 0;
    const opinionFont = hex_(fcs[r][1]);

    if (w <= 0) continue;

    if (opinionColor === "#000000") {
      fuckPresent = true;
      rawCount++;
      fuckWeight += w;

      const fontKey = hex_(opinionFont);
      if (splitPairs[fontKey]) {
        const [lowHex, highHex] = splitPairs[fontKey].map(hex_);
        const lowTier = difficultyColorNames[lowHex];
        const highTier = difficultyColorNames[highHex];
        const i1 = tierIdxOf(lowTier);
        const i2 = tierIdxOf(highTier);

        if (i1 >= 0 && i2 >= 0) rawPoints.push((i1 + i2) / 2);
        if (lowTier && highTier) {
          countedRowFlags[r] = true;
          const halfW = w / 2;
          addTierWeight(lowTier, halfW);
          addTierWeight(highTier, halfW);
          totalWeight += w;
        }
        continue;
      }

      const fontTier = difficultyColorNames[fontKey];
      if (fontTier && fontTier !== "Insane") {
        const it = tierIdxOf(fontTier);
        if (it >= 0) rawPoints.push(it);
        countedRowFlags[r] = true;
        addTierWeight(fontTier, w);
        totalWeight += w;
        continue;
      }

      if (String(opinionText).toLowerCase().includes("insane")) {
        const it = tierIdxOf("Insane");
        if (it >= 0) rawPoints.push(it);
        countedRowFlags[r] = true;
        addTierWeight("Insane", w);
        totalWeight += w;
        continue;
      }

      totalWeight += w;
      continue;
    }

    if (difficultyColorNames[opinionColor]) {
      const tName = difficultyColorNames[opinionColor];
      const it = tierIdxOf(tName);
      if (it >= 0) rawPoints.push(it);

      countedRowFlags[r] = true;
      addTierWeight(tName, w);
      rawCount++;
      totalWeight += w;
      continue;
    }

    const oc = hex_(opinionColor);
    if (splitPairs[oc]) {
      const [c1Raw, c2Raw] = splitPairs[oc];
      const c1 = hex_(c1Raw);
      const c2 = hex_(c2Raw);
      const t1 = difficultyColorNames[c1];
      const t2 = difficultyColorNames[c2];
      const i1 = tierIdxOf(t1);
      const i2 = tierIdxOf(t2);

      if (i1 >= 0 && i2 >= 0) rawPoints.push((i1 + i2) / 2);
      if (t1 && t2) {
        countedRowFlags[r] = true;
        const halfW = w / 2;
        addTierWeight(t1, halfW);
        addTierWeight(t2, halfW);
        rawCount++;
        totalWeight += w;
      }
    }
  }

  const rawMeanIdx = mean_(rawPoints);
  const rawMedianIdx = median_(rawPoints);
  const rawMeanLabel = tierLabelFromIndex_(orderedTierNames, rawMeanIdx);
  const rawMedianLabel = tierLabelFromIndex_(orderedTierNames, rawMedianIdx);

  const entries = Object.entries(weightsByTier).sort((a, b) => b[1] - a[1]);
  const top = entries[0] || ["(none)", 0];
  const second = entries[1] || ["(none)", 0];
  const topTier = top[0];
  const topWeight = top[1];
  const secondTier = second[0];
  const secondWeight = second[1];

  const tierWeightSum = orderedTierNames.reduce((s, t) => s + (weightsByTier[t] || 0), 0);
  const allWeight = tierWeightSum;
  const allAndFuck = allWeight + fuckWeight;

  const idx = {};
  orderedTierNames.forEach((t, i) => idx[t] = i);
  let sw = 0;
  let mu = 0;
  orderedTierNames.forEach(t => {
    const wt = weightsByTier[t] || 0;
    sw += wt;
    mu += wt * idx[t];
  });
  mu = sw ? mu / sw : 0;

  let v = 0;
  orderedTierNames.forEach(t => {
    const wt = weightsByTier[t] || 0;
    const d = idx[t] - mu;
    v += wt * d * d;
  });
  const sd = sw ? Math.sqrt(v / sw) : 0;

  let outlierK = 1.5;
  if (sd > 1.5) outlierK = 2.5;
  else if (sd > 1) outlierK = 2;

  const outlierPctObj = outlierPct_(rawPoints, rawMeanIdx, outlierK);
  const outlierBounds = outlierCutoffLabels_(orderedTierNames, rawMeanIdx, outlierK);

  const sumRange = (names, weightMap, startIdx, endIdx) => {
    let s = 0;
    for (let i = startIdx; i <= endIdx; i++) s += (weightMap[names[i]] || 0);
    return s;
  };

  const currentTier = String(tierName || "").trim();
  const currentIdx = orderedTierNames.indexOf(currentTier);
  const isPending = currentTier.toLowerCase() === "pending";
  const topIdx = orderedTierNames.indexOf(topTier);
  const runnerIdx = orderedTierNames.indexOf(secondTier);
  const splitLowIdx = pickSplitLowIdx_(
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
  const splitLeftTotal = sumRange(orderedTierNames, weightsByTier, 0, splitLowIdx);
  const splitRightTotal = sumRange(orderedTierNames, weightsByTier, splitHighIdx, orderedTierNames.length - 1);

  const decisionTierName = (name) => {
    return (name === "Beginner" || name === "Insane Demon") ? "Beginner" : name;
  };
  const decisionOrderedTierNames = orderedTierNames.filter(name => name !== "Insane Demon");
  const decisionWeightsByTier = {};
  for (const n of decisionOrderedTierNames) decisionWeightsByTier[n] = 0;
  for (const n of orderedTierNames) {
    const decisionName = decisionTierName(n);
    if (decisionName in decisionWeightsByTier) {
      decisionWeightsByTier[decisionName] += (weightsByTier[n] || 0);
    }
  }

  const decisionEntries = Object.entries(decisionWeightsByTier).sort((a, b) => b[1] - a[1]);
  const decisionTop = decisionEntries[0] || ["(none)", 0];
  const decisionSecond = decisionEntries[1] || ["(none)", 0];
  const decisionTopTier = decisionTop[0];
  const decisionTopWeight = decisionTop[1];
  const decisionSecondTier = decisionSecond[0];
  const decisionSecondWeight = decisionSecond[1];
  const decisionTopIdx = decisionOrderedTierNames.indexOf(decisionTopTier);
  const decisionRunnerIdx = decisionOrderedTierNames.indexOf(decisionSecondTier);
  const decisionCurrentTier = isPending ? currentTier : decisionTierName(currentTier);
  const decisionCurrentIdx = decisionOrderedTierNames.indexOf(decisionCurrentTier);
  const decisionCenterTier = decisionTierName(rawMedianLabel);
  const decisionCenterIdx = decisionOrderedTierNames.indexOf(decisionCenterTier);
  const decisionSplitLowIdx = pickSplitLowIdx_(
    decisionOrderedTierNames,
    decisionWeightsByTier,
    decisionTopIdx,
    decisionRunnerIdx,
    decisionCurrentIdx,
    decisionCenterIdx,
    isPending
  );
  const decisionSplitHighIdx = decisionSplitLowIdx + 1;
  const decisionSplitLeftTotal = sumRange(decisionOrderedTierNames, decisionWeightsByTier, 0, decisionSplitLowIdx);
  const decisionSplitRightTotal = sumRange(
    decisionOrderedTierNames,
    decisionWeightsByTier,
    decisionSplitHighIdx,
    decisionOrderedTierNames.length - 1
  );
  const decisionLockSideTotal = Math.max(decisionSplitLeftTotal, decisionSplitRightTotal);
  const lockSharePct = allWeight > 0 ? decisionLockSideTotal / allWeight : 0;
  const isLockWorthy = lockSharePct >= FLAG_LOCK_SHARE &&
    rawCount >= FLAG_LOCK_MIN_RAW_COUNT &&
    allWeight >= FLAG_LOCK_MIN_WEIGHT;

  const toppct = allWeight > 0 ? topWeight / allWeight : 0;
  const fuckpct = allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0;
  const hasMinimumTotalOpinions = allWeight >= FLAG_LOW_OPINION_WEIGHT;
  const passesMajority = hasMinimumTotalOpinions;

  let verdictTier = "";
  if (fuckPresent) {
    if ((((fuckpct >= 0.2 && toppct < 0.35) || fuckpct >= 0.5) || (sd >= 2 && passesMajority))) {
      verdictTier = "Fuck";
    } else {
      verdictTier = (decisionSplitLeftTotal > decisionSplitRightTotal)
        ? decisionOrderedTierNames[decisionSplitLowIdx]
        : decisionOrderedTierNames[decisionSplitHighIdx];
    }
  } else {
    verdictTier = (decisionSplitLeftTotal > decisionSplitRightTotal)
      ? decisionOrderedTierNames[decisionSplitLowIdx]
      : decisionOrderedTierNames[decisionSplitHighIdx];
  }

  const verdictDiffersFromCurrent = verdictTier !== "" && verdictTier !== decisionCurrentTier;
  let splitThreshold = 3;
  if (allWeight >= 50 && allWeight < 100) splitThreshold = 4;
  else if (allWeight >= 100) splitThreshold = 5;

  const splitMargin = Math.abs(decisionSplitLeftTotal - decisionSplitRightTotal);
  const splitMarginPct = totalWeight > 0 ? (splitMargin / totalWeight) : 0;
  const passesSplitMajority = splitMargin >= splitThreshold;
  const passesSplitPct = isPending ? (splitMarginPct >= 0.20) : (splitMarginPct >= 0.05);
  const verdictIdx = (decisionSplitLeftTotal > decisionSplitRightTotal) ? decisionSplitLowIdx : decisionSplitHighIdx;
  const verdictBaseName = decisionOrderedTierNames[verdictIdx];
  const decisionWinningSideTotal = verdictIdx === decisionSplitLowIdx
    ? decisionSplitLeftTotal
    : decisionSplitRightTotal;
  const decisionLosingSideTotal = verdictIdx === decisionSplitLowIdx
    ? decisionSplitRightTotal
    : decisionSplitLeftTotal;
  const requiresFuckPlacementMargin = fuckPresent && verdictTier !== "Fuck";
  const fuckPlacementMargin = decisionWinningSideTotal - fuckWeight;
  const passesFuckPlacementMargin = !requiresFuckPlacementMargin || fuckPlacementMargin >= splitThreshold;
  const showsFuckPlacementSplit = requiresFuckPlacementMargin && fuckWeight > decisionLosingSideTotal;

  let displaySplitLowTier = splitLowTier;
  let displaySplitHighTier = splitHighTier;
  let displaySplitLeftTotal = splitLeftTotal;
  let displaySplitRightTotal = splitRightTotal;

  if (showsFuckPlacementSplit) {
    if (verdictIdx === decisionSplitLowIdx) {
      displaySplitLowTier = verdictBaseName;
      displaySplitHighTier = "Fuck";
      displaySplitLeftTotal = decisionWinningSideTotal;
      displaySplitRightTotal = fuckWeight;
    } else {
      displaySplitLowTier = "Fuck";
      displaySplitHighTier = verdictBaseName;
      displaySplitLeftTotal = fuckWeight;
      displaySplitRightTotal = decisionWinningSideTotal;
    }
  }

  let canMove = false;
  if (fuckPresent) {
    const fuckShare = (allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0);
    if (
      hasMinimumTotalOpinions &&
      (
        (((fuckShare >= 0.2 && toppct < 0.35) || fuckpct >= 0.5) && verdictDiffersFromCurrent) ||
        (passesSplitMajority && passesSplitPct && verdictDiffersFromCurrent)
      )
    ) {
      canMove = true;
    }
  } else if (
    hasMinimumTotalOpinions &&
    passesSplitMajority &&
    passesSplitPct &&
    verdictDiffersFromCurrent
  ) {
    canMove = true;
  }

  if (canMove && !passesFuckPlacementMargin) {
    canMove = false;
  }

  let verdictTierName = verdictTier;
  if (fuckPresent) {
    if (((fuckpct >= 0.2) && toppct < 0.35) || (sd >= 2 && passesMajority) || fuckpct >= 0.5) {
      verdictTierName = "Fuck";
    } else if (currentTier !== "Fuck") {
      let arrow = "";
      if (decisionCurrentIdx >= 0) {
        if (verdictIdx < currentIdx) arrow = "\u2193    ";
        else if (verdictIdx > currentIdx) arrow = "\u2191    ";
      }
      verdictTierName = arrow + verdictBaseName;
    }
  } else if (isPending) {
    verdictTierName = verdictBaseName;
  } else {
    let arrow = "";
    if (decisionCurrentIdx >= 0) {
      if (verdictIdx < currentIdx) arrow = "\u2193    ";
      else if (verdictIdx > currentIdx) arrow = "\u2191    ";
    }
    verdictTierName = arrow + verdictBaseName;
  }

  const difference = buildSplitDifference_(
    decisionWeightsByTier,
    decisionOrderedTierNames,
    decisionCurrentIdx,
    decisionSplitLowIdx,
    decisionSplitHighIdx,
    decisionSplitLeftTotal,
    decisionSplitRightTotal
  );
  const bookAlert = buildBookAlertFromPlacementSplit_(difference, decisionCurrentIdx, decisionSplitLowIdx);
  const lockSideLabel = decisionSplitLeftTotal >= decisionSplitRightTotal
    ? decisionOrderedTierNames[decisionSplitLowIdx]
    : decisionOrderedTierNames[decisionSplitHighIdx];
  const moveDirection = (verdictTier !== "Fuck" && decisionCurrentIdx >= 0)
    ? movementDirection_(decisionCurrentIdx, verdictIdx)
    : "";

  return {
    tierName,
    levelName,
    weightsByTier,
    countedRowFlags,
    rawPoints,
    rawCount,
    rawMeanIdx,
    rawMeanLabel,
    rawMedianIdx,
    rawMedianLabel,
    outlierK,
    outlierPctObj,
    outlierBounds,
    sd,
    fuckPresent,
    fuckWeight,
    totalWeight,
    tierWeightSum,
    allWeight,
    allAndFuck,
    topTier,
    topWeight,
    secondTier,
    secondWeight,
    decisionOrderedTierNames,
    decisionWeightsByTier,
    decisionTopTier,
    decisionTopWeight,
    decisionSecondTier,
    decisionSecondWeight,
    minimumOpinionWeight: decisionTopWeight,
    splitLowIdx,
    splitHighIdx,
    splitLowTier,
    splitHighTier,
    splitLeftTotal,
    splitRightTotal,
    displaySplitLowTier,
    displaySplitHighTier,
    displaySplitLeftTotal,
    displaySplitRightTotal,
    currentTier,
    currentIdx,
    isPending,
    decisionCurrentTier,
    decisionCurrentIdx,
    decisionSplitLowIdx,
    decisionSplitHighIdx,
    decisionSplitLeftTotal,
    decisionSplitRightTotal,
    toppct,
    fuckpct,
    passesMajority,
    verdictTier,
    verdictDiffersFromCurrent,
    splitThreshold,
    splitMargin,
    splitMarginPct,
    passesSplitMajority,
    passesSplitPct,
    decisionWinningSideTotal,
    decisionLosingSideTotal,
    requiresFuckPlacementMargin,
    fuckPlacementMargin,
    passesFuckPlacementMargin,
    showsFuckPlacementSplit,
    verdictIdx,
    verdictBaseName,
    verdictTierName,
    canMove,
    lockSharePct,
    isLockWorthy,
    lockSideLabel,
    difference,
    bookAlert,
    moveDirection
  };
}

function buildSplitDifference_(
  weightsByTier,
  orderedNames,
  currentIdx,
  splitLowIdx,
  splitHighIdx,
  leftTotal,
  rightTotal
) {
  const lean = Math.abs(leftTotal - rightTotal);
  const winningSide = leftTotal > rightTotal ? "left" : "right";
  const startIdx = winningSide === "left" ? 0 : splitHighIdx;
  const endIdx = winningSide === "left" ? splitLowIdx : orderedNames.length - 1;
  const target = strongestTierInRange_(weightsByTier, orderedNames, startIdx, endIdx, currentIdx);

  return {
    direction: winningSide === "right" ? "green" : "red",
    side: winningSide === "right" ? "higher side" : "lower side",
    tier: target.tier,
    weight: target.weight,
    winningSide,
    lean
  };
}

function buildBookAlertFromPlacementSplit_(difference, currentIdx, splitLowIdx) {
  if (!difference || currentIdx < 0) return null;
  if (difference.lean < FLAG_BOOK_LEAN_WEIGHT) return null;

  const currentSide = currentIdx <= splitLowIdx ? "left" : "right";
  if (difference.winningSide === currentSide) return null;

  return difference;
}

function strongestTierInRange_(weightsByTier, orderedNames, startIdx, endIdx, currentIdx) {
  let best = { tier: "", weight: 0, distance: Number.MAX_VALUE };

  for (let i = startIdx; i <= endIdx; i++) {
    const tier = orderedNames[i];
    if (!tier) continue;

    const weight = weightsByTier[tier] || 0;
    const distance = Math.abs(i - currentIdx);
    if (
      !best.tier ||
      weight > best.weight ||
      (weight === best.weight && distance < best.distance)
    ) {
      best = { tier, weight, distance };
    }
  }

  return best;
}

function movementDirection_(currentIdx, verdictIdx) {
  if (verdictIdx < currentIdx) return "down";
  if (verdictIdx > currentIdx) return "up";
  return "";
}
