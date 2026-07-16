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
      countedRowFlags[r] = true;

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
  let splitThreshold = 3;
  if (allWeight >= 50 && allWeight < 100) splitThreshold = 4;
  else if (allWeight >= 100) splitThreshold = 5;
  const requiredSplitPct = 0.05;
  const tierSplit = buildTierSplit_(
    orderedTierNames,
    weightsByTier,
    splitLowIdx,
    totalWeight,
    splitThreshold,
    requiredSplitPct
  );

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
  const decisionTierSplit = buildTierSplit_(
    decisionOrderedTierNames,
    decisionWeightsByTier,
    decisionSplitLowIdx,
    totalWeight,
    splitThreshold,
    requiredSplitPct
  );
  const decisionLockSideTotal = decisionTierSplit.winner.weight;
  const lockSharePct = allWeight > 0 ? decisionLockSideTotal / allWeight : 0;
  const isLockWorthy = lockSharePct >= FLAG_LOCK_SHARE &&
    rawCount >= FLAG_LOCK_MIN_RAW_COUNT &&
    allWeight >= FLAG_LOCK_MIN_WEIGHT;

  const toppct = allWeight > 0 ? topWeight / allWeight : 0;
  const fuckpct = allAndFuck > 0 ? fuckWeight / (totalWeight || 1) : 0;
  const hasMinimumTotalOpinions = allWeight >= FLAG_LOW_OPINION_WEIGHT;
  const passesMajority = hasMinimumTotalOpinions;
  const fuckRules = evaluateFuckRules_(fuckPresent, fuckpct, toppct, sd, passesMajority);

  const splitVerdictTier = decisionTierSplit.winner.label;
  const verdictTier = fuckRules.isFuckVerdict ? "Fuck" : splitVerdictTier;

  const verdictDiffersFromCurrent = verdictTier !== "" && verdictTier !== decisionCurrentTier;
  const passesSplitMajority = decisionTierSplit.passesWeightThreshold;
  const passesSplitPct = decisionTierSplit.passesPctThreshold;
  const verdictIdx = decisionTierSplit.winner.index;
  const verdictBaseName = decisionOrderedTierNames[verdictIdx];
  const requiresFuckPlacementMargin = fuckPresent && verdictTier !== "Fuck";
  const fuckPlacementComparison = requiresFuckPlacementMargin
    ? buildFuckPlacementComparison_(decisionTierSplit, verdictBaseName, fuckWeight, totalWeight, splitThreshold)
    : null;
  const passesFuckPlacementMargin = !fuckPlacementComparison || fuckPlacementComparison.passesWeightThreshold;
  const placementGate = selectPlacementComparison_(decisionTierSplit, fuckPlacementComparison, "");

  let canMove = false;
  if (fuckPresent) {
    if (
      hasMinimumTotalOpinions &&
      (
        (fuckRules.hasShareBasedMovementSignal && verdictDiffersFromCurrent) ||
        (placementGate.passesWeightThreshold && placementGate.passesPctThreshold && verdictDiffersFromCurrent)
      )
    ) {
      canMove = true;
    }
  } else if (
    hasMinimumTotalOpinions &&
    placementGate.passesWeightThreshold &&
    placementGate.passesPctThreshold &&
    verdictDiffersFromCurrent
  ) {
    canMove = true;
  }

  let verdictTierName = verdictTier;
  if (fuckPresent) {
    if (fuckRules.isFuckVerdict) {
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
    decisionTierSplit.lowIndex,
    decisionTierSplit.highIndex,
    decisionTierSplit.left.weight,
    decisionTierSplit.right.weight
  );
  const bookAlert = buildBookAlertFromPlacementSplit_(difference, decisionCurrentIdx, decisionTierSplit.lowIndex);
  const lockSideLabel = decisionTierSplit.left.weight >= decisionTierSplit.right.weight
    ? decisionTierSplit.left.label
    : decisionTierSplit.right.label;
  const moveDirection = (verdictTier !== "Fuck" && decisionCurrentIdx >= 0)
    ? movementDirection_(decisionCurrentIdx, verdictIdx)
    : "";
  const moveFailureReason = determinePlaceMoveFailureReason_({
    totalWeightedOpinions: allWeight,
    fuckRules,
    isPending,
    passesMajority,
    passesSplitMajority,
    verdictDiffersFromCurrent,
    passesFuckPlacementMargin,
    passesSplitPct,
    isLockWorthy
  });
  const placementComparison = selectPlacementComparison_(
    decisionTierSplit,
    fuckPlacementComparison,
    moveFailureReason
  );

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
    tierSplit,
    decisionTierSplit,
    fuckPlacementComparison,
    placementComparison,
    currentTier,
    currentIdx,
    isPending,
    decisionCurrentTier,
    decisionCurrentIdx,
    toppct,
    fuckpct,
    fuckRules,
    passesMajority,
    verdictTier,
    verdictDiffersFromCurrent,
    splitThreshold,
    passesSplitMajority,
    passesSplitPct,
    requiresFuckPlacementMargin,
    passesFuckPlacementMargin,
    verdictIdx,
    verdictBaseName,
    verdictTierName,
    canMove,
    lockSharePct,
    isLockWorthy,
    lockSideLabel,
    difference,
    bookAlert,
    moveDirection,
    moveFailureReason
  };
}

function buildTierSplit_(orderedNames, weightsByTier, lowIndex, totalWeight, requiredMargin, requiredMarginPct) {
  const highIndex = lowIndex + 1;
  let leftWeight = 0;
  let rightWeight = 0;

  for (let i = 0; i <= lowIndex; i++) leftWeight += (weightsByTier[orderedNames[i]] || 0);
  for (let i = highIndex; i < orderedNames.length; i++) rightWeight += (weightsByTier[orderedNames[i]] || 0);

  const left = { label: orderedNames[lowIndex], weight: leftWeight, index: lowIndex };
  const right = { label: orderedNames[highIndex], weight: rightWeight, index: highIndex };
  const winnerSide = leftWeight > rightWeight ? "left" : "right";
  const winner = winnerSide === "left" ? left : right;
  const loser = winnerSide === "left" ? right : left;
  const marginWeight = Math.abs(leftWeight - rightWeight);
  const marginPct = totalWeight > 0 ? marginWeight / totalWeight : 0;

  return {
    kind: "tier-vs-tier",
    lowIndex,
    highIndex,
    left,
    right,
    winnerSide,
    winner,
    loser,
    marginWeight,
    decisionMarginWeight: marginWeight,
    marginPct,
    requiredMargin,
    requiredMarginPct,
    passesWeightThreshold: marginWeight >= requiredMargin,
    passesPctThreshold: marginPct >= requiredMarginPct
  };
}

function buildFuckPlacementComparison_(tierSplit, verdictBaseName, fuckWeight, totalWeight, requiredMargin) {
  const tierSide = {
    label: verdictBaseName,
    weight: tierSplit.winner.weight,
    index: tierSplit.winner.index
  };
  const fuckSide = { label: "Fuck", weight: fuckWeight, index: -1 };
  const tierIsLeft = tierSplit.winnerSide === "left";
  const left = tierIsLeft ? tierSide : fuckSide;
  const right = tierIsLeft ? fuckSide : tierSide;
  const winnerSide = left.weight > right.weight ? "left" : "right";
  const winner = winnerSide === "left" ? left : right;
  const loser = winnerSide === "left" ? right : left;
  const decisionMarginWeight = tierSide.weight - fuckSide.weight;
  const marginWeight = Math.abs(decisionMarginWeight);

  return {
    kind: "tier-vs-fuck",
    left,
    right,
    winnerSide,
    winner,
    loser,
    decisionSide: tierIsLeft ? "left" : "right",
    decisionLabel: verdictBaseName,
    marginWeight,
    decisionMarginWeight,
    marginPct: totalWeight > 0 ? marginWeight / totalWeight : 0,
    requiredMargin,
    requiredMarginPct: 0,
    passesWeightThreshold: decisionMarginWeight >= requiredMargin,
    passesPctThreshold: true
  };
}

function selectPlacementComparison_(tierSplit, fuckComparison, moveFailureReason) {
  let active = tierSplit;

  if (
    fuckComparison &&
    moveFailureReason !== "low_split_margin" &&
    fuckComparison.decisionMarginWeight < tierSplit.decisionMarginWeight
  ) {
    active = fuckComparison;
  }

  return {
    ...active,
    tierSplit,
    fuckComparison,
    passesWeightThreshold: tierSplit.passesWeightThreshold &&
      (!fuckComparison || fuckComparison.passesWeightThreshold),
    passesPctThreshold: tierSplit.passesPctThreshold
  };
}

function evaluateFuckRules_(fuckPresent, fuckpct, toppct, sd, passesMajority) {
  const meetsShareThreshold = fuckpct >= 0.25;
  const hasLowTopShare = toppct < 0.35;
  const hasMajorityFuckShare = fuckpct >= 0.5;
  const hasHighVolatility = sd >= 2;
  const hasLowTopShareSignal = meetsShareThreshold && hasLowTopShare;
  const hasVolatilityVerdictSignal = hasHighVolatility && passesMajority;

  return {
    meetsShareThreshold,
    hasLowTopShare,
    hasMajorityFuckShare,
    hasHighVolatility,
    hasShareBasedMovementSignal: hasLowTopShareSignal || hasMajorityFuckShare,
    isFuckVerdict: !!fuckPresent && (
      hasLowTopShareSignal ||
      hasMajorityFuckShare ||
      hasVolatilityVerdictSignal
    ),
    isFuckMode: !!fuckPresent &&
      (hasMajorityFuckShare || hasLowTopShare) &&
      (meetsShareThreshold || hasHighVolatility)
  };
}

function determinePlaceMoveFailureReason_(ctx) {
  if (ctx.totalWeightedOpinions < 4) return "needs_more_opinions";

  if (ctx.fuckRules.isFuckMode) {
    if (ctx.isPending && !ctx.passesMajority) return "needs_more_opinions";
    if (!ctx.verdictDiffersFromCurrent || ctx.fuckRules.hasMajorityFuckShare) return "no_movement_f";
    if (!ctx.passesSplitMajority && ctx.fuckRules.hasHighVolatility) return "split_not_decisive_f";
    return "fuck_rules_not_met";
  }

  if (ctx.isPending && !ctx.passesSplitMajority) return "split_not_decisive";
  if (ctx.isPending && !ctx.passesMajority) return "needs_more_opinions";
  if (!ctx.passesSplitMajority) return "split_not_decisive";
  if (!ctx.passesMajority) return "needs_more_opinions";
  if (!ctx.passesFuckPlacementMargin) return "fuck_placement_not_decisive";
  if (!ctx.passesSplitPct && ctx.verdictDiffersFromCurrent) return "low_split_margin";
  if (ctx.isLockWorthy) return "lock_threshold_met";
  if (!ctx.verdictDiffersFromCurrent) return "no_movement";
  return "rules_not_met";
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
