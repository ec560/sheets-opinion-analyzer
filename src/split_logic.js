function sumTierRange_(orderedTierNames, weightsByTier, startIdx, endIdx) {
  let total = 0;
  for (let i = startIdx; i <= endIdx; i++) {
    total += (weightsByTier[orderedTierNames[i]] || 0);
  }
  return total;
}

function rankTierWeights_(orderedTierNames, weightsByTier) {
  const ranked = orderedTierNames
    .map((name, index) => ({ name, weight: weightsByTier[name] || 0, index }))
    .sort((a, b) => {
      const weightDifference = b.weight - a.weight;
      return weightDifference !== 0 ? weightDifference : a.index - b.index;
    });

  const entries = ranked.map(item => [item.name, item.weight]);
  return {
    entries,
    top: entries[0] || ["(none)", 0],
    second: entries[1] || ["(none)", 0]
  };
}

function buildTierSplit_(orderedTierNames, weightsByTier, lowIndex, totalWeight, requiredMargin, requiredMarginPct) {
  const highIndex = lowIndex + 1;
  const leftWeight = sumTierRange_(orderedTierNames, weightsByTier, 0, lowIndex);
  const rightWeight = sumTierRange_(orderedTierNames, weightsByTier, highIndex, orderedTierNames.length - 1);
  const left = { label: orderedTierNames[lowIndex], weight: leftWeight, index: lowIndex };
  const right = { label: orderedTierNames[highIndex], weight: rightWeight, index: highIndex };
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

function calculateSplit_(
  orderedTierNames,
  weightsByTier,
  ranking,
  currentIdx,
  centerIdx,
  isPending,
  totalWeight,
  requiredMargin,
  requiredMarginPct
) {
  const topIdx = orderedTierNames.indexOf(ranking.top[0]);
  const runnerIdx = orderedTierNames.indexOf(ranking.second[0]);
  const lowIndex = pickSplitLowIdx_(
    orderedTierNames,
    weightsByTier,
    topIdx,
    runnerIdx,
    currentIdx,
    centerIdx,
    isPending
  );

  return buildTierSplit_(
    orderedTierNames,
    weightsByTier,
    lowIndex,
    totalWeight,
    requiredMargin,
    requiredMarginPct
  );
}

function pickSplitLowIdx_(orderedTierNames, weightsByTier, topIdx, runnerIdx, currentIdx, centerIdx, isPending) {
  const n = orderedTierNames.length;

  const boundaryScore = (lowIdx) => {
    const leftTotal = sumTierRange_(orderedTierNames, weightsByTier, 0, lowIdx);
    const rightTotal = sumTierRange_(orderedTierNames, weightsByTier, lowIdx + 1, n - 1);
    return { leftTotal, rightTotal, score: Math.abs(leftTotal - rightTotal) };
  };

  const areAdjacent = (a, b) => a >= 0 && b >= 0 && Math.abs(a - b) === 1;

  // Constrain split to include the current placed tier unless:
  // - is Pending
  // - placed tier is > 2 tiers away from the center (median)
  const constrain =
    !isPending &&
    currentIdx >= 0 &&
    Math.abs(currentIdx - centerIdx) <= 2;

  if (constrain) {
    const candidates = [];

    // (current-1 | current)
    if (currentIdx > 0) {
      const lowIdx = currentIdx - 1;
      candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
    }

    // (current | current+1)
    if (currentIdx < n - 1) {
      const lowIdx = currentIdx;
      candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
    }

    candidates.sort((a, b) => a.score - b.score);
    return candidates.length ? candidates[0].lowIdx : Math.max(0, Math.min(n - 2, currentIdx));
  }

  // Placed levels: adjacent top/runner forced
  // Pending levels: closer side around the winning tier
  if (!isPending && areAdjacent(topIdx, runnerIdx)) return Math.min(topIdx, runnerIdx);

  const candidates = [];

  if (topIdx > 0) {
    const lowIdx = topIdx - 1;
    candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
  }
  if (topIdx < n - 1) {
    const lowIdx = topIdx;
    candidates.push({ lowIdx, ...boundaryScore(lowIdx) });
  }

  candidates.sort((a, b) => a.score - b.score);
  return candidates.length ? candidates[0].lowIdx : Math.max(0, topIdx - 1);
}
