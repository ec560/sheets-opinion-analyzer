function pickSplitLowIdx_(orderedTierNames, weightsByTier, topIdx, runnerIdx, currentIdx, centerIdx, isPending) {
  const n = orderedTierNames.length;

  const sumRange = (a, b) => {
    let s = 0;
    for (let i = a; i <= b; i++) s += (weightsByTier[orderedTierNames[i]] || 0);
    return s;
  };

  const boundaryScore = (lowIdx) => {
    const leftTotal = sumRange(0, lowIdx);
    const rightTotal = sumRange(lowIdx + 1, n - 1);
    return { leftTotal, rightTotal, score: Math.abs(leftTotal - rightTotal) };
  };

  const areAdjacent = (a, b) => a >= 0 && b >= 0 && Math.abs(a - b) === 1;

  // Constrain split to include the current placed tier unless:
  // - is Pending
  // OR
  // - placed tier is > 2 tiers away from the center (median)
  const currentIsTopOrRunner = (currentIdx === topIdx || currentIdx === runnerIdx);

  const constrain =
    !isPending &&
    currentIdx >= 0 &&
    currentIsTopOrRunner &&     // only anchor if current is top/runner-up
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

  // adjacent top/runner forced, else pick boundary adjacent to top that best balances the split
  if (areAdjacent(topIdx, runnerIdx)) return Math.min(topIdx, runnerIdx);

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