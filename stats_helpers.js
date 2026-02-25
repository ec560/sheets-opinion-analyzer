function hasInsaneTokenNotDemon_(text) {
  const words = String(text || "")
    .toLowerCase()
    // turn anything that isn't a letter into spaces
    .replace(/[^a-z]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  for (let i = 0; i < words.length; i++) {
    if (words[i] === "insane") {
      // if the next token is "demon", it refers to "insane demon"
      if (words[i + 1] === "demon") {
        return false;
      }
      return true;
    }
  }
  return false;
}

function mean_(xs) {
  if (!xs || xs.length === 0) return 0;
  let s = 0;
  for (let i = 0; i < xs.length; i++) s += xs[i];
  return s / xs.length;
}

function median_(xs) {
  if (!xs || xs.length === 0) return 0;
  const a = xs.slice().sort((x, y) => x - y);
  const n = a.length;
  const mid = Math.floor(n / 2);
  return (n % 2) ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

// Turns an index into tier label
// ex: 4.0 becomes "Hard"
function tierLabelFromIndex_(orderedTierNames, idxFloat) {
  const n = orderedTierNames.length;
  if (!n) return "";

  const x = Number(idxFloat);
  if (!isFinite(x)) return "";

  const i = clamp_(Math.round(x), 0, n - 1);
  return orderedTierNames[i] || "";
}

// Outlier rule: farther than k tiers from the median/mean
function outlierPct_(rawPoints, centerIdx, k) {
  const total = rawPoints.length;
  if (!total) return { count: 0, total: 0, pct: 0 };

  const cutoff = Math.abs(k);
  let outliers = 0;

  for (const p of rawPoints) {
    if (Math.abs(p - centerIdx) > cutoff) outliers++;
  }

  return { count: outliers, total, pct: outliers / total };
}

function buildCumEnds_(orderedTierNames, weightsByTier) {
  const cumEnds = [];
  let cum = 0;
  for (let i = 0; i < orderedTierNames.length; i++) {
    cum += (weightsByTier[orderedTierNames[i]] || 0);
    cumEnds.push(cum);
  }
  return cumEnds; // cumEnds[i] = cumulative after tier i
}

function intervalDistance_(half, low, high) {
  if (half < low) return low - half;
  if (half > high) return half - high;
  return 0; // half inside interval
}

// pulled out from analyzeSelectedLevel (same logic)
function outlierCutoffLabels_(orderedTierNames, centerIdx, k) {
  const n = orderedTierNames.length;

  // clamps values; does not round
  const clampF = (x, lo, hi) => Math.max(lo, Math.min(hi, x));

  const labelAtGrid = (g) => {
    const gi = clampF(g, 0, n - 1);
    const lo = Math.floor(gi);
    const frac = gi - lo;

    if (Math.abs(frac) < 1e-9) return orderedTierNames[lo] || "";

    const hi = Math.min(lo + 1, n - 1);
    return `${orderedTierNames[lo]} / ${orderedTierNames[hi]}`;
  };

  const lowBoundary  = Math.floor((centerIdx - k) * 2) / 2; // down to nearest 0.5
  const highBoundary = Math.ceil ((centerIdx + k) * 2) / 2; // up to nearest 0.5

  return {
    lowBoundary,
    highBoundary,
    lowLabel: labelAtGrid(lowBoundary),
    highLabel: labelAtGrid(highBoundary),
  };
}