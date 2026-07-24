// functions related to color configurations and reading cells with color values

// maps difficulty color to tier name
function hex_(color) {
  if (color == null) return "";
  // If it's already a hex string like "#ff0000"
  if (typeof color === "string") {
    return color.trim().toLowerCase();
  }

  // Google Apps Script Color object (from getFontColorObjects)
  try {
    if (typeof color.asRgbColor === "function") {
      return color.asRgbColor().asHexString().toLowerCase();
    }
  } catch (e) {}

  // Fallback
  return String(color).trim().toLowerCase();
}

function tierTextColor_(tierName) {
  if (tierName === "Fuck") return "#ff0000"; // red on black
  return tierFontColors[tierName] || "#000000";
}

function configuredOpinionFontColor_(fillColor, fallbackFontColor) {
  const fill = hex_(fillColor);
  if (fill === "#000000") return hex_(fallbackFontColor) || "#000000";
  return opinionFontColorsByFill[fill] || hex_(fallbackFontColor) || "#000000";
}

function buildTierNameToColor_() {
  const out = {};
  Object.keys(difficultyColorNames).forEach(hex => {
    const name = difficultyColorNames[hex];
    if (!(name in out)) out[name] = hex_(hex); // don't overwrite
  });
  out["Fuck"] = "#000000";
  return out;
}

function tierToSparkColor_(tierName) {
  const nameToColor = buildTierNameToColor_();
  return nameToColor[tierName] || "#999999";
}

// currently not heavily used, but kept
const tierNameFromTextColor = (fontHex) => {
  const h = hex_(fontHex);
  // If the font color matches one of the tier fill colors, treat it as specifying that tier
  return difficultyColorNames[h] || null;
};
