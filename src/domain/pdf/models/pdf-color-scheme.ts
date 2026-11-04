/**
 * Domain models and services for Zotero color palettes and customizable PDF color schemes.
 * Supports semantic color tagging (e.g., Green = "New Info", Red = "Important Points", Yellow = "Processes"),
 * color scheme templates/profiles, per-PDF color scheme overrides, and plugin settings persistence.
 */

import { PdfRgbColor, hexToRgb, rgbToHex } from "./pdf-annotation";

export interface PdfTaggedColor {
  id: string;
  hex: string; // e.g. "#ffd400"
  name: string; // Display name, e.g. "Yellow"
  tag: string; // Semantic tag, e.g. "Processes", "Important Points", "New Info"
  rgb: PdfRgbColor; // Normalized [r, g, b] (0.0 - 1.0)
}

export interface PdfColorSchemeProfile {
  id: string; // Unique profile identifier
  name: string; // Display profile name, e.g. "Academic Research"
  description?: string;
  colors: PdfTaggedColor[];
  isBuiltIn?: boolean;
  modifiedAt?: number;
}

/** Official Zotero Color Palette Definitions */
export const ZOTERO_HEX_PALETTE = {
  yellow: "#ffd400",
  red: "#ff6666",
  green: "#5fb236",
  blue: "#2ea8e5",
  purple: "#a28ae5",
  magenta: "#e56eee",
  orange: "#f19837",
  gray: "#aaaaaa",
} as const;

/** Built-in Template: Zotero Standard */
export const ZOTERO_DEFAULT_PROFILE: PdfColorSchemeProfile = {
  id: "zotero-default",
  name: "Zotero Standard",
  description: "Standard Zotero palette with versatile academic & research tags",
  isBuiltIn: true,
  colors: [
    {
      id: "green",
      hex: ZOTERO_HEX_PALETTE.green,
      name: "Green",
      tag: "New Info",
      rgb: [0.37, 0.70, 0.22],
    },
    {
      id: "red",
      hex: ZOTERO_HEX_PALETTE.red,
      name: "Red",
      tag: "Important Points",
      rgb: [1.0, 0.40, 0.40],
    },
    {
      id: "yellow",
      hex: ZOTERO_HEX_PALETTE.yellow,
      name: "Yellow",
      tag: "Processes",
      rgb: [1.0, 0.83, 0.0],
    },
    {
      id: "blue",
      hex: ZOTERO_HEX_PALETTE.blue,
      name: "Blue",
      tag: "Definitions & Questions",
      rgb: [0.18, 0.66, 0.90],
    },
    {
      id: "purple",
      hex: ZOTERO_HEX_PALETTE.purple,
      name: "Purple",
      tag: "Theories & Claims",
      rgb: [0.64, 0.54, 0.90],
    },
    {
      id: "magenta",
      hex: ZOTERO_HEX_PALETTE.magenta,
      name: "Magenta",
      tag: "Methodology & Proofs",
      rgb: [0.90, 0.43, 0.93],
    },
    {
      id: "orange",
      hex: ZOTERO_HEX_PALETTE.orange,
      name: "Orange",
      tag: "Caveats & Objections",
      rgb: [0.95, 0.60, 0.22],
    },
    {
      id: "gray",
      hex: ZOTERO_HEX_PALETTE.gray,
      name: "Gray",
      tag: "References & Context",
      rgb: [0.67, 0.67, 0.67],
    },
  ],
};

/** Built-in Template: Academic Research */
export const ACADEMIC_RESEARCH_PROFILE: PdfColorSchemeProfile = {
  id: "academic-research",
  name: "Academic Research",
  description: "Tailored for literature review, experimental analysis, and thesis papers",
  isBuiltIn: true,
  colors: [
    {
      id: "red",
      hex: ZOTERO_HEX_PALETTE.red,
      name: "Red",
      tag: "Core Thesis / Claim",
      rgb: [1.0, 0.40, 0.40],
    },
    {
      id: "green",
      hex: ZOTERO_HEX_PALETTE.green,
      name: "Green",
      tag: "Key Evidence / Finding",
      rgb: [0.37, 0.70, 0.22],
    },
    {
      id: "yellow",
      hex: ZOTERO_HEX_PALETTE.yellow,
      name: "Yellow",
      tag: "Methodology & Setup",
      rgb: [1.0, 0.83, 0.0],
    },
    {
      id: "blue",
      hex: ZOTERO_HEX_PALETTE.blue,
      name: "Blue",
      tag: "Related Literature",
      rgb: [0.18, 0.66, 0.90],
    },
    {
      id: "purple",
      hex: ZOTERO_HEX_PALETTE.purple,
      name: "Purple",
      tag: "Theoretical Framework",
      rgb: [0.64, 0.54, 0.90],
    },
    {
      id: "orange",
      hex: ZOTERO_HEX_PALETTE.orange,
      name: "Orange",
      tag: "Limitation / Counterargument",
      rgb: [0.95, 0.60, 0.22],
    },
    {
      id: "magenta",
      hex: ZOTERO_HEX_PALETTE.magenta,
      name: "Magenta",
      tag: "Novel Contribution",
      rgb: [0.90, 0.43, 0.93],
    },
    {
      id: "gray",
      hex: ZOTERO_HEX_PALETTE.gray,
      name: "Gray",
      tag: "Future Work & Citation",
      rgb: [0.67, 0.67, 0.67],
    },
  ],
};

/** Built-in Template: Literature & Book Review */
export const LITERATURE_REVIEW_PROFILE: PdfColorSchemeProfile = {
  id: "literature-review",
  name: "Literature & Book Review",
  description: "Designed for reading novels, essays, and humanistic analysis",
  isBuiltIn: true,
  colors: [
    {
      id: "yellow",
      hex: ZOTERO_HEX_PALETTE.yellow,
      name: "Yellow",
      tag: "Memorable Quote",
      rgb: [1.0, 0.83, 0.0],
    },
    {
      id: "green",
      hex: ZOTERO_HEX_PALETTE.green,
      name: "Green",
      tag: "Theme & Symbolism",
      rgb: [0.37, 0.70, 0.22],
    },
    {
      id: "red",
      hex: ZOTERO_HEX_PALETTE.red,
      name: "Red",
      tag: "Plot Climax & Conflict",
      rgb: [1.0, 0.40, 0.40],
    },
    {
      id: "blue",
      hex: ZOTERO_HEX_PALETTE.blue,
      name: "Blue",
      tag: "Character Development",
      rgb: [0.18, 0.66, 0.90],
    },
    {
      id: "purple",
      hex: ZOTERO_HEX_PALETTE.purple,
      name: "Purple",
      tag: "Historical Context",
      rgb: [0.64, 0.54, 0.90],
    },
    {
      id: "orange",
      hex: ZOTERO_HEX_PALETTE.orange,
      name: "Orange",
      tag: "Author Tone & Voice",
      rgb: [0.95, 0.60, 0.22],
    },
    {
      id: "magenta",
      hex: ZOTERO_HEX_PALETTE.magenta,
      name: "Magenta",
      tag: "Rhetoric & Style",
      rgb: [0.90, 0.43, 0.93],
    },
    {
      id: "gray",
      hex: ZOTERO_HEX_PALETTE.gray,
      name: "Gray",
      tag: "Personal Reflection",
      rgb: [0.67, 0.67, 0.67],
    },
  ],
};

/** Built-in Template: Legal & Case Analysis */
export const LEGAL_ANALYSIS_PROFILE: PdfColorSchemeProfile = {
  id: "legal-analysis",
  name: "Legal & Case Brief",
  description: "Optimized for case law, statutes, and jurisprudence briefs",
  isBuiltIn: true,
  colors: [
    {
      id: "red",
      hex: ZOTERO_HEX_PALETTE.red,
      name: "Red",
      tag: "Holding / Court Ruling",
      rgb: [1.0, 0.40, 0.40],
    },
    {
      id: "green",
      hex: ZOTERO_HEX_PALETTE.green,
      name: "Green",
      tag: "Legal Rule & Statute",
      rgb: [0.37, 0.70, 0.22],
    },
    {
      id: "yellow",
      hex: ZOTERO_HEX_PALETTE.yellow,
      name: "Yellow",
      tag: "Procedural History",
      rgb: [1.0, 0.83, 0.0],
    },
    {
      id: "blue",
      hex: ZOTERO_HEX_PALETTE.blue,
      name: "Blue",
      tag: "Material Facts",
      rgb: [0.18, 0.66, 0.90],
    },
    {
      id: "purple",
      hex: ZOTERO_HEX_PALETTE.purple,
      name: "Purple",
      tag: "Ratio Decidendi",
      rgb: [0.64, 0.54, 0.90],
    },
    {
      id: "orange",
      hex: ZOTERO_HEX_PALETTE.orange,
      name: "Orange",
      tag: "Dissenting Opinion",
      rgb: [0.95, 0.60, 0.22],
    },
    {
      id: "magenta",
      hex: ZOTERO_HEX_PALETTE.magenta,
      name: "Magenta",
      tag: "Statutory Interpretation",
      rgb: [0.90, 0.43, 0.93],
    },
    {
      id: "gray",
      hex: ZOTERO_HEX_PALETTE.gray,
      name: "Gray",
      tag: "Precedent & Citations",
      rgb: [0.67, 0.67, 0.67],
    },
  ],
};

export const BUILT_IN_PROFILES: PdfColorSchemeProfile[] = [
  ZOTERO_DEFAULT_PROFILE,
  ACADEMIC_RESEARCH_PROFILE,
  LITERATURE_REVIEW_PROFILE,
  LEGAL_ANALYSIS_PROFILE,
];

/** Normalizes a hex string to standard lowercase `#rrggbb` */
export function normalizeHexColor(hex: string): string {
  if (!hex) return "#ffd400";
  let clean = hex.trim().toLowerCase();
  if (!clean.startsWith("#")) clean = "#" + clean;
  if (clean.length === 4) {
    clean = "#" + clean[1] + clean[1] + clean[2] + clean[2] + clean[3] + clean[3];
  }
  return clean;
}

/** Calculates Euclidean distance between two RGB colors (0.0 - 1.0) */
export function colorDistance(c1: PdfRgbColor, c2: PdfRgbColor): number {
  const dr = (c1[0] ?? 0) - (c2[0] ?? 0);
  const dg = (c1[1] ?? 0) - (c2[1] ?? 0);
  const db = (c1[2] ?? 0) - (c2[2] ?? 0);
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

/**
 * Finds matching tagged color in a profile by hex string or RGB array.
 * If exact match isn't found, finds closest match within threshold.
 */
export function findTagForColor(
  scheme: PdfColorSchemeProfile | undefined | null,
  color: PdfRgbColor | string | undefined | null
): PdfTaggedColor | undefined {
  if (!scheme || !scheme.colors || scheme.colors.length === 0 || !color) {
    return undefined;
  }

  let targetHex: string;
  let targetRgb: PdfRgbColor;

  if (typeof color === "string") {
    targetHex = normalizeHexColor(color);
    targetRgb = hexToRgb(targetHex);
  } else {
    targetHex = normalizeHexColor(rgbToHex(color));
    targetRgb = color;
  }

  // Exact hex match
  for (const c of scheme.colors) {
    if (normalizeHexColor(c.hex) === targetHex) {
      return c;
    }
  }

  // Closest Euclidean distance match (threshold ~0.15)
  let bestMatch: PdfTaggedColor | undefined = undefined;
  let minDistance = 0.15;

  for (const c of scheme.colors) {
    const cRgb = c.rgb || hexToRgb(c.hex);
    const dist = colorDistance(targetRgb, cRgb);
    if (dist < minDistance) {
      minDistance = dist;
      bestMatch = c;
    }
  }

  return bestMatch;
}

/** Deep clone of a profile with optional new ID and name */
export function cloneProfile(
  profile: PdfColorSchemeProfile,
  newId?: string,
  newName?: string
): PdfColorSchemeProfile {
  return {
    id: newId || `profile_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    name: newName || (profile.name + " (Copy)"),
    description: profile.description,
    isBuiltIn: false,
    modifiedAt: Date.now(),
    colors: profile.colors.map((c) => ({
      id: c.id,
      hex: normalizeHexColor(c.hex),
      name: c.name,
      tag: c.tag,
      rgb: c.rgb ? [c.rgb[0], c.rgb[1], c.rgb[2]] : hexToRgb(c.hex),
    })),
  };
}

/** Validates and normalizes a profile object, ensuring colors array is well-formed */
export function ensureValidProfile(raw: any, fallbackId = "zotero-default"): PdfColorSchemeProfile {
  if (!raw || typeof raw !== "object") {
    return cloneProfile(ZOTERO_DEFAULT_PROFILE);
  }

  const id = typeof raw.id === "string" && raw.id.trim() ? raw.id.trim() : fallbackId;
  const name = typeof raw.name === "string" && raw.name.trim() ? raw.name.trim() : "Custom Color Scheme";
  const description = typeof raw.description === "string" ? raw.description : undefined;
  const isBuiltIn = Boolean(raw.isBuiltIn);
  const modifiedAt = typeof raw.modifiedAt === "number" ? raw.modifiedAt : Date.now();

  const colors: PdfTaggedColor[] = [];
  if (Array.isArray(raw.colors)) {
    for (const c of raw.colors) {
      if (c && typeof c === "object" && typeof c.hex === "string") {
        const hex = normalizeHexColor(c.hex);
        const name = typeof c.name === "string" && c.name.trim() ? c.name.trim() : "Color";
        const tag = typeof c.tag === "string" ? c.tag.trim() : "";
        const id = typeof c.id === "string" && c.id.trim() ? c.id.trim() : hex.replace("#", "");
        colors.push({
          id,
          hex,
          name,
          tag,
          rgb: Array.isArray(c.rgb) && c.rgb.length === 3 ? [c.rgb[0], c.rgb[1], c.rgb[2]] : hexToRgb(hex),
        });
      }
    }
  }

  // If no colors provided, populate with Zotero defaults
  if (colors.length === 0) {
    return cloneProfile(ZOTERO_DEFAULT_PROFILE, id, name);
  }

  return {
    id,
    name,
    description,
    isBuiltIn,
    modifiedAt,
    colors,
  };
}
