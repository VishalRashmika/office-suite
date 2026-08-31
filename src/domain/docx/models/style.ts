import { NS } from "../../../core/utils/xml";

export interface DocxStyleInfo {
  name?: string;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
}

export const HEADING_STYLE_TO_LEVEL: Record<string, number> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Title: 1,
  Subtitle: 2,
};

export const LEVEL_TO_HEADING_STYLE: Record<number, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
  4: "Heading4",
  5: "Heading5",
  6: "Heading6",
};

export const WORD_HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGreen: "#006400",
  darkMagenta: "#8b008b",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#d3d3d3",
  black: "#000000",
  white: "#ffffff",
};

export const WORD_THEME_COLORS: Record<string, string> = {
  accent1: "#4472C4",
  accent2: "#ED7D31",
  accent3: "#A5A5A5",
  accent4: "#FFC000",
  accent5: "#5B9BD5",
  accent6: "#70AD47",
  dark1: "#000000",
  light1: "#FFFFFF",
  dark2: "#44546A",
  light2: "#E7E6E6",
  hyperlink: "#0563C1",
  followedHyperlink: "#954F72",
};

export function normalizeDocxColor(c: string | null | undefined, el?: Element | null): string | null {
  if (el) {
    const themeCol = el.getAttributeNS(NS.w, "themeColor") ?? el.getAttribute("w:themeColor");
    const themeFill = el.getAttributeNS(NS.w, "themeFill") ?? el.getAttribute("w:themeFill");
    const theme = themeCol || themeFill;
    if (theme && WORD_THEME_COLORS[theme]) return WORD_THEME_COLORS[theme];
  }

  if (!c || c === "auto" || c === "none") return null;
  const lower = c.toLowerCase();
  if (WORD_THEME_COLORS[c]) return WORD_THEME_COLORS[c];
  if (WORD_THEME_COLORS[lower]) return WORD_THEME_COLORS[lower];
  if (WORD_HIGHLIGHT_COLORS[c]) return WORD_HIGHLIGHT_COLORS[c];
  if (WORD_HIGHLIGHT_COLORS[lower]) return WORD_HIGHLIGHT_COLORS[lower];
  if (/^[0-9A-Fa-f]{6}$/.test(c)) return `#${c}`;
  if (/^[0-9A-Fa-f]{8}$/.test(c)) return `#${c.slice(2)}`; // strip alpha
  if (c.startsWith("#")) return c;
  return null;
}
