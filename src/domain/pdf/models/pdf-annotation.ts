/**
 * Domain models for PDF annotations and edit operations.
 * Designed to map 1:1 with ISO 32000-1 (PDF Standard) dictionaries for 100% cross-compatibility
 * with native PDF viewers (Adobe Acrobat, Apple Preview, Google Chrome, Foxit, PDF-XChange).
 */

export type PdfAnnotationType =
  | "text" // Sticky Note (/Subtype /Text)
  | "highlight" // Text Highlight (/Subtype /Highlight)
  | "ink" // Freehand drawing (/Subtype /Ink)
  | "freetext" // Text box directly on PDF (/Subtype /FreeText)
  | "square" // Rectangle shape (/Subtype /Square)
  | "circle" // Circle/oval shape (/Subtype /Circle)
  | "line"; // Line shape (/Subtype /Line)

/**
 * Standard PDF User Space coordinates [llx, lly, urx, ury]:
 * - llx: lower-left X
 * - lly: lower-left Y
 * - urx: upper-right X
 * - ury: upper-right Y
 * Origin (0,0) is at the lower-left corner of the page in PDF coordinate space.
 */
export type PdfRect = [number, number, number, number];

/** Normalized RGB color [r, g, b] where each value is between 0.0 and 1.0 */
export type PdfRgbColor = [number, number, number];

export interface BasePdfAnnotation {
  id: string;
  type: PdfAnnotationType;
  pageIndex: number; // 0-based page index
  rect: PdfRect;
  color: PdfRgbColor;
  opacity?: number;
  contents?: string; // Text content, note, or comment
  author?: string; // /T in PDF standard
  createdAt?: Date; // /CreationDate
  modifiedAt?: Date; // /M in PDF standard
  linkedNote?: string; // Vault path to linked Obsidian markdown note (e.g. "Notes/Summary.md")
  highlightedText?: string; // Text snippet if attached to a selection
}

/** Sticky Note annotation (/Subtype /Text) */
export interface PdfNoteAnnotation extends BasePdfAnnotation {
  type: "text";
  iconName?: "Comment" | "Key" | "Note" | "Help" | "NewParagraph" | "Paragraph" | "Insert";
  isOpen?: boolean;
}

/**
 * Highlight markup annotation (/Subtype /Highlight).
 * quadPoints contains 8 coordinates per line: [x1, y1, x2, y2, x3, y3, x4, y4]
 * (top-left, top-right, bottom-left, bottom-right in PDF user space)
 */
export interface PdfHighlightAnnotation extends BasePdfAnnotation {
  type: "highlight";
  quadPoints: number[]; // Flat array of (8 * numLines) coordinates
  highlightedText: string;
}

/** Freehand Ink drawing annotation (/Subtype /Ink) */
export interface PdfInkAnnotation extends BasePdfAnnotation {
  type: "ink";
  paths: number[][][]; // Array of paths, each path is array of [x, y] coordinates in PDF user space
  strokeWidth: number;
}

/** FreeText box annotation (/Subtype /FreeText) */
export interface PdfFreeTextAnnotation extends BasePdfAnnotation {
  type: "freetext";
  fontSize: number;
  fontFamily?: string;
  textColor: PdfRgbColor;
  borderColor?: PdfRgbColor;
  fillColor?: PdfRgbColor;
}

/** Vector geometric shape annotation (/Subtype /Square, /Circle, /Line) */
export interface PdfShapeAnnotation extends BasePdfAnnotation {
  type: "square" | "circle" | "line";
  strokeWidth: number;
  strokeColor: PdfRgbColor;
  fillColor?: PdfRgbColor;
  linePoints?: [number, number, number, number]; // [x1, y1, x2, y2] for lines
}

export type PdfAnnotation =
  | PdfNoteAnnotation
  | PdfHighlightAnnotation
  | PdfInkAnnotation
  | PdfFreeTextAnnotation
  | PdfShapeAnnotation;

/** Preset colors matching modern PDF viewers & highlight palettes */
export const PDF_COLOR_PRESETS: { name: string; hex: string; rgb: PdfRgbColor }[] = [
  { name: "Yellow", hex: "#fde047", rgb: [0.99, 0.88, 0.28] },
  { name: "Green", hex: "#86efac", rgb: [0.52, 0.94, 0.67] },
  { name: "Cyan", hex: "#67e8f9", rgb: [0.40, 0.91, 0.98] },
  { name: "Pink", hex: "#f472b6", rgb: [0.96, 0.45, 0.71] },
  { name: "Orange", hex: "#fb923c", rgb: [0.98, 0.57, 0.24] },
  { name: "Purple", hex: "#c084fc", rgb: [0.75, 0.52, 0.99] },
  { name: "Red", hex: "#f87171", rgb: [0.97, 0.44, 0.44] },
  { name: "Blue", hex: "#60a5fa", rgb: [0.38, 0.65, 0.98] },
];

/** Utility to convert hex color to [r, g, b] (0.0 - 1.0) */
export function hexToRgb(hex: string): PdfRgbColor {
  const cleanHex = hex.replace("#", "");
  const num = parseInt(cleanHex, 16);
  if (cleanHex.length === 3) {
    const r = ((num >> 8) & 0xf) * 17;
    const g = ((num >> 4) & 0xf) * 17;
    const b = (num & 0xf) * 17;
    return [r / 255, g / 255, b / 255];
  }
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return [r / 255, g / 255, b / 255];
}

/** Utility to convert [r, g, b] (0.0 - 1.0) to hex string */
export function rgbToHex(rgb?: PdfRgbColor | null): string {
  if (!rgb || !Array.isArray(rgb) || rgb.length < 3) {
    return "#ffd700";
  }
  const r = Math.round(Math.max(0, Math.min(1, rgb[0] ?? 1)) * 255)
    .toString(16)
    .padStart(2, "0");
  const g = Math.round(Math.max(0, Math.min(1, rgb[1] ?? 0.85)) * 255)
    .toString(16)
    .padStart(2, "0");
  const b = Math.round(Math.max(0, Math.min(1, rgb[2] ?? 0)) * 255)
    .toString(16)
    .padStart(2, "0");
  return `#${r}${g}${b}`;
}
