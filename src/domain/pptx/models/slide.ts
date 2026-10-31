import { PptxTransition } from "./transition";

export interface PptxRun {
  text: string;
  fontSize?: number; // In points (pt), e.g. 18, 24, 32
  fontFamily?: string;
  color?: string; // Hex color, e.g. #000000
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  isSlideNumber?: boolean; // True if this run is a slide number placeholder (<a:fld type="slNum">)
  hyperlink?: {
    url: string;
    tooltip?: string;
  };
}

export interface PptxParagraph {
  align?: "left" | "center" | "right" | "justify";
  bullet?: {
    type: "char" | "number" | "none";
    char?: string;
    number?: number;
    font?: string;
    color?: string;
  };
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  indent?: number;
  runs: PptxRun[];
}

export type PptxElementType = "shape" | "text" | "image" | "table" | "group" | "line";

export interface PptxTableCell {
  text: string;
  paragraphs: PptxParagraph[];
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  rowSpan?: number;
  colSpan?: number;
  verticalAlign?: "top" | "middle" | "bottom";
}

export interface PptxTableRow {
  height: number; // in pixels
  cells: PptxTableCell[];
}

export interface PptxTable {
  columns: number[]; // column widths in pixels
  rows: PptxTableRow[];
}

export interface PptxShapeFill {
  type: "solid" | "gradient" | "none";
  color?: string;
  gradientStops?: Array<{ color: string; position: number }>;
}

export interface PptxShapeBorder {
  color: string;
  width: number; // in pixels
  style?: "solid" | "dashed" | "dotted";
}

export interface PptxElement {
  id: string;
  name: string;
  type: PptxElementType;
  x: number; // pixels relative to slide
  y: number; // pixels relative to slide
  width: number; // pixels
  height: number; // pixels
  rotation: number; // degrees
  zIndex: number;
  // Shape/Text specific
  shapeType?: string; // rect, roundRect, ellipse, triangle, line, arrow, etc.
  fill?: PptxShapeFill;
  border?: PptxShapeBorder;
  paragraphs?: PptxParagraph[];
  verticalAlign?: "top" | "middle" | "bottom";
  padding?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  // Image specific
  imageSrc?: string; // Data URL or object URL
  imageAlt?: string;
  opacity?: number;
  // Table specific
  table?: PptxTable;
  // Group specific
  children?: PptxElement[];
}

export interface PptxSlide {
  index: number; // 0-based
  slideNumber: number; // 1-based (for display e.g. 1, 2, 3...)
  id: string;
  title: string;
  background?: string; // Hex color, CSS gradient, or image URL
  elements: PptxElement[];
  transition?: PptxTransition;
  notes: string; // Speaker notes
  hidden?: boolean;
}

export interface PptxTheme {
  colors: Record<string, string>; // e.g. dk1, lt1, accent1-6
  majorFont?: string;
  minorFont?: string;
}

export interface PptxPresentation {
  title: string;
  width: number; // Viewport width in pixels, e.g. 960 (16:9 widescreen) or 960 (4:3)
  height: number; // Viewport height in pixels, e.g. 540 (16:9) or 720 (4:3)
  emuWidth: number;
  emuHeight: number;
  slides: PptxSlide[];
  theme: PptxTheme;
}
