import { CellStyle } from "../models/cell";
import { escapeXml } from "../../../core/utils/xml";

const THEME_PALETTE: Record<number, string> = {
  0: "#ffffff",
  1: "#000000",
  2: "#e7e6e6",
  3: "#44546a",
  4: "#4472c4",
  5: "#ed7d31",
  6: "#a5a5a5",
  7: "#ffc000",
  8: "#5b9bd5",
  9: "#70ad47",
};

export interface ParsedFont {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  xmlSnippet?: string;
}

export interface ParsedFill {
  patternType: string;
  background?: string;
  xmlSnippet?: string;
}

export interface ParsedBorder {
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  xmlSnippet?: string;
}

export interface ParsedXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  xmlSnippet?: string;
}

function applyTint(hex: string, tint: number): string {
  let r = parseInt(hex.substring(1, 3), 16);
  let g = parseInt(hex.substring(3, 5), 16);
  let b = parseInt(hex.substring(5, 7), 16);
  if (tint > 0) {
    r = Math.round(r + (255 - r) * tint);
    g = Math.round(g + (255 - g) * tint);
    b = Math.round(b + (255 - b) * tint);
  } else if (tint < 0) {
    r = Math.round(r * (1 + tint));
    g = Math.round(g * (1 + tint));
    b = Math.round(b * (1 + tint));
  }
  const toHex = (n: number) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

export function toArgb(hexColor: string): string {
  let clean = hexColor.replace("#", "").trim();
  if (clean.length === 3) {
    clean = clean.split("").map((c) => c + c).join("");
  }
  if (clean.length === 6) {
    return "FF" + clean.toUpperCase();
  }
  if (clean.length === 8) {
    return clean.toUpperCase();
  }
  return "FF000000";
}

function parseColorElement(colorEl: Element | null): string | undefined {
  if (!colorEl) return undefined;
  const rgb = colorEl.getAttribute("rgb");
  if (rgb) {
    if (rgb.length === 8) {
      return "#" + rgb.substring(2);
    }
    return "#" + rgb;
  }
  const themeAttr = colorEl.getAttribute("theme");
  if (themeAttr !== null) {
    const theme = parseInt(themeAttr, 10);
    const baseColor = THEME_PALETTE[theme] || "#000000";
    const tintAttr = colorEl.getAttribute("tint");
    if (tintAttr) {
      const tint = parseFloat(tintAttr);
      if (!isNaN(tint)) {
        return applyTint(baseColor, tint);
      }
    }
    return baseColor;
  }
  const indexed = colorEl.getAttribute("indexed");
  if (indexed !== null) {
    const idx = parseInt(indexed, 10);
    if (idx === 64) return "#000000";
    if (idx === 65) return "#ffffff";
  }
  return undefined;
}

export class XlsxStyleService {
  private fonts: ParsedFont[] = [];
  private fills: ParsedFill[] = [];
  private borders: ParsedBorder[] = [];
  private cellStyleXfs: ParsedXf[] = [];
  private cellXfs: ParsedXf[] = [];
  private numFmts: Map<number, string> = new Map();
  private otherSections: string[] = [];

  constructor(stylesXml?: string) {
    if (stylesXml) {
      this.parse(stylesXml);
    } else {
      this.initDefaults();
    }
  }

  private initDefaults(): void {
    this.fonts = [
      {
        fontSize: 11,
        fontFamily: "Calibri",
        xmlSnippet: '<font><sz val="11"/><name val="Calibri"/></font>',
      },
    ];
    this.fills = [
      {
        patternType: "none",
        xmlSnippet: '<fill><patternFill patternType="none"/></fill>',
      },
      {
        patternType: "gray125",
        xmlSnippet: '<fill><patternFill patternType="gray125"/></fill>',
      },
    ];
    this.borders = [
      {
        xmlSnippet: "<border><left/><right/><top/><bottom/><diagonal/></border>",
      },
    ];
    this.cellStyleXfs = [
      {
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        xmlSnippet: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
      },
    ];
    this.cellXfs = [
      {
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        xmlSnippet: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      },
    ];
  }

  public parse(stylesXml: string): void {
    const parser =
      typeof DOMParser !== "undefined"
        ? new DOMParser()
        : new (globalThis as any).DOMParser();
    const doc = parser.parseFromString(stylesXml, "application/xml");

    // 1. Custom Number Formats
    const numFmtNodes = doc.getElementsByTagName("numFmt");
    for (let i = 0; i < numFmtNodes.length; i++) {
      const el = numFmtNodes[i];
      const id = parseInt(el.getAttribute("numFmtId") || "-1", 10);
      const code = el.getAttribute("formatCode") || "";
      if (id >= 0 && code) {
        this.numFmts.set(id, code);
      }
    }

    // 2. Fonts
    const fontNodes = doc.getElementsByTagName("font");
    for (let i = 0; i < fontNodes.length; i++) {
      const el = fontNodes[i];
      const font: ParsedFont = {};

      const bEl = el.getElementsByTagName("b")[0];
      if (bEl && bEl.getAttribute("val") !== "0" && bEl.getAttribute("val") !== "false") {
        font.bold = true;
      }
      const iEl = el.getElementsByTagName("i")[0];
      if (iEl && iEl.getAttribute("val") !== "0" && iEl.getAttribute("val") !== "false") {
        font.italic = true;
      }
      const uEl = el.getElementsByTagName("u")[0];
      if (uEl && uEl.getAttribute("val") !== "none") {
        font.underline = true;
      }
      const strikeEl = el.getElementsByTagName("strike")[0];
      if (strikeEl && strikeEl.getAttribute("val") !== "0") {
        font.strike = true;
      }

      const szEl = el.getElementsByTagName("sz")[0];
      if (szEl && szEl.getAttribute("val")) {
        font.fontSize = parseFloat(szEl.getAttribute("val")!);
      }

      const nameEl = el.getElementsByTagName("name")[0];
      if (nameEl && nameEl.getAttribute("val")) {
        font.fontFamily = nameEl.getAttribute("val")!;
      }

      const colorEl = el.getElementsByTagName("color")[0];
      font.color = parseColorElement(colorEl);

      const serializer =
        typeof XMLSerializer !== "undefined"
          ? new XMLSerializer()
          : new (globalThis as any).XMLSerializer();
      font.xmlSnippet = serializer.serializeToString(el);
      this.fonts.push(font);
    }

    // 3. Fills
    const fillNodes = doc.getElementsByTagName("fill");
    for (let i = 0; i < fillNodes.length; i++) {
      const el = fillNodes[i];
      const patternFill = el.getElementsByTagName("patternFill")[0];
      const fill: ParsedFill = {
        patternType: patternFill ? patternFill.getAttribute("patternType") || "none" : "none",
      };

      if (patternFill) {
        const fgColor = patternFill.getElementsByTagName("fgColor")[0];
        const bgColor = patternFill.getElementsByTagName("bgColor")[0];
        fill.background = parseColorElement(fgColor) || parseColorElement(bgColor);
      }

      const serializer =
        typeof XMLSerializer !== "undefined"
          ? new XMLSerializer()
          : new (globalThis as any).XMLSerializer();
      fill.xmlSnippet = serializer.serializeToString(el);
      this.fills.push(fill);
    }

    // 4. Borders
    const borderNodes = doc.getElementsByTagName("border");
    for (let i = 0; i < borderNodes.length; i++) {
      const el = borderNodes[i];
      const border: ParsedBorder = {};

      const parseSide = (tag: string): string | undefined => {
        const sideEl = el.getElementsByTagName(tag)[0];
        if (!sideEl) return undefined;
        const style = sideEl.getAttribute("style");
        if (!style || style === "none") return undefined;
        const color = parseColorElement(sideEl.getElementsByTagName("color")[0]) || "#000000";
        let widthPx = "1px";
        let borderStyle = "solid";
        if (style === "medium") widthPx = "2px";
        else if (style === "thick") widthPx = "3px";
        else if (style === "double") {
          widthPx = "3px";
          borderStyle = "double";
        } else if (style === "dashed") borderStyle = "dashed";
        else if (style === "dotted") borderStyle = "dotted";
        return `${widthPx} ${borderStyle} ${color}`;
      };

      border.borderTop = parseSide("top");
      border.borderBottom = parseSide("bottom");
      border.borderLeft = parseSide("left");
      border.borderRight = parseSide("right");

      const serializer =
        typeof XMLSerializer !== "undefined"
          ? new XMLSerializer()
          : new (globalThis as any).XMLSerializer();
      border.xmlSnippet = serializer.serializeToString(el);
      this.borders.push(border);
    }

    // 5. cellXfs
    const cellXfsEl = doc.getElementsByTagName("cellXfs")[0];
    if (cellXfsEl) {
      const xfNodes = cellXfsEl.getElementsByTagName("xf");
      for (let i = 0; i < xfNodes.length; i++) {
        const xfEl = xfNodes[i];
        const xf: ParsedXf = {
          numFmtId: parseInt(xfEl.getAttribute("numFmtId") || "0", 10),
          fontId: parseInt(xfEl.getAttribute("fontId") || "0", 10),
          fillId: parseInt(xfEl.getAttribute("fillId") || "0", 10),
          borderId: parseInt(xfEl.getAttribute("borderId") || "0", 10),
          applyFont: xfEl.getAttribute("applyFont") === "1",
          applyFill: xfEl.getAttribute("applyFill") === "1",
          applyBorder: xfEl.getAttribute("applyBorder") === "1",
          applyAlignment: xfEl.getAttribute("applyAlignment") === "1",
        };

        const alignEl = xfEl.getElementsByTagName("alignment")[0];
        if (alignEl) {
          const horiz = alignEl.getAttribute("horizontal");
          if (horiz === "left" || horiz === "center" || horiz === "right") {
            xf.align = horiz;
          }
          const vert = alignEl.getAttribute("vertical");
          if (vert === "top" || vert === "bottom") {
            xf.verticalAlign = vert;
          } else if (vert === "center") {
            xf.verticalAlign = "middle";
          }
          if (alignEl.getAttribute("wrapText") === "1" || alignEl.getAttribute("wrapText") === "true") {
            xf.wrap = true;
          }
        }

        const serializer =
          typeof XMLSerializer !== "undefined"
            ? new XMLSerializer()
            : new (globalThis as any).XMLSerializer();
        xf.xmlSnippet = serializer.serializeToString(xfEl);
        this.cellXfs.push(xf);
      }
    }

    // Preserve any remaining nodes like cellStyles, dxfs, tableStyles
    for (const tag of ["cellStyles", "dxfs", "tableStyles", "colors"]) {
      const found = doc.getElementsByTagName(tag)[0];
      if (found) {
        const serializer =
          typeof XMLSerializer !== "undefined"
            ? new XMLSerializer()
            : new (globalThis as any).XMLSerializer();
        this.otherSections.push(serializer.serializeToString(found));
      }
    }

    if (this.fonts.length === 0 || this.cellXfs.length === 0) {
      this.initDefaults();
    }
  }

  public getStyle(styleIndex?: number): CellStyle {
    if (styleIndex === undefined || styleIndex < 0 || styleIndex >= this.cellXfs.length) {
      return {};
    }
    const xf = this.cellXfs[styleIndex];
    const style: CellStyle = {};

    const font = this.fonts[xf.fontId];
    if (font) {
      if (font.bold) style.bold = true;
      if (font.italic) style.italic = true;
      if (font.underline) style.underline = true;
      if (font.strike) style.strike = true;
      if (font.fontSize) style.fontSize = font.fontSize;
      if (font.fontFamily) style.fontFamily = font.fontFamily;
      if (font.color) style.color = font.color;
    }

    const fill = this.fills[xf.fillId];
    if (fill && fill.background && fill.patternType !== "none") {
      style.background = fill.background;
    }

    const border = this.borders[xf.borderId];
    if (border) {
      if (border.borderTop) style.borderTop = border.borderTop;
      if (border.borderBottom) style.borderBottom = border.borderBottom;
      if (border.borderLeft) style.borderLeft = border.borderLeft;
      if (border.borderRight) style.borderRight = border.borderRight;
    }

    if (xf.align) style.align = xf.align;
    if (xf.verticalAlign) style.verticalAlign = xf.verticalAlign;
    if (xf.wrap) style.wrap = xf.wrap;

    // Number format mapping
    style.numberFormat = this.resolveNumberFormat(xf.numFmtId);

    return style;
  }

  private resolveNumberFormat(numFmtId: number): string | undefined {
    if (numFmtId === 0) return undefined;
    if (
      (numFmtId >= 5 && numFmtId <= 8) ||
      numFmtId === 44 ||
      numFmtId === 42 ||
      numFmtId === 43
    ) {
      return "currency";
    }
    if (numFmtId === 9 || numFmtId === 10) {
      return "percent";
    }
    if (
      (numFmtId >= 14 && numFmtId <= 22) ||
      numFmtId === 45 ||
      numFmtId === 46 ||
      numFmtId === 47
    ) {
      return "date";
    }
    if (numFmtId >= 1 && numFmtId <= 4) {
      return "number";
    }

    const custom = this.numFmts.get(numFmtId);
    if (custom) {
      if (custom.includes("$") || custom.includes("€") || custom.includes("£") || custom.includes("¥")) {
        return "currency";
      }
      if (custom.includes("%")) {
        return "percent";
      }
      if (custom.includes("yy") || custom.includes("mm") || custom.includes("dd")) {
        return "date";
      }
      if (custom.includes("0") || custom.includes("#")) {
        return "number";
      }
    }
    return undefined;
  }

  public getOrCreateStyleIndex(style?: CellStyle, originalIndex?: number): number {
    if (!style || Object.keys(style).length === 0) {
      return originalIndex !== undefined && originalIndex >= 0 ? originalIndex : 0;
    }

    // If originalIndex is provided, verify whether style has diverged
    if (originalIndex !== undefined && originalIndex >= 0 && originalIndex < this.cellXfs.length) {
      const origStyle = this.getStyle(originalIndex);
      if (this.stylesMatch(origStyle, style)) {
        return originalIndex;
      }
    }

    // Check if an existing xf already matches the target style
    for (let i = 0; i < this.cellXfs.length; i++) {
      if (this.stylesMatch(this.getStyle(i), style)) {
        return i;
      }
    }

    // Need to create or reuse font, fill, border
    const fontId = this.getOrCreateFont(style);
    const fillId = this.getOrCreateFill(style);
    const borderId = this.getOrCreateBorder(style);
    const numFmtId = this.resolveNumFmtId(style.numberFormat);

    const newXf: ParsedXf = {
      numFmtId,
      fontId,
      fillId,
      borderId,
      applyFont: fontId > 0,
      applyFill: fillId > 0,
      applyBorder: borderId > 0,
      applyAlignment: !!(style.align || style.verticalAlign || style.wrap),
      align: style.align,
      verticalAlign: style.verticalAlign,
      wrap: style.wrap,
    };

    const newIdx = this.cellXfs.length;
    this.cellXfs.push(newXf);
    return newIdx;
  }

  private stylesMatch(a: CellStyle, b: CellStyle): boolean {
    return (
      (a.bold ?? false) === (b.bold ?? false) &&
      (a.italic ?? false) === (b.italic ?? false) &&
      (a.underline ?? false) === (b.underline ?? false) &&
      (a.strike ?? false) === (b.strike ?? false) &&
      (a.fontSize || 11) === (b.fontSize || 11) &&
      (a.fontFamily || "Calibri").toLowerCase() === (b.fontFamily || "Calibri").toLowerCase() &&
      (a.color || "").toLowerCase() === (b.color || "").toLowerCase() &&
      (a.background || "").toLowerCase() === (b.background || "").toLowerCase() &&
      (a.align || "left") === (b.align || "left") &&
      (a.verticalAlign || "bottom") === (b.verticalAlign || "bottom") &&
      (a.wrap ?? false) === (b.wrap ?? false) &&
      (a.numberFormat || "") === (b.numberFormat || "") &&
      (a.borderTop || "") === (b.borderTop || "") &&
      (a.borderBottom || "") === (b.borderBottom || "") &&
      (a.borderLeft || "") === (b.borderLeft || "") &&
      (a.borderRight || "") === (b.borderRight || "")
    );
  }

  private getOrCreateFont(style: CellStyle): number {
    for (let i = 0; i < this.fonts.length; i++) {
      const f = this.fonts[i];
      if (
        (f.bold ?? false) === (style.bold ?? false) &&
        (f.italic ?? false) === (style.italic ?? false) &&
        (f.underline ?? false) === (style.underline ?? false) &&
        (f.strike ?? false) === (style.strike ?? false) &&
        (f.fontSize || 11) === (style.fontSize || 11) &&
        (f.fontFamily || "Calibri").toLowerCase() === (style.fontFamily || "Calibri").toLowerCase() &&
        (f.color || "").toLowerCase() === (style.color || "").toLowerCase()
      ) {
        return i;
      }
    }

    const newFont: ParsedFont = {
      bold: style.bold,
      italic: style.italic,
      underline: style.underline,
      strike: style.strike,
      fontSize: style.fontSize || 11,
      fontFamily: style.fontFamily || "Calibri",
      color: style.color,
    };

    let xml = "<font>";
    if (newFont.bold) xml += "<b/>";
    if (newFont.italic) xml += "<i/>";
    if (newFont.underline) xml += "<u/>";
    if (newFont.strike) xml += "<strike/>";
    xml += `<sz val="${newFont.fontSize}"/>`;
    if (newFont.color) xml += `<color rgb="${toArgb(newFont.color)}"/>`;
    xml += `<name val="${escapeXml(newFont.fontFamily || "Calibri")}"/>`;
    xml += "</font>";
    newFont.xmlSnippet = xml;

    const idx = this.fonts.length;
    this.fonts.push(newFont);
    return idx;
  }

  private getOrCreateFill(style: CellStyle): number {
    if (!style.background) return 0; // none

    for (let i = 0; i < this.fills.length; i++) {
      const fl = this.fills[i];
      if (fl.background && fl.background.toLowerCase() === style.background.toLowerCase()) {
        return i;
      }
    }

    const argb = toArgb(style.background);
    const xml = `<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`;
    const newFill: ParsedFill = {
      patternType: "solid",
      background: style.background,
      xmlSnippet: xml,
    };
    const idx = this.fills.length;
    this.fills.push(newFill);
    return idx;
  }

  private getOrCreateBorder(style: CellStyle): number {
    if (!style.borderTop && !style.borderBottom && !style.borderLeft && !style.borderRight) {
      return 0;
    }

    for (let i = 0; i < this.borders.length; i++) {
      const b = this.borders[i];
      if (
        (b.borderTop || "") === (style.borderTop || "") &&
        (b.borderBottom || "") === (style.borderBottom || "") &&
        (b.borderLeft || "") === (style.borderLeft || "") &&
        (b.borderRight || "") === (style.borderRight || "")
      ) {
        return i;
      }
    }

    const serializeSide = (tag: string, cssVal?: string): string => {
      if (!cssVal || cssVal === "none") return `<${tag}/>`;
      let styleAttr = "thin";
      let color = "#000000";
      const parts = cssVal.split(/\s+/);
      for (const part of parts) {
        if (part.startsWith("#")) color = part;
        else if (part === "double") styleAttr = "double";
        else if (part === "dashed") styleAttr = "dashed";
        else if (part === "dotted") styleAttr = "dotted";
        else if (part.endsWith("px")) {
          const px = parseInt(part, 10);
          if (px >= 3) styleAttr = "thick";
          else if (px === 2) styleAttr = "medium";
        }
      }
      return `<${tag} style="${styleAttr}"><color rgb="${toArgb(color)}"/></${tag}>`;
    };

    let xml = "<border>";
    xml += serializeSide("left", style.borderLeft);
    xml += serializeSide("right", style.borderRight);
    xml += serializeSide("top", style.borderTop);
    xml += serializeSide("bottom", style.borderBottom);
    xml += "<diagonal/>";
    xml += "</border>";

    const newBorder: ParsedBorder = {
      borderTop: style.borderTop,
      borderBottom: style.borderBottom,
      borderLeft: style.borderLeft,
      borderRight: style.borderRight,
      xmlSnippet: xml,
    };
    const idx = this.borders.length;
    this.borders.push(newBorder);
    return idx;
  }

  private resolveNumFmtId(format?: string): number {
    if (!format) return 0;
    switch (format) {
      case "currency":
        return 7;
      case "percent":
        return 10;
      case "date":
        return 14;
      case "number":
        return 2;
      default:
        return 0;
    }
  }

  public toXml(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n';

    // 1. numFmts
    if (this.numFmts.size > 0) {
      xml += `  <numFmts count="${this.numFmts.size}">\n`;
      this.numFmts.forEach((code, id) => {
        xml += `    <numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>\n`;
      });
      xml += "  </numFmts>\n";
    }

    // 2. fonts
    xml += `  <fonts count="${this.fonts.length}">\n`;
    for (const f of this.fonts) {
      if (f.xmlSnippet) {
        xml += `    ${f.xmlSnippet}\n`;
      } else {
        let fXml = "<font>";
        if (f.bold) fXml += "<b/>";
        if (f.italic) fXml += "<i/>";
        if (f.underline) fXml += "<u/>";
        if (f.strike) fXml += "<strike/>";
        fXml += `<sz val="${f.fontSize || 11}"/>`;
        if (f.color) fXml += `<color rgb="${toArgb(f.color)}"/>`;
        fXml += `<name val="${escapeXml(f.fontFamily || "Calibri")}"/>`;
        fXml += "</font>";
        xml += `    ${fXml}\n`;
      }
    }
    xml += "  </fonts>\n";

    // 3. fills
    xml += `  <fills count="${this.fills.length}">\n`;
    for (const fl of this.fills) {
      if (fl.xmlSnippet) {
        xml += `    ${fl.xmlSnippet}\n`;
      } else {
        const bg = fl.background ? `<fgColor rgb="${toArgb(fl.background)}"/><bgColor indexed="64"/>` : "";
        xml += `    <fill><patternFill patternType="${fl.patternType}">${bg}</patternFill></fill>\n`;
      }
    }
    xml += "  </fills>\n";

    // 4. borders
    xml += `  <borders count="${this.borders.length}">\n`;
    for (const b of this.borders) {
      if (b.xmlSnippet) {
        xml += `    ${b.xmlSnippet}\n`;
      } else {
        xml += "    <border><left/><right/><top/><bottom/><diagonal/></border>\n";
      }
    }
    xml += "  </borders>\n";

    // 5. cellStyleXfs
    xml += `  <cellStyleXfs count="${this.cellStyleXfs.length}">\n`;
    for (const x of this.cellStyleXfs) {
      xml += `    ${x.xmlSnippet || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'}\n`;
    }
    xml += "  </cellStyleXfs>\n";

    // 6. cellXfs
    xml += `  <cellXfs count="${this.cellXfs.length}">\n`;
    for (const x of this.cellXfs) {
      if (x.xmlSnippet) {
        xml += `    ${x.xmlSnippet}\n`;
      } else {
        let alignXml = "";
        if (x.align || x.verticalAlign || x.wrap) {
          const hAttr = x.align ? ` horizontal="${x.align}"` : "";
          const vAttr = x.verticalAlign ? ` vertical="${x.verticalAlign === "middle" ? "center" : x.verticalAlign}"` : "";
          const wAttr = x.wrap ? ` wrapText="1"` : "";
          alignXml = `<alignment${hAttr}${vAttr}${wAttr}/>`;
        }
        const appF = x.applyFont ? ' applyFont="1"' : "";
        const appFl = x.applyFill ? ' applyFill="1"' : "";
        const appB = x.applyBorder ? ' applyBorder="1"' : "";
        const appA = x.applyAlignment ? ' applyAlignment="1"' : "";
        xml += `    <xf numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"${appF}${appFl}${appB}${appA}>${alignXml}</xf>\n`;
      }
    }
    xml += "  </cellXfs>\n";

    // 7. other preserved sections
    for (const other of this.otherSections) {
      xml += `  ${other}\n`;
    }

    xml += "</styleSheet>";
    return xml;
  }
}
