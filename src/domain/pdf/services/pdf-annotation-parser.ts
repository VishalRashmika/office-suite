/**
 * Parses existing ISO 32000-1 annotations from PDF documents (via pdf-lib and/or PDF.js)
 * into PdfAnnotation domain models so they can be viewed, listed, and edited natively.
 */

import {
  PDFDocument,
  PDFPage,
  PDFDict,
  PDFArray,
  PDFName,
  PDFString,
  PDFHexString,
  PDFNumber,
  PDFRef,
  PDFContext,
} from "pdf-lib";
import {
  PdfAnnotation,
  PdfNoteAnnotation,
  PdfHighlightAnnotation,
  PdfUnderlineAnnotation,
  PdfInkAnnotation,
  PdfFreeTextAnnotation,
  PdfShapeAnnotation,
  PdfRect,
  PdfRgbColor,
} from "../models/pdf-annotation";

export class PdfAnnotationParser {
  /**
   * Parses all annotations from a pdf-lib PDFPage.
   */
  static parsePageAnnotations(pdfDoc: PDFDocument, page: PDFPage, pageIndex: number): PdfAnnotation[] {
    const annotations: PdfAnnotation[] = [];
    const context = pdfDoc.context;

    // 1. Try page.node.Annots() first
    let annots = page.node.Annots();

    // 2. Fallback: page.node.get("Annots") with reference resolution
    if (!annots) {
      const rawAnnots = page.node.get(PDFName.of("Annots"));
      if (rawAnnots) {
        let lookedUp: any = rawAnnots;
        while (lookedUp instanceof PDFRef && context) {
          lookedUp = context.lookup(lookedUp);
        }
        if (lookedUp instanceof PDFArray) {
          annots = lookedUp;
        }
      }
    }

    if (!annots) return annotations;

    const annotCount = annots.size();
    for (let i = 0; i < annotCount; i++) {
      try {
        let item: any = annots.get(i);
        while (item instanceof PDFRef && context) {
          item = context.lookup(item);
        }
        if (!(item instanceof PDFDict)) continue;

        const parsed = this.parseAnnotationDict(item, pageIndex, context);
        if (parsed) {
          annotations.push(parsed);
        }
      } catch (err) {
        console.warn(`Failed to parse PDF annotation ${i} on page ${pageIndex + 1}:`, err);
      }
    }

    return annotations;
  }

  /**
   * Resolves a key in a dictionary, following any chain of indirect references (PDFRef).
   */
  private static lookupItem(dict: PDFDict, key: string, context?: PDFContext): any {
    const raw = dict.get(PDFName.of(key));
    if (!raw) return undefined;
    const ctx = context || dict.context;
    let current: any = raw;
    while (current instanceof PDFRef && ctx) {
      current = ctx.lookup(current);
    }
    return current;
  }

  /**
   * Parses a single PDFDict annotation.
   */
  static parseAnnotationDict(dict: PDFDict, pageIndex: number, context?: PDFContext): PdfAnnotation | null {
    const subtypeItem = this.lookupItem(dict, "Subtype", context);
    if (!(subtypeItem instanceof PDFName)) return null;

    const subtype = subtypeItem.asString().replace(/^\//, "");

    // Skip non-markup system annotations
    if (subtype === "Popup" || subtype === "Link" || subtype === "Widget") {
      return null;
    }

    // Common keys
    const id = this.extractString(dict, "NM", context) || `annot_${pageIndex}_${Math.random().toString(36).slice(2, 9)}`;
    const rect = this.extractRect(dict, context);
    if (!rect) return null;

    const color = this.extractColor(dict, context) || [1, 0.85, 0];
    let contents = this.extractString(dict, "Contents", context) || "";

    // If contents is empty, check associated /Popup dictionary (common in Adobe Acrobat & Apple Preview)
    if (!contents) {
      const popupItem = this.lookupItem(dict, "Popup", context);
      if (popupItem instanceof PDFDict) {
        contents = this.extractString(popupItem, "Contents", context) || "";
      }
    }

    const author = this.extractString(dict, "T", context) || "Author";
    const linkedNote = this.extractString(dict, "ObsidianNote", context);

    switch (subtype) {
      case "Text":
      case "Caret":
      case "Stamp": {
        const iconName = (this.extractName(dict, "Name", context) as any) || "Comment";
        const note: PdfNoteAnnotation = {
          id,
          type: "text",
          pageIndex,
          rect,
          color,
          contents,
          author,
          iconName,
          linkedNote,
        };
        return note;
      }

      case "Highlight":
      case "StrikeOut":
      case "Squiggly": {
        const quadPoints = this.extractNumberArray(dict, "QuadPoints", context);
        const highlight: PdfHighlightAnnotation = {
          id,
          type: "highlight",
          pageIndex,
          rect,
          color,
          contents,
          highlightedText: contents,
          quadPoints: quadPoints.length > 0 ? quadPoints : [
            rect[0], rect[3], rect[2], rect[3],
            rect[0], rect[1], rect[2], rect[1]
          ],
          author,
          linkedNote,
        };
        return highlight;
      }

      case "Underline": {
        const quadPoints = this.extractNumberArray(dict, "QuadPoints", context);
        const underline: PdfUnderlineAnnotation = {
          id,
          type: "underline",
          pageIndex,
          rect,
          color,
          contents,
          highlightedText: contents,
          quadPoints: quadPoints.length > 0 ? quadPoints : [
            rect[0], rect[3], rect[2], rect[3],
            rect[0], rect[1], rect[2], rect[1]
          ],
          author,
          linkedNote,
        };
        return underline;
      }

      case "Ink":
      case "PolyLine": {
        const paths = this.extractInkList(dict, context);
        const strokeWidth = this.extractBorderWidth(dict, context) || 2;
        const ink: PdfInkAnnotation = {
          id,
          type: "ink",
          pageIndex,
          rect,
          color,
          contents,
          paths,
          strokeWidth,
          author,
          linkedNote,
        };
        return ink;
      }

      case "FreeText": {
        const da = this.extractString(dict, "DA", context) || "";
        const sizeMatch = da.match(/(\d+(?:\.\d+)?)\s+Tf/) || da.match(/Tf\s+(\d+(?:\.\d+)?)/);
        const fontSize = sizeMatch ? parseFloat(sizeMatch[1]) : 12;

        let textColor = color;
        const colorMatch = da.match(/([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+rg/);
        if (colorMatch) {
          textColor = [
            parseFloat(colorMatch[1]),
            parseFloat(colorMatch[2]),
            parseFloat(colorMatch[3]),
          ];
        }

        const isBold = /(Bold|Black)/i.test(da);
        const isItalic = /(Italic|Oblique)/i.test(da);

        const qVal = this.extractNumber(dict, "Q", context);
        const textAlign = qVal === 1 ? "center" : qVal === 2 ? "right" : "left";

        const fillColor = this.extractColorByKey(dict, "IC", context);
        const borderWidth = this.extractBorderWidth(dict, context);
        let borderColor: PdfRgbColor | undefined;
        let borderStyle: "none" | "solid" | "dashed" | undefined;
        if (borderWidth && borderWidth > 0) {
          borderColor = color;
          borderStyle = "solid";
        }

        const freeText: PdfFreeTextAnnotation = {
          id,
          type: "freetext",
          pageIndex,
          rect,
          color,
          contents,
          fontSize,
          textColor,
          isBold: isBold || undefined,
          isItalic: isItalic || undefined,
          textAlign,
          fillColor,
          borderColor,
          borderWidth,
          borderStyle,
          author,
          linkedNote,
        };
        return freeText;
      }

      case "Square":
      case "Circle":
      case "Line":
      case "Polygon": {
        const strokeWidth = this.extractBorderWidth(dict, context) || 2;
        const fillColor = this.extractColorByKey(dict, "IC", context);
        const shapeType = (subtype === "Circle" ? "circle" : subtype === "Line" ? "line" : "square") as "square" | "circle" | "line";
        let linePoints: [number, number, number, number] | undefined;
        if (subtype === "Line") {
          const lArr = this.extractNumberArray(dict, "L", context);
          if (lArr.length >= 4) {
            linePoints = [lArr[0], lArr[1], lArr[2], lArr[3]];
          }
        }
        const shape: PdfShapeAnnotation = {
          id,
          type: shapeType,
          pageIndex,
          rect,
          color,
          strokeColor: color,
          fillColor,
          linePoints,
          strokeWidth,
          contents,
          author,
          linkedNote,
        };
        return shape;
      }

      default:
        return null;
    }
  }

  private static extractNumber(dict: PDFDict, key: string, context?: PDFContext): number | undefined {
    const val = this.lookupItem(dict, key, context);
    if (val instanceof PDFNumber) return val.asNumber();
    return undefined;
  }

  private static extractString(dict: PDFDict, key: string, context?: PDFContext): string | undefined {
    const val = this.lookupItem(dict, key, context);
    if (!val) return undefined;
    if (val instanceof PDFString || val instanceof PDFHexString) {
      if (typeof (val as any).decodeText === "function") {
        return (val as any).decodeText();
      }
      return val.asString();
    }
    return undefined;
  }

  private static extractName(dict: PDFDict, key: string, context?: PDFContext): string | undefined {
    const val = this.lookupItem(dict, key, context);
    if (val instanceof PDFName) {
      return val.asString().replace(/^\//, "");
    }
    return undefined;
  }

  private static extractRect(dict: PDFDict, context?: PDFContext): PdfRect | null {
    const rectArr = this.lookupItem(dict, "Rect", context);
    if (!(rectArr instanceof PDFArray) || rectArr.size() < 4) return null;
    const ctx = context || dict.context;
    const nums: number[] = [];
    for (let i = 0; i < 4; i++) {
      let item: any = rectArr.get(i);
      while (item instanceof PDFRef && ctx) {
        item = ctx.lookup(item);
      }
      nums.push(item instanceof PDFNumber ? item.asNumber() : 0);
    }
    return [nums[0], nums[1], nums[2], nums[3]];
  }

  private static extractColor(dict: PDFDict, context?: PDFContext): PdfRgbColor | undefined {
    return this.extractColorByKey(dict, "C", context);
  }

  private static extractColorByKey(dict: PDFDict, key: string, context?: PDFContext): PdfRgbColor | undefined {
    const cArr = this.lookupItem(dict, key, context);
    if (!(cArr instanceof PDFArray) || cArr.size() < 3) return undefined;
    const ctx = context || dict.context;
    const nums: number[] = [];
    for (let i = 0; i < 3; i++) {
      let item: any = cArr.get(i);
      while (item instanceof PDFRef && ctx) {
        item = ctx.lookup(item);
      }
      nums.push(item instanceof PDFNumber ? item.asNumber() : 0);
    }
    return [nums[0], nums[1], nums[2]];
  }

  private static extractNumberArray(dict: PDFDict, key: string, context?: PDFContext): number[] {
    const arr = this.lookupItem(dict, key, context);
    if (!(arr instanceof PDFArray)) return [];
    const ctx = context || dict.context;
    const res: number[] = [];
    const size = arr.size();
    for (let i = 0; i < size; i++) {
      let item: any = arr.get(i);
      while (item instanceof PDFRef && ctx) {
        item = ctx.lookup(item);
      }
      if (item instanceof PDFNumber) res.push(item.asNumber());
    }
    return res;
  }

  private static extractInkList(dict: PDFDict, context?: PDFContext): number[][][] {
    const inkList = this.lookupItem(dict, "InkList", context);
    if (!(inkList instanceof PDFArray)) return [];
    const ctx = context || dict.context;
    const paths: number[][][] = [];
    const count = inkList.size();
    for (let i = 0; i < count; i++) {
      let strokeArr: any = inkList.get(i);
      while (strokeArr instanceof PDFRef && ctx) {
        strokeArr = ctx.lookup(strokeArr);
      }
      if (!(strokeArr instanceof PDFArray)) continue;
      const strokeCoords: number[][] = [];
      const coordCount = strokeArr.size();
      for (let j = 0; j < coordCount; j += 2) {
        let xItem: any = strokeArr.get(j);
        let yItem: any = strokeArr.get(j + 1);
        while (xItem instanceof PDFRef && ctx) xItem = ctx.lookup(xItem);
        while (yItem instanceof PDFRef && ctx) yItem = ctx.lookup(yItem);
        const x = xItem instanceof PDFNumber ? xItem.asNumber() : 0;
        const y = yItem instanceof PDFNumber ? yItem.asNumber() : 0;
        strokeCoords.push([x, y]);
      }
      if (strokeCoords.length > 0) {
        paths.push(strokeCoords);
      }
    }
    return paths;
  }

  private static extractBorderWidth(dict: PDFDict, context?: PDFContext): number | undefined {
    const bs = this.lookupItem(dict, "BS", context);
    if (bs instanceof PDFDict) {
      const w = this.lookupItem(bs, "W", context);
      if (w instanceof PDFNumber) return w.asNumber();
    }
    return undefined;
  }
}
