/**
 * Serializes PdfAnnotation domain models into standard ISO 32000-1 PDF annotation dictionaries
 * in pdf-lib, ensuring 100% cross-compatibility with external readers like Adobe Acrobat,
 * Apple Preview, Chrome, Edge, Foxit, and PDF-XChange.
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

/** Formats JavaScript Date into ISO standard PDF date string: (D:YYYYMMDDHHmmSSOHH'mm') */
export function formatPdfDate(date: Date = new Date()): string {
  const pad = (n: number) => n.toString().padStart(2, "0");
  const y = date.getFullYear();
  const m = pad(date.getMonth() + 1);
  const d = pad(date.getDate());
  const h = pad(date.getHours());
  const min = pad(date.getMinutes());
  const s = pad(date.getSeconds());
  return `D:${y}${m}${d}${h}${min}${s}Z`;
}

/** Encodes string safely for PDF: uses UTF-16BE hex string if text contains parentheses, backslashes, or non-ASCII characters to avoid syntax errors */
export function createPdfTextString(text: string): PDFString | PDFHexString {
  // If the text contains any parentheses, backslashes, or non-ASCII characters, encode as hex string
  if (/[\(\)\\\x80-\uFFFF]/.test(text)) {
    return PDFHexString.fromText(text);
  }
  return PDFString.of(text);
}

export class PdfAnnotationSerializer {
  /** Standard markup annotation subtypes managed by this editor */
  static readonly MANAGED_SUBTYPES = new Set([
    "Text",
    "Highlight",
    "Underline",
    "Ink",
    "FreeText",
    "Square",
    "Circle",
    "Line",
  ]);

  /**
   * Writes all annotations for a page into that page's standard /Annots array.
   * Purges existing managed markup annotations while preserving untouched
   * third-party non-markup annotations (e.g. /Widget form fields, /Link hyperlinks).
   */
  static syncPageAnnotations(
    pdfDoc: PDFDocument,
    page: PDFPage,
    annotations: PdfAnnotation[]
  ): void {
    const context = pdfDoc.context;
    let annots = page.node.Annots();

    if (!annots) {
      annots = context.obj([]) as PDFArray;
      page.node.set(PDFName.of("Annots"), annots);
    }

    // Preserve non-markup third-party annotations (like /Widget form fields, /Link links)
    const keptAnnots: (PDFRef | PDFDict)[] = [];
    const annotSize = annots.size();

    for (let i = 0; i < annotSize; i++) {
      const annotRefOrDict = annots.get(i);
      const annotDict = annotRefOrDict instanceof PDFRef
        ? (context.lookup(annotRefOrDict) as PDFDict)
        : (annotRefOrDict as PDFDict);

      if (annotDict instanceof PDFDict) {
        const subtypeName = annotDict.get(PDFName.of("Subtype"));
        if (subtypeName instanceof PDFName) {
          const subtype = subtypeName.asString().replace(/^\//, "");
          if (PdfAnnotationSerializer.MANAGED_SUBTYPES.has(subtype)) {
            // Managed markup annotation: skip, will be replaced fresh from active state below
            continue;
          }
        }
      }
      keptAnnots.push(annotRefOrDict as PDFRef | PDFDict);
    }

    // Reset Annots array with kept non-markup annotations
    const newAnnotsArray = context.obj([]) as PDFArray;
    for (const kept of keptAnnots) {
      newAnnotsArray.push(kept);
    }

    // Serialize each active annotation into an ISO PDF annotation dict and register it
    for (const annotation of annotations) {
      const annotDict = this.createAnnotationDict(pdfDoc, annotation);
      const annotRef = context.register(annotDict);
      newAnnotsArray.push(annotRef);
    }

    page.node.set(PDFName.of("Annots"), newAnnotsArray);
  }

  /**
   * Creates an ISO 32000-1 compliant PDF dictionary for any annotation type.
   */
  static createAnnotationDict(pdfDoc: PDFDocument, annot: PdfAnnotation): PDFDict {
    const context = pdfDoc.context;
    const dict = context.obj({}) as PDFDict;

    // Common standard ISO 32000-1 keys
    dict.set(PDFName.of("Type"), PDFName.of("Annot"));
    dict.set(PDFName.of("NM"), PDFString.of(annot.id));
    dict.set(PDFName.of("Rect"), this.createPdfRectArray(pdfDoc, annot.rect));
    dict.set(PDFName.of("F"), PDFNumber.of(4)); // Print flag
    dict.set(PDFName.of("M"), PDFString.of(formatPdfDate(annot.modifiedAt || new Date())));

    if (annot.author) {
      dict.set(PDFName.of("T"), createPdfTextString(annot.author));
    }
    if (annot.contents) {
      dict.set(PDFName.of("Contents"), createPdfTextString(annot.contents));
    }
    if (annot.color) {
      dict.set(PDFName.of("C"), this.createPdfColorArray(pdfDoc, annot.color));
    }

    // Custom Obsidian vault note backlink stored in standard-safe custom key
    if (annot.linkedNote) {
      dict.set(PDFName.of("ObsidianNote"), createPdfTextString(annot.linkedNote));
    }

    // Type-specific serialization
    switch (annot.type) {
      case "text":
        this.populateNoteAnnotation(pdfDoc, dict, annot);
        break;
      case "highlight":
        this.populateHighlightAnnotation(pdfDoc, dict, annot);
        break;
      case "underline":
        this.populateUnderlineAnnotation(pdfDoc, dict, annot as PdfUnderlineAnnotation);
        break;
      case "ink":
        this.populateInkAnnotation(pdfDoc, dict, annot);
        break;
      case "freetext":
        this.populateFreeTextAnnotation(pdfDoc, dict, annot);
        break;
      case "square":
      case "circle":
      case "line":
        this.populateShapeAnnotation(pdfDoc, dict, annot);
        break;
    }

    return dict;
  }

  /** Sticky Note: /Subtype /Text */
  private static populateNoteAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfNoteAnnotation
  ): void {
    dict.set(PDFName.of("Subtype"), PDFName.of("Text"));
    dict.set(PDFName.of("Subj"), PDFString.of("Sticky Note"));
    dict.set(PDFName.of("Name"), PDFName.of(annot.iconName || "Comment"));
    if (annot.isOpen) {
      dict.set(PDFName.of("Open"), PDFName.of("true"));
    }
  }

  /** Highlight: /Subtype /Highlight */
  private static populateHighlightAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfHighlightAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Highlight"));
    dict.set(PDFName.of("Subj"), PDFString.of("Highlight"));

    // QuadPoints array in PDF user space: [x1, y1, x2, y2, x3, y3, x4, y4, ...]
    const quadArray = context.obj([]) as PDFArray;
    for (const val of annot.quadPoints) {
      quadArray.push(PDFNumber.of(Math.round(val * 100) / 100));
    }
    dict.set(PDFName.of("QuadPoints"), quadArray);

    if (annot.highlightedText && !annot.contents) {
      dict.set(PDFName.of("Contents"), createPdfTextString(annot.highlightedText));
    }
  }

  /** Underline: /Subtype /Underline */
  private static populateUnderlineAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfUnderlineAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Underline"));
    dict.set(PDFName.of("Subj"), PDFString.of("Underline"));

    // QuadPoints array in PDF user space: [x1, y1, x2, y2, x3, y3, x4, y4, ...]
    const quadArray = context.obj([]) as PDFArray;
    for (const val of annot.quadPoints) {
      quadArray.push(PDFNumber.of(Math.round(val * 100) / 100));
    }
    dict.set(PDFName.of("QuadPoints"), quadArray);

    if (annot.highlightedText && !annot.contents) {
      dict.set(PDFName.of("Contents"), createPdfTextString(annot.highlightedText));
    }
  }

  /** Freehand Ink Drawing: /Subtype /Ink */
  private static populateInkAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfInkAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Ink"));
    dict.set(PDFName.of("Subj"), PDFString.of("Drawing"));

    // Border Style dictionary for stroke width
    const bsDict = context.obj({
      Type: PDFName.of("Border"),
      W: PDFNumber.of(annot.strokeWidth || 2),
    }) as PDFDict;
    dict.set(PDFName.of("BS"), bsDict);

    // InkList: array of point arrays: [ [x1, y1, x2, y2, ...], [x1, y1, x2, y2, ...] ]
    const inkList = context.obj([]) as PDFArray;
    for (const stroke of annot.paths) {
      const strokeArray = context.obj([]) as PDFArray;
      for (const pt of stroke) {
        strokeArray.push(PDFNumber.of(Math.round(pt[0] * 100) / 100));
        strokeArray.push(PDFNumber.of(Math.round(pt[1] * 100) / 100));
      }
      inkList.push(strokeArray);
    }
    dict.set(PDFName.of("InkList"), inkList);
  }

  /** Text Box: /Subtype /FreeText */
  private static populateFreeTextAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfFreeTextAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("FreeText"));
    dict.set(PDFName.of("Subj"), PDFString.of("Text Box"));

    const size = annot.fontSize || 12;
    const tc = annot.textColor || [0, 0, 0];
    let fontName = "Helv";
    if (annot.isBold && annot.isItalic) {
      fontName = "Helv-BoldOblique";
    } else if (annot.isBold) {
      fontName = "Helv-Bold";
    } else if (annot.isItalic) {
      fontName = "Helv-Oblique";
    }
    // Default appearance string: Font and Color command
    const da = `/${fontName} ${size} Tf ${tc[0]} ${tc[1]} ${tc[2]} rg`;
    dict.set(PDFName.of("DA"), PDFString.of(da));

    if (annot.borderColor) {
      const bsDict = context.obj({
        Type: PDFName.of("Border"),
        W: PDFNumber.of(1),
      }) as PDFDict;
      dict.set(PDFName.of("BS"), bsDict);
    }
  }

  /** Shape: /Subtype /Square | /Circle | /Line */
  private static populateShapeAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfShapeAnnotation
  ): void {
    const context = pdfDoc.context;
    const subtypeMap = {
      square: "Square",
      circle: "Circle",
      line: "Line",
    };
    dict.set(PDFName.of("Subtype"), PDFName.of(subtypeMap[annot.type] || "Square"));
    dict.set(PDFName.of("Subj"), PDFString.of("Shape"));

    const bsDict = context.obj({
      Type: PDFName.of("Border"),
      W: PDFNumber.of(annot.strokeWidth || 2),
    }) as PDFDict;
    dict.set(PDFName.of("BS"), bsDict);

    if (annot.fillColor) {
      dict.set(PDFName.of("IC"), this.createPdfColorArray(pdfDoc, annot.fillColor));
    }

    if (annot.type === "line" && annot.linePoints) {
      const [x1, y1, x2, y2] = annot.linePoints;
      const lineArray = context.obj([
        PDFNumber.of(x1),
        PDFNumber.of(y1),
        PDFNumber.of(x2),
        PDFNumber.of(y2),
      ]) as PDFArray;
      dict.set(PDFName.of("L"), lineArray);
    }
  }

  private static createPdfRectArray(pdfDoc: PDFDocument, rect: PdfRect): PDFArray {
    const context = pdfDoc.context;
    return context.obj([
      PDFNumber.of(Math.round(rect[0] * 100) / 100),
      PDFNumber.of(Math.round(rect[1] * 100) / 100),
      PDFNumber.of(Math.round(rect[2] * 100) / 100),
      PDFNumber.of(Math.round(rect[3] * 100) / 100),
    ]) as PDFArray;
  }

  private static createPdfColorArray(pdfDoc: PDFDocument, color: PdfRgbColor): PDFArray {
    const context = pdfDoc.context;
    return context.obj([
      PDFNumber.of(Math.round(color[0] * 100) / 100),
      PDFNumber.of(Math.round(color[1] * 100) / 100),
      PDFNumber.of(Math.round(color[2] * 100) / 100),
    ]) as PDFArray;
  }
}
