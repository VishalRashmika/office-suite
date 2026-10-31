/**
 * Serializes PdfAnnotation domain models into standard ISO 32000-1 PDF annotation dictionaries
 * in pdf-lib, generating compliant /AP appearance streams to ensure 100% cross-compatibility
 * with external readers like Adobe Acrobat, Apple Preview, Chrome, Edge, Foxit, and PDF-XChange.
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
  PDFBool,
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
  if (/[\(\)\\\x80-\uFFFF]/.test(text)) {
    return PDFHexString.fromText(text);
  }
  return PDFString.of(text);
}

/** Mapping of Unicode punctuation and symbols to WinAnsi single-byte encoding */
const WIN_ANSI_MAP: Record<number, number> = {
  8364: 0x80, // Euro €
  8218: 0x82, // Single low-9 quote ‚
  402:  0x83, // Latin small f with hook ƒ
  8222: 0x84, // Double low-9 quote „
  8230: 0x85, // Ellipsis …
  8224: 0x86, // Dagger †
  8225: 0x87, // Double dagger ‡
  710:  0x88, // Circumflex accent ˆ
  8240: 0x89, // Per mille ‰
  352:  0x8a, // S with caron Š
  8249: 0x8b, // Single left-pointing angle quote ‹
  338:  0x8c, // Latin capital ligature OE Œ
  381:  0x8e, // Z with caron Ž
  8216: 0x91, // Left single quote ‘
  8217: 0x92, // Right single quote ’
  8220: 0x93, // Left double quote “
  8221: 0x94, // Right double quote ”
  8226: 0x95, // Bullet •
  8211: 0x96, // En dash –
  8212: 0x97, // Em dash —
  732:  0x98, // Small tilde ˜
  8482: 0x99, // Trade mark ™
  353:  0x9a, // s with caron š
  8250: 0x9b, // Single right-pointing angle quote ›
  339:  0x9c, // Latin small ligature oe œ
  382:  0x9e, // z with caron ž
  376:  0x9f, // Y with diaeresis Ÿ
};

/** Encodes string to WinAnsi hex bytes for PDF text operator (<hex> Tj) */
export function encodeWinAnsiHex(str: string): string {
  let hex = "";
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    if (code >= 32 && code <= 255) {
      hex += code.toString(16).padStart(2, "0");
    } else if (WIN_ANSI_MAP[code]) {
      hex += WIN_ANSI_MAP[code].toString(16).padStart(2, "0");
    } else if (code < 32) {
      hex += "20"; // space
    } else {
      hex += "3f"; // ?
    }
  }
  return hex;
}

/** Wraps text lines to fit within a specific character column width */
export function wrapFreeText(text: string, maxCharsPerLine: number): string[] {
  const lines: string[] = [];
  const rawLines = text.split("\n");
  for (const raw of rawLines) {
    if (raw.length <= maxCharsPerLine) {
      lines.push(raw);
      continue;
    }
    const words = raw.split(" ");
    let currentLine = "";
    for (const w of words) {
      if (!currentLine) {
        currentLine = w;
      } else if ((currentLine + " " + w).length <= maxCharsPerLine) {
        currentLine += " " + w;
      } else {
        lines.push(currentLine);
        currentLine = w;
      }
    }
    if (currentLine) lines.push(currentLine);
  }
  return lines.length > 0 ? lines : [""];
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
          if (subtype === "Popup") {
            // Check if this popup is attached to a managed markup annotation
            const parentRef = annotDict.get(PDFName.of("Parent"));
            let isParentManaged = true;
            if (parentRef instanceof PDFRef) {
              const parentDict = context.lookup(parentRef);
              if (parentDict instanceof PDFDict) {
                const parentSubtypeName = parentDict.get(PDFName.of("Subtype"));
                if (parentSubtypeName instanceof PDFName) {
                  const pSub = parentSubtypeName.asString().replace(/^\//, "");
                  if (!PdfAnnotationSerializer.MANAGED_SUBTYPES.has(pSub)) {
                    isParentManaged = false;
                  }
                }
              }
            }
            if (isParentManaged) {
              continue;
            }
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
      const annotDict = this.createAnnotationDict(pdfDoc, annotation, page);
      const annotRef = context.register(annotDict);
      newAnnotsArray.push(annotRef);

      // Create an associated /Popup annotation for annotations with comments or notes
      // This allows Acrobat, Foxit, PDF-XChange, and Preview to display the comment popup natively
      if (annotation.contents || annotation.type === "text") {
        const popupDict = context.obj({
          Type: PDFName.of("Annot"),
          Subtype: PDFName.of("Popup"),
          Parent: annotRef,
          Rect: context.obj([
            PDFNumber.of(Math.round(annotation.rect[0] * 100) / 100),
            PDFNumber.of(Math.round(annotation.rect[1] * 100) / 100),
            PDFNumber.of(Math.round((annotation.rect[0] + 180) * 100) / 100),
            PDFNumber.of(Math.round((annotation.rect[1] + 90) * 100) / 100),
          ]),
          Contents: createPdfTextString(annotation.contents || ""),
          F: PDFNumber.of(28),
          P: page.ref,
          Open: annotation.type === "text" && (annotation as PdfNoteAnnotation).isOpen ? PDFBool.True : PDFBool.False,
        }) as PDFDict;

        if (annotation.author) {
          popupDict.set(PDFName.of("T"), createPdfTextString(annotation.author));
        }
        if (annotation.modifiedAt) {
          popupDict.set(PDFName.of("M"), PDFString.of(formatPdfDate(annotation.modifiedAt)));
        }

        const popupRef = context.register(popupDict);
        annotDict.set(PDFName.of("Popup"), popupRef);
        newAnnotsArray.push(popupRef);
      }
    }

    page.node.set(PDFName.of("Annots"), newAnnotsArray);
  }

  /**
   * Creates an ISO 32000-1 compliant PDF dictionary for any annotation type,
   * including Appearance Streams (/AP) for cross-reader compatibility.
   */
  static createAnnotationDict(pdfDoc: PDFDocument, annot: PdfAnnotation, page?: PDFPage): PDFDict {
    const context = pdfDoc.context;
    const dict = context.obj({}) as PDFDict;

    // Common standard ISO 32000-1 keys
    dict.set(PDFName.of("Type"), PDFName.of("Annot"));
    dict.set(PDFName.of("NM"), PDFString.of(annot.id));
    dict.set(PDFName.of("Rect"), this.createPdfRectArray(pdfDoc, annot.rect));
    dict.set(PDFName.of("F"), PDFNumber.of(annot.type === "text" ? 28 : 4)); // Print flag (28 = Print + NoZoom + NoRotate for text icons)
    dict.set(PDFName.of("M"), PDFString.of(formatPdfDate(annot.modifiedAt || new Date())));

    if (page?.ref) {
      dict.set(PDFName.of("P"), page.ref);
    }
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

    // Type-specific serialization & appearance stream creation
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

  // ==================== Type-Specific Populators ====================

  /** Sticky Note: /Subtype /Text with appearance stream */
  private static populateNoteAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfNoteAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Text"));
    dict.set(PDFName.of("Subj"), PDFString.of("Sticky Note"));
    dict.set(PDFName.of("Name"), PDFName.of(annot.iconName || "Comment"));
    if (annot.isOpen) {
      dict.set(PDFName.of("Open"), PDFBool.True);
    }

    const normRect = this.getNormalizedRect(annot.rect);
    const width = Math.max(1, normRect.width);
    const height = Math.max(1, normRect.height);

    const apRef = this.generateNoteAppearance(pdfDoc, annot, width, height);
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  /** Highlight: /Subtype /Highlight with QuadPoints, transparency, and appearance stream */
  private static populateHighlightAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfHighlightAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Highlight"));
    dict.set(PDFName.of("Subj"), PDFString.of("Highlight"));
    dict.set(PDFName.of("CA"), PDFNumber.of(0.5));
    dict.set(PDFName.of("ca"), PDFNumber.of(0.5));

    // QuadPoints array in PDF user space: [x1, y1, x2, y2, x3, y3, x4, y4, ...]
    const normRect = this.getNormalizedRect(annot.rect);
    const quadPoints =
      annot.quadPoints && annot.quadPoints.length > 0
        ? annot.quadPoints
        : [normRect[0], normRect[3], normRect[2], normRect[3], normRect[0], normRect[1], normRect[2], normRect[1]];
    const quadArray = context.obj([]) as PDFArray;
    for (const val of quadPoints) {
      quadArray.push(PDFNumber.of(Math.round(val * 100) / 100));
    }
    dict.set(PDFName.of("QuadPoints"), quadArray);

    if (annot.highlightedText && !annot.contents) {
      dict.set(PDFName.of("Contents"), createPdfTextString(annot.highlightedText));
    }

    const apRef = this.generateHighlightAppearance(
      pdfDoc,
      annot,
      normRect.width,
      normRect.height,
      normRect.llx,
      normRect.lly
    );
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  /** Underline: /Subtype /Underline with QuadPoints and appearance stream */
  private static populateUnderlineAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfUnderlineAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("Underline"));
    dict.set(PDFName.of("Subj"), PDFString.of("Underline"));

    // Border Style dictionary
    const bsDict = context.obj({
      Type: PDFName.of("Border"),
      W: PDFNumber.of(1.5),
    }) as PDFDict;
    dict.set(PDFName.of("BS"), bsDict);

    // QuadPoints array in PDF user space: [x1, y1, x2, y2, x3, y3, x4, y4, ...]
    const normRect = this.getNormalizedRect(annot.rect);
    const quadPoints =
      annot.quadPoints && annot.quadPoints.length > 0
        ? annot.quadPoints
        : [normRect[0], normRect[3], normRect[2], normRect[3], normRect[0], normRect[1], normRect[2], normRect[1]];
    const quadArray = context.obj([]) as PDFArray;
    for (const val of quadPoints) {
      quadArray.push(PDFNumber.of(Math.round(val * 100) / 100));
    }
    dict.set(PDFName.of("QuadPoints"), quadArray);

    if (annot.highlightedText && !annot.contents) {
      dict.set(PDFName.of("Contents"), createPdfTextString(annot.highlightedText));
    }

    const apRef = this.generateUnderlineAppearance(
      pdfDoc,
      annot,
      normRect.width,
      normRect.height,
      normRect.llx,
      normRect.lly
    );
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  /** Freehand Ink Drawing: /Subtype /Ink with InkList, BS, and appearance stream */
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

    const normRect = this.getNormalizedRect(annot.rect);
    const apRef = this.generateInkAppearance(
      pdfDoc,
      annot,
      normRect.width,
      normRect.height,
      normRect.llx,
      normRect.lly
    );
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  /** Text Box: /Subtype /FreeText with DA, quadding, font resource, and appearance stream */
  private static populateFreeTextAnnotation(
    pdfDoc: PDFDocument,
    dict: PDFDict,
    annot: PdfFreeTextAnnotation
  ): void {
    const context = pdfDoc.context;
    dict.set(PDFName.of("Subtype"), PDFName.of("FreeText"));
    dict.set(PDFName.of("Subj"), PDFString.of("Text Box"));
    
    // Quadding / text alignment: 0=Left, 1=Center, 2=Right
    const qVal = annot.textAlign === "center" ? 1 : annot.textAlign === "right" ? 2 : 0;
    dict.set(PDFName.of("Q"), PDFNumber.of(qVal));

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

    if (annot.fillColor) {
      dict.set(PDFName.of("IC"), this.createPdfColorArray(pdfDoc, annot.fillColor));
    }

    if (annot.borderColor && annot.borderStyle !== "none") {
      dict.set(PDFName.of("C"), this.createPdfColorArray(pdfDoc, annot.borderColor));
      const bsDict = context.obj({
        Type: PDFName.of("Border"),
        W: PDFNumber.of(annot.borderWidth || 1),
        S: PDFName.of(annot.borderStyle === "dashed" ? "D" : "S"),
      }) as PDFDict;
      dict.set(PDFName.of("BS"), bsDict);
    }

    const normRect = this.getNormalizedRect(annot.rect);
    const apRef = this.generateFreeTextAppearance(pdfDoc, annot, normRect.width, normRect.height);
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  /** Shape: /Subtype /Square | /Circle | /Line with BS, IC, and appearance stream */
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

    if (annot.type === "line") {
      const normRect = this.getNormalizedRect(annot.rect);
      const [lx1, ly1, lx2, ly2] = annot.linePoints || [
        normRect.llx,
        normRect.lly,
        normRect.urx,
        normRect.ury,
      ];
      const lineArray = context.obj([
        PDFNumber.of(Math.round(lx1 * 100) / 100),
        PDFNumber.of(Math.round(ly1 * 100) / 100),
        PDFNumber.of(Math.round(lx2 * 100) / 100),
        PDFNumber.of(Math.round(ly2 * 100) / 100),
      ]) as PDFArray;
      dict.set(PDFName.of("L"), lineArray);
    }

    const normRect = this.getNormalizedRect(annot.rect);
    const apRef = this.generateShapeAppearance(
      pdfDoc,
      annot,
      normRect.width,
      normRect.height,
      normRect.llx,
      normRect.lly
    );
    dict.set(PDFName.of("AP"), context.obj({ N: apRef }));
  }

  // ==================== Appearance Stream Generators ====================

  /** Generates compliant appearance stream for Sticky Note icon */
  private static generateNoteAppearance(
    pdfDoc: PDFDocument,
    annot: PdfNoteAnnotation,
    width: number,
    height: number
  ): PDFRef {
    const context = pdfDoc.context;
    const c = annot.color || [1, 0.85, 0];
    const r = Math.round(c[0] * 100) / 100;
    const g = Math.round(c[1] * 100) / 100;
    const b = Math.round(c[2] * 100) / 100;
    const darkR = Math.round(Math.max(0, r * 0.7) * 100) / 100;
    const darkG = Math.round(Math.max(0, g * 0.7) * 100) / 100;
    const darkB = Math.round(Math.max(0, b * 0.7) * 100) / 100;

    const sx = (width / 24).toFixed(3);
    const sy = (height / 24).toFixed(3);

    const streamOps = `q
${sx} 0 0 ${sy} 0 0 cm
${r} ${g} ${b} rg
${darkR} ${darkG} ${darkB} RG
1 w 1 j 1 J
2 2 m 2 22 l 22 22 l 22 8 l 16 2 l 2 2 l B
16 2 m 16 8 l 22 8 l B
${darkR} ${darkG} ${darkB} RG
1 w
6 17 m 18 17 l S
6 13 m 18 13 l S
6 9 m 13 9 l S
Q`;

    const stream = context.flateStream(streamOps, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Resources: context.obj({}),
    });
    return context.register(stream);
  }

  /** Generates compliant appearance stream for Highlight with blend mode Multiply & transparency */
  private static generateHighlightAppearance(
    pdfDoc: PDFDocument,
    annot: PdfHighlightAnnotation,
    width: number,
    height: number,
    llx: number,
    lly: number
  ): PDFRef {
    const context = pdfDoc.context;
    const gsRef = this.getOrCreateHighlightExtGState(pdfDoc);
    const c = annot.color || [1, 0.9, 0.2];
    const r = Math.round(c[0] * 100) / 100;
    const g = Math.round(c[1] * 100) / 100;
    const b = Math.round(c[2] * 100) / 100;

    let ops = `q\n/GS0 gs\n${r} ${g} ${b} rg\n`;
    if (annot.quadPoints && annot.quadPoints.length >= 8) {
      for (let i = 0; i < annot.quadPoints.length; i += 8) {
        const x1 = Math.round((annot.quadPoints[i] - llx) * 100) / 100;
        const y1 = Math.round((annot.quadPoints[i + 1] - lly) * 100) / 100;
        const x2 = Math.round((annot.quadPoints[i + 2] - llx) * 100) / 100;
        const y2 = Math.round((annot.quadPoints[i + 3] - lly) * 100) / 100;
        const x3 = Math.round((annot.quadPoints[i + 4] - llx) * 100) / 100;
        const y3 = Math.round((annot.quadPoints[i + 5] - lly) * 100) / 100;
        const x4 = Math.round((annot.quadPoints[i + 6] - llx) * 100) / 100;
        const y4 = Math.round((annot.quadPoints[i + 7] - lly) * 100) / 100;

        ops += `${x1} ${y1} m ${x2} ${y2} l ${x4} ${y4} l ${x3} ${y3} l h f\n`;
      }
    } else {
      ops += `0 0 ${width} ${height} re f\n`;
    }
    ops += `Q`;

    const stream = context.flateStream(ops, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Group: context.obj({
        Type: PDFName.of("Group"),
        S: PDFName.of("Transparency"),
        CS: PDFName.of("DeviceRGB"),
      }),
      Resources: context.obj({
        ExtGState: context.obj({
          GS0: gsRef,
        }),
      }),
    });
    return context.register(stream);
  }

  /** Generates compliant appearance stream for Underline */
  private static generateUnderlineAppearance(
    pdfDoc: PDFDocument,
    annot: PdfUnderlineAnnotation,
    width: number,
    height: number,
    llx: number,
    lly: number
  ): PDFRef {
    const context = pdfDoc.context;
    const c = annot.color || [0.2, 0.4, 0.9];
    const r = Math.round(c[0] * 100) / 100;
    const g = Math.round(c[1] * 100) / 100;
    const b = Math.round(c[2] * 100) / 100;

    let ops = `q\n1.5 w\n1 J\n${r} ${g} ${b} RG\n`;
    if (annot.quadPoints && annot.quadPoints.length >= 8) {
      for (let i = 0; i < annot.quadPoints.length; i += 8) {
        const x3 = Math.round((annot.quadPoints[i + 4] - llx) * 100) / 100;
        const y3 = Math.round((annot.quadPoints[i + 5] - lly) * 100) / 100;
        const x4 = Math.round((annot.quadPoints[i + 6] - llx) * 100) / 100;
        const y4 = Math.round((annot.quadPoints[i + 7] - lly) * 100) / 100;
        ops += `${x3} ${y3} m ${x4} ${y4} l S\n`;
      }
    } else {
      ops += `0 0 m ${width} 0 l S\n`;
    }
    ops += `Q`;

    const stream = context.flateStream(ops, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Resources: context.obj({}),
    });
    return context.register(stream);
  }

  /** Generates compliant appearance stream for Ink drawing */
  private static generateInkAppearance(
    pdfDoc: PDFDocument,
    annot: PdfInkAnnotation,
    width: number,
    height: number,
    llx: number,
    lly: number
  ): PDFRef {
    const context = pdfDoc.context;
    const c = annot.color || [0.9, 0.2, 0.2];
    const r = Math.round(c[0] * 100) / 100;
    const g = Math.round(c[1] * 100) / 100;
    const b = Math.round(c[2] * 100) / 100;
    const sw = annot.strokeWidth || 2;

    let ops = `q\n1 J 1 j\n${sw} w\n${r} ${g} ${b} RG\n`;
    for (const stroke of annot.paths) {
      if (stroke.length === 1) {
        const x = Math.round((stroke[0][0] - llx) * 100) / 100;
        const y = Math.round((stroke[0][1] - lly) * 100) / 100;
        ops += `${x} ${y} m ${x + 0.1} ${y} l S\n`;
      } else if (stroke.length > 1) {
        const startX = Math.round((stroke[0][0] - llx) * 100) / 100;
        const startY = Math.round((stroke[0][1] - lly) * 100) / 100;
        ops += `${startX} ${startY} m\n`;
        for (let j = 1; j < stroke.length; j++) {
          const px = Math.round((stroke[j][0] - llx) * 100) / 100;
          const py = Math.round((stroke[j][1] - lly) * 100) / 100;
          ops += `${px} ${py} l\n`;
        }
        ops += `S\n`;
      }
    }
    ops += `Q`;

    const stream = context.flateStream(ops, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Resources: context.obj({}),
    });
    return context.register(stream);
  }

  /** Generates compliant appearance stream for FreeText text box */
  private static generateFreeTextAppearance(
    pdfDoc: PDFDocument,
    annot: PdfFreeTextAnnotation,
    width: number,
    height: number
  ): PDFRef {
    const context = pdfDoc.context;
    const fontSize = annot.fontSize || 12;
    const tc = annot.textColor || [0, 0, 0];
    const tr = Math.round(tc[0] * 100) / 100;
    const tg = Math.round(tc[1] * 100) / 100;
    const tb = Math.round(tc[2] * 100) / 100;

    let fontName = "Helvetica";
    if (annot.isBold && annot.isItalic) {
      fontName = "Helvetica-BoldOblique";
    } else if (annot.isBold) {
      fontName = "Helvetica-Bold";
    } else if (annot.isItalic) {
      fontName = "Helvetica-Oblique";
    }
    const fontRef = this.getOrCreateStandardFont(pdfDoc, fontName);

    let ops = `q\n0 0 ${width} ${height} re W n\n`;

    // Fill background if specified
    if (annot.fillColor) {
      const fr = Math.round(annot.fillColor[0] * 100) / 100;
      const fg = Math.round(annot.fillColor[1] * 100) / 100;
      const fb = Math.round(annot.fillColor[2] * 100) / 100;
      ops += `${fr} ${fg} ${fb} rg\n0 0 ${width} ${height} re f\n`;
    }

    // Border if specified
    if (annot.borderColor && annot.borderStyle !== "none") {
      const br = Math.round(annot.borderColor[0] * 100) / 100;
      const bg = Math.round(annot.borderColor[1] * 100) / 100;
      const bb = Math.round(annot.borderColor[2] * 100) / 100;
      const bw = annot.borderWidth || 1;
      if (annot.borderStyle === "dashed") {
        ops += `[3 3] 0 d\n`;
      }
      ops += `${br} ${bg} ${bb} RG\n${bw} w\n0.5 0.5 ${Math.max(0.1, width - 1)} ${Math.max(0.1, height - 1)} re S\n`;
      if (annot.borderStyle === "dashed") {
        ops += `[] 0 d\n`;
      }
    }

    // Wrap and render text lines
    const textContent = annot.contents || "";
    const maxCharsPerLine = Math.max(1, Math.floor((width - 8) / (fontSize * 0.52)));
    const lines = wrapFreeText(textContent, maxCharsPerLine);
    const leading = Math.round(fontSize * 1.25 * 10) / 10;
    const startY = Math.max(4, Math.round((height - fontSize - 4) * 10) / 10);

    ops += `BT\n/Helv ${fontSize} Tf\n${tr} ${tg} ${tb} rg\n${leading} TL\n`;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      let lineX = 4;
      const approxLineWidth = line.length * fontSize * 0.52;
      if (annot.textAlign === "center") {
        lineX = Math.max(4, (width - approxLineWidth) / 2);
      } else if (annot.textAlign === "right") {
        lineX = Math.max(4, width - approxLineWidth - 4);
      }
      const yPos = startY - i * leading;
      if (yPos >= -leading) {
        ops += `1 0 0 1 ${lineX.toFixed(2)} ${yPos.toFixed(2)} Tm\n`;
        const hex = encodeWinAnsiHex(line);
        ops += `<${hex}> Tj\n`;
      }
    }
    ops += `ET\nQ`;

    const stream = context.flateStream(ops, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Resources: context.obj({
        Font: context.obj({
          Helv: fontRef,
        }),
      }),
    });
    return context.register(stream);
  }

  /** Generates compliant appearance stream for Shapes (Square, Circle, Line) */
  private static generateShapeAppearance(
    pdfDoc: PDFDocument,
    annot: PdfShapeAnnotation,
    width: number,
    height: number,
    llx: number,
    lly: number
  ): PDFRef {
    const context = pdfDoc.context;
    const sw = annot.strokeWidth || 2;
    const sc = annot.strokeColor || annot.color || [0.2, 0.4, 0.9];
    const sr = Math.round(sc[0] * 100) / 100;
    const sg = Math.round(sc[1] * 100) / 100;
    const sb = Math.round(sc[2] * 100) / 100;

    let ops = `q\n${sw} w\n${sr} ${sg} ${sb} RG\n`;
    if (annot.fillColor) {
      const fc = annot.fillColor;
      const fcr = Math.round(fc[0] * 100) / 100;
      const fcg = Math.round(fc[1] * 100) / 100;
      const fcb = Math.round(fc[2] * 100) / 100;
      ops += `${fcr} ${fcg} ${fcb} rg\n`;
    }

    if (annot.type === "line") {
      const [lx1, ly1, lx2, ly2] = annot.linePoints || [llx, lly, llx + width, lly + height];
      const x1 = Math.round((lx1 - llx) * 100) / 100;
      const y1 = Math.round((ly1 - lly) * 100) / 100;
      const x2 = Math.round((lx2 - llx) * 100) / 100;
      const y2 = Math.round((ly2 - lly) * 100) / 100;
      ops = `q\n1 J 1 j\n${sw} w\n${sr} ${sg} ${sb} RG\n${x1} ${y1} m ${x2} ${y2} l S\nQ`;
    } else if (annot.type === "circle") {
      const cx = width / 2;
      const cy = height / 2;
      const rx = Math.max(0.1, (width - sw) / 2);
      const ry = Math.max(0.1, (height - sw) / 2);
      const kappa = 0.552284749831;
      const ox = rx * kappa;
      const oy = ry * kappa;

      const pLeft = Math.round((cx - rx) * 100) / 100;
      const pRight = Math.round((cx + rx) * 100) / 100;
      const pTop = Math.round((cy + ry) * 100) / 100;
      const pBottom = Math.round((cy - ry) * 100) / 100;
      const mCy = Math.round(cy * 100) / 100;
      const mCx = Math.round(cx * 100) / 100;

      const c1x = Math.round((cx - rx) * 100) / 100;
      const c1y = Math.round((cy + oy) * 100) / 100;
      const c2x = Math.round((cx - ox) * 100) / 100;
      const c2y = Math.round((cy + ry) * 100) / 100;

      const c3x = Math.round((cx + ox) * 100) / 100;
      const c3y = Math.round((cy + ry) * 100) / 100;
      const c4x = Math.round((cx + rx) * 100) / 100;
      const c4y = Math.round((cy + oy) * 100) / 100;

      const c5x = Math.round((cx + rx) * 100) / 100;
      const c5y = Math.round((cy - oy) * 100) / 100;
      const c6x = Math.round((cx + ox) * 100) / 100;
      const c6y = Math.round((cy - ry) * 100) / 100;

      const c7x = Math.round((cx - ox) * 100) / 100;
      const c7y = Math.round((cy - ry) * 100) / 100;
      const c8x = Math.round((cx - rx) * 100) / 100;
      const c8y = Math.round((cy - oy) * 100) / 100;

      ops += `${pLeft} ${mCy} m\n`;
      ops += `${c1x} ${c1y} ${c2x} ${c2y} ${mCx} ${pTop} c\n`;
      ops += `${c3x} ${c3y} ${c4x} ${c4y} ${pRight} ${mCy} c\n`;
      ops += `${c5x} ${c5y} ${c6x} ${c6y} ${mCx} ${pBottom} c\n`;
      ops += `${c7x} ${c7y} ${c8x} ${c8y} ${pLeft} ${mCy} c\n`;
      ops += `h ${annot.fillColor ? "B" : "S"}\nQ`;
    } else {
      // Square
      const halfW = sw / 2;
      const rw = Math.max(0.1, width - sw);
      const rh = Math.max(0.1, height - sw);
      ops += `${halfW} ${halfW} ${rw} ${rh} re ${annot.fillColor ? "B" : "S"}\nQ`;
    }

    const stream = context.flateStream(ops, {
      Type: PDFName.of("XObject"),
      Subtype: PDFName.of("Form"),
      FormType: PDFNumber.of(1),
      BBox: context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(width), PDFNumber.of(height)]),
      Resources: context.obj({}),
    });
    return context.register(stream);
  }

  // ==================== Resource Caching Helpers ====================

  /** Caches standard Type 1 font in PDF document context to avoid duplicates */
  private static getOrCreateStandardFont(pdfDoc: PDFDocument, fontName: string): PDFRef {
    const context = pdfDoc.context;
    const cacheKey = `__CachedFont_${fontName}`;
    if ((pdfDoc as any)[cacheKey]) {
      return (pdfDoc as any)[cacheKey];
    }

    const fontDict = context.obj({
      Type: PDFName.of("Font"),
      Subtype: PDFName.of("Type1"),
      BaseFont: PDFName.of(fontName),
      Encoding: PDFName.of("WinAnsiEncoding"),
    });
    const fontRef = context.register(fontDict);
    (pdfDoc as any)[cacheKey] = fontRef;
    return fontRef;
  }

  /** Caches ExtGState for Highlight multiply transparency to avoid duplicates */
  private static getOrCreateHighlightExtGState(pdfDoc: PDFDocument): PDFRef {
    const context = pdfDoc.context;
    const cacheKey = `__CachedExtGState_Multiply`;
    if ((pdfDoc as any)[cacheKey]) {
      return (pdfDoc as any)[cacheKey];
    }

    const gsDict = context.obj({
      Type: PDFName.of("ExtGState"),
      BM: PDFName.of("Multiply"),
      ca: PDFNumber.of(0.45),
      CA: PDFNumber.of(0.45),
    });
    const gsRef = context.register(gsDict);
    (pdfDoc as any)[cacheKey] = gsRef;
    return gsRef;
  }

  // ==================== Geometry & Array Helpers ====================

  /** Normalizes rect coordinates so llx < urx and lly < ury */
  static getNormalizedRect(rect: PdfRect): {
    llx: number;
    lly: number;
    urx: number;
    ury: number;
    width: number;
    height: number;
  } {
    const minX = Math.min(rect[0], rect[2]);
    const maxX = Math.max(rect[0], rect[2]);
    const minY = Math.min(rect[1], rect[3]);
    const maxY = Math.max(rect[1], rect[3]);
    const width = Math.max(1, maxX - minX);
    const height = Math.max(1, maxY - minY);
    return {
      llx: Math.round(minX * 100) / 100,
      lly: Math.round(minY * 100) / 100,
      urx: Math.round((minX + width) * 100) / 100,
      ury: Math.round((minY + height) * 100) / 100,
      width: Math.round(width * 100) / 100,
      height: Math.round(height * 100) / 100,
    };
  }

  static createPdfRectArray(pdfDoc: PDFDocument, rect: PdfRect): PDFArray {
    const context = pdfDoc.context;
    const norm = this.getNormalizedRect(rect);
    return context.obj([
      PDFNumber.of(norm.llx),
      PDFNumber.of(norm.lly),
      PDFNumber.of(norm.urx),
      PDFNumber.of(norm.ury),
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
