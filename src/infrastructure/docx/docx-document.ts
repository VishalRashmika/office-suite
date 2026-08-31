import JSZip from "jszip";
import { NS } from "../../core/utils/xml";
import { DocxStyleInfo, DocxPageSetup, parseDocxPageSetup } from "../../domain/docx";
import { IDocxParseContext } from "../../domain/docx/services/ooxml-parser";
import { IDocxSerializeContext } from "../../domain/docx/services/ooxml-serializer";

/**
 * Wraps a .docx file's zip contents. Keeps the *entire* document.xml DOM tree
 * around and only mutates the <w:body> subtree on save, so headers, footers,
 * sectPr, styles.xml, numbering.xml, media, theme, etc. all pass through
 * untouched even though our editor only understands a subset of Word features.
 */
export class DocxDocument implements IDocxParseContext, IDocxSerializeContext {
  zip: JSZip;
  documentXmlDoc: Document;
  relsXmlDoc: Document | null;
  stylesXmlDoc: Document | null = null;
  numberingXmlDoc: Document | null = null;
  mediaCache = new Map<string, string>(); // rId -> data URL, populated lazily
  private originalSectPr: Element | null = null;
  private styleMap = new Map<string, DocxStyleInfo>();

  private constructor(
    zip: JSZip,
    documentXmlDoc: Document,
    relsXmlDoc: Document | null,
    stylesXmlDoc: Document | null,
    numberingXmlDoc: Document | null
  ) {
    this.zip = zip;
    this.documentXmlDoc = documentXmlDoc;
    this.relsXmlDoc = relsXmlDoc;
    this.stylesXmlDoc = stylesXmlDoc;
    this.numberingXmlDoc = numberingXmlDoc;
    this.originalSectPr = this.findSectPr();
    this.parseStyles();
  }

  static async load(data: ArrayBuffer): Promise<DocxDocument> {
    const zip = await JSZip.loadAsync(data);
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
      throw new Error("Not a valid .docx file: word/document.xml is missing.");
    }
    const docXmlText = await docXmlFile.async("text");
    const parser = new DOMParser();
    const documentXmlDoc = parser.parseFromString(docXmlText, "application/xml");
    if (documentXmlDoc.getElementsByTagName("parsererror").length) {
      throw new Error("Failed to parse word/document.xml (file may be corrupt).");
    }

    let relsXmlDoc: Document | null = null;
    const relsFile = zip.file("word/_rels/document.xml.rels");
    if (relsFile) {
      const relsText = await relsFile.async("text");
      relsXmlDoc = parser.parseFromString(relsText, "application/xml");
    }

    let stylesXmlDoc: Document | null = null;
    const stylesFile = zip.file("word/styles.xml");
    if (stylesFile) {
      try {
        const stylesText = await stylesFile.async("text");
        stylesXmlDoc = parser.parseFromString(stylesText, "application/xml");
      } catch {
        // styles.xml optional
      }
    }

    let numberingXmlDoc: Document | null = null;
    const numberingFile = zip.file("word/numbering.xml");
    if (numberingFile) {
      try {
        const numText = await numberingFile.async("text");
        numberingXmlDoc = parser.parseFromString(numText, "application/xml");
      } catch {
        // numbering.xml optional
      }
    }

    return new DocxDocument(zip, documentXmlDoc, relsXmlDoc, stylesXmlDoc, numberingXmlDoc);
  }

  private parseStyles(): void {
    if (!this.stylesXmlDoc) return;
    const styles = this.stylesXmlDoc.getElementsByTagNameNS(NS.w, "style");
    for (let i = 0; i < styles.length; i++) {
      const styleEl = styles[i];
      const styleId = styleEl.getAttributeNS(NS.w, "styleId") || styleEl.getAttribute("w:styleId");
      if (!styleId) continue;

      const info: DocxStyleInfo = {};
      const nameEl = styleEl.getElementsByTagNameNS(NS.w, "name")[0];
      if (nameEl) {
        info.name = nameEl.getAttributeNS(NS.w, "val") || nameEl.getAttribute("w:val") || undefined;
      }

      const rPr = styleEl.getElementsByTagNameNS(NS.w, "rPr")[0];
      if (rPr) {
        const colorEl = rPr.getElementsByTagNameNS(NS.w, "color")[0];
        const colorVal = colorEl?.getAttributeNS(NS.w, "val") || colorEl?.getAttribute("w:val");
        if (colorVal && colorVal !== "auto") {
          info.color = `#${colorVal}`;
        }

        const rFonts = rPr.getElementsByTagNameNS(NS.w, "rFonts")[0];
        const font = rFonts?.getAttributeNS(NS.w, "ascii") || rFonts?.getAttribute("w:ascii");
        if (font) info.fontFamily = font;

        const szEl = rPr.getElementsByTagNameNS(NS.w, "sz")[0];
        const szVal = szEl?.getAttributeNS(NS.w, "val") || szEl?.getAttribute("w:val");
        if (szVal) info.fontSizePt = parseInt(szVal, 10) / 2;

        const bEl = rPr.getElementsByTagNameNS(NS.w, "b")[0];
        if (bEl && bEl.getAttributeNS(NS.w, "val") !== "0" && bEl.getAttribute("w:val") !== "0") {
          info.bold = true;
        }

        const iEl = rPr.getElementsByTagNameNS(NS.w, "i")[0];
        if (iEl && iEl.getAttributeNS(NS.w, "val") !== "0" && iEl.getAttribute("w:val") !== "0") {
          info.italic = true;
        }
      }

      this.styleMap.set(styleId, info);
    }
  }

  getStyle(styleId: string): DocxStyleInfo | null {
    return this.styleMap.get(styleId) ?? null;
  }

  get body(): Element {
    const body = this.documentXmlDoc.getElementsByTagNameNS(NS.w, "body")[0];
    if (!body) throw new Error("word/document.xml has no <w:body>.");
    return body;
  }

  /** Resolve a relationship id (e.g. an image rId) to its zip target path. */
  resolveRelTarget(rId: string): string | null {
    if (!this.relsXmlDoc) return null;
    const rels = this.relsXmlDoc.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute("Id") === rId) {
        const target = rels[i].getAttribute("Target");
        if (!target) return null;
        return target.startsWith("/") ? target.slice(1) : `word/${target}`;
      }
    }
    return null;
  }

  /** Load an embedded media file (image) as a data URL, cached. */
  async getMediaDataUrl(rId: string): Promise<string | null> {
    if (this.mediaCache.has(rId)) return this.mediaCache.get(rId)!;
    const path = this.resolveRelTarget(rId);
    if (!path) return null;
    const file = this.zip.file(path);
    if (!file) return null;
    const base64 = await file.async("base64");
    const ext = path.split(".").pop()?.toLowerCase() ?? "png";
    const mime =
      {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        bmp: "image/bmp",
        svg: "image/svg+xml",
      }[ext] ?? "application/octet-stream";
    const dataUrl = `data:${mime};base64,${base64}`;
    this.mediaCache.set(rId, dataUrl);
    return dataUrl;
  }

  /** Add a new media file (e.g. image) into the zip archive and register it in relationships. */
  async addMediaFile(data: ArrayBuffer | Uint8Array, ext: string): Promise<{ rId: string; dataUrl: string }> {
    if (!this.relsXmlDoc) {
      const parser = new DOMParser();
      this.relsXmlDoc = parser.parseFromString(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
        "application/xml"
      );
    }

    const rels = this.relsXmlDoc.getElementsByTagName("Relationship");
    let maxId = 0;
    for (let i = 0; i < rels.length; i++) {
      const idStr = rels[i].getAttribute("Id");
      if (idStr && idStr.startsWith("rId")) {
        const num = parseInt(idStr.slice(3), 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    }
    const rId = `rId${maxId + 1}`;
    const cleanExt = (ext || "png").toLowerCase().replace(".", "");
    const filename = `image_${Date.now()}_${maxId + 1}.${cleanExt}`;
    const zipPath = `word/media/${filename}`;

    this.zip.file(zipPath, data);

    const relEl = this.relsXmlDoc.createElementNS(null, "Relationship");
    relEl.setAttribute("Id", rId);
    relEl.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    relEl.setAttribute("Target", `media/${filename}`);
    this.relsXmlDoc.documentElement.appendChild(relEl);

    const mime =
      {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        bmp: "image/bmp",
        svg: "image/svg+xml",
        webp: "image/webp",
      }[cleanExt] ?? "image/png";

    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    let binary = "";
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    const base64 = btoa(binary);
    const dataUrl = `data:${mime};base64,${base64}`;
    this.mediaCache.set(rId, dataUrl);

    return { rId, dataUrl };
  }

  /** Determine if a given numId is a bullet list or ordered (decimal) list */
  getListFormat(numId: string): "bullet" | "ordered" {
    if (!this.numberingXmlDoc) return numId === "2" ? "ordered" : "bullet";
    const nums = this.numberingXmlDoc.getElementsByTagNameNS(NS.w, "num");
    for (let i = 0; i < nums.length; i++) {
      const id = nums[i].getAttributeNS(NS.w, "numId") || nums[i].getAttribute("w:numId");
      if (id === numId) {
        const absEl =
          nums[i].getElementsByTagNameNS(NS.w, "abstractNumId")[0] ||
          nums[i].querySelector("abstractNumId");
        const absId = absEl?.getAttributeNS(NS.w, "val") || absEl?.getAttribute("w:val");
        if (absId) {
          const absNums = this.numberingXmlDoc.getElementsByTagNameNS(NS.w, "abstractNum");
          for (let j = 0; j < absNums.length; j++) {
            const currentAbsId =
              absNums[j].getAttributeNS(NS.w, "abstractNumId") ||
              absNums[j].getAttribute("w:abstractNumId");
            if (currentAbsId === absId) {
              const lvl = absNums[j].getElementsByTagNameNS(NS.w, "lvl")[0] || absNums[j].querySelector("lvl");
              const numFmt = lvl?.getElementsByTagNameNS(NS.w, "numFmt")[0] || lvl?.querySelector("numFmt");
              const fmtVal = numFmt?.getAttributeNS(NS.w, "val") || numFmt?.getAttribute("w:val");
              if (
                fmtVal === "decimal" ||
                fmtVal === "lowerLetter" ||
                fmtVal === "upperLetter" ||
                fmtVal === "lowerRoman" ||
                fmtVal === "upperRoman"
              ) {
                return "ordered";
              }
              return "bullet";
            }
          }
        }
      }
    }
    return numId === "2" ? "ordered" : "bullet";
  }

  /** Return the appropriate numId for bullet ("1") or ordered ("2") lists, ensuring numbering.xml exists */
  getOrCreateNumId(kind: "bullet" | "ordered"): string {
    this.ensureNumberingXml();
    return kind === "bullet" ? "1" : "2";
  }

  private ensureNumberingXml(): void {
    if (!this.numberingXmlDoc) {
      const defaultNumbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val=""/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="2"/>
  </w:num>
</w:numbering>`;

      const parser = new DOMParser();
      this.numberingXmlDoc = parser.parseFromString(defaultNumbering, "application/xml");
    }

    // Ensure relationship for numbering.xml in document.xml.rels
    if (this.relsXmlDoc) {
      let hasRel = false;
      const rels = this.relsXmlDoc.getElementsByTagName("Relationship");
      for (let i = 0; i < rels.length; i++) {
        const target = rels[i].getAttribute("Target");
        if (target === "numbering.xml" || target === "word/numbering.xml") {
          hasRel = true;
          break;
        }
      }
      if (!hasRel) {
        let maxId = 0;
        for (let i = 0; i < rels.length; i++) {
          const idStr = rels[i].getAttribute("Id");
          if (idStr && idStr.startsWith("rId")) {
            const num = parseInt(idStr.slice(3), 10);
            if (!isNaN(num) && num > maxId) maxId = num;
          }
        }
        const relEl = this.relsXmlDoc.createElementNS(null, "Relationship");
        relEl.setAttribute("Id", `rId${maxId + 1}`);
        relEl.setAttribute(
          "Type",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
        );
        relEl.setAttribute("Target", "numbering.xml");
        this.relsXmlDoc.documentElement.appendChild(relEl);
      }
    }
  }

  /** Replace <w:body>'s children with newly-built ones and serialize the whole docx back to bytes. */
  async save(newBodyChildren: Element[]): Promise<ArrayBuffer> {
    const body = this.body;
    while (body.firstChild) body.removeChild(body.firstChild);
    for (const child of newBodyChildren) body.appendChild(child);

    const serializer = new XMLSerializer();
    const xmlString = serializer.serializeToString(this.documentXmlDoc);
    this.zip.file("word/document.xml", xmlString);

    if (this.numberingXmlDoc) {
      const numString = serializer.serializeToString(this.numberingXmlDoc);
      this.zip.file("word/numbering.xml", numString);
    }

    if (this.relsXmlDoc) {
      const relsString = serializer.serializeToString(this.relsXmlDoc);
      this.zip.file("word/_rels/document.xml.rels", relsString);
    }

    return this.zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  }

  private findSectPr(): Element | null {
    try {
      const body = this.body;
      for (let i = body.children.length - 1; i >= 0; i--) {
        if (body.children[i].localName === "sectPr" && body.children[i].namespaceURI === NS.w) {
          return body.children[i].cloneNode(true) as Element;
        }
      }
    } catch {
      // Body may not be available yet
    }
    return null;
  }

  /** The trailing <w:sectPr> (page setup) that must be preserved as the final body child. */
  getSectPr(): Element | null {
    if (this.originalSectPr) {
      return this.originalSectPr.cloneNode(true) as Element;
    }
    return this.findSectPr();
  }

  /** Extract page size and margin parameters matching Microsoft Word's layout. */
  getPageSetup(): DocxPageSetup {
    return parseDocxPageSetup(this.getSectPr());
  }

  createEl(localName: string): Element {
    return this.documentXmlDoc.createElementNS(NS.w, `w:${localName}`);
  }

  setAttr(el: Element, localName: string, value: string): void {
    el.setAttributeNS(NS.w, `w:${localName}`, value);
  }

  setXmlAttr(el: Element, localName: string, value: string): void {
    el.setAttributeNS(NS.xml, `xml:${localName}`, value);
  }
}
