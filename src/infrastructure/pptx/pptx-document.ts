import JSZip from "jszip";
import {
  PptxPresentation,
  PptxSlide,
  PptxElement,
  PptxParagraph,
  PptxRun,
  PptxTable,
  PptxTableRow,
  PptxTableCell,
  PptxTheme,
  PptxTransition,
  TransitionType,
  TransitionDirection,
  DEFAULT_TRANSITION,
  PptxExport,
} from "../../domain/pptx";

export interface PptxSlideEntry {
  rId: string;
  slideId: number;
  path: string;       // e.g. "ppt/slides/slide1.xml"
  relsPath: string;   // e.g. "ppt/slides/_rels/slide1.xml.rels"
  notesPath?: string; // e.g. "ppt/notesSlides/notesSlide1.xml"
  notesRId?: string;
}

export class PptxDocument {
  private zip: JSZip;
  private presentation: PptxPresentation;
  private slideEntries: PptxSlideEntry[] = [];
  private presDoc: Document;
  private presRelsDoc: Document;
  private contentTypesDoc: Document;
  private slideXmlDocs: Map<string, Document> = new Map();
  private slideRelsDocs: Map<string, Document> = new Map();
  private notesXmlDocs: Map<string, Document> = new Map();
  private mediaDataUrls: Map<string, string> = new Map();
  private deletedFiles: Set<string> = new Set();
  public dirty = false;

  private constructor(
    zip: JSZip,
    presentation: PptxPresentation,
    slideEntries: PptxSlideEntry[],
    presDoc: Document,
    presRelsDoc: Document,
    contentTypesDoc: Document
  ) {
    this.zip = zip;
    this.presentation = presentation;
    this.slideEntries = slideEntries;
    this.presDoc = presDoc;
    this.presRelsDoc = presRelsDoc;
    this.contentTypesDoc = contentTypesDoc;
  }

  static async load(bytes: ArrayBuffer): Promise<PptxDocument> {
    const zip = await JSZip.loadAsync(bytes);
    const parser = getXmlParser();

    // 1. Parse [Content_Types].xml
    const ctFile = zip.file("[Content_Types].xml");
    if (!ctFile) {
      throw new Error("Invalid .pptx file: [Content_Types].xml is missing.");
    }
    const ctXml = await ctFile.async("text");
    const contentTypesDoc = parser.parseFromString(ctXml, "application/xml");

    // 2. Parse ppt/presentation.xml
    const presFile = zip.file("ppt/presentation.xml");
    if (!presFile) {
      throw new Error("Invalid .pptx file: ppt/presentation.xml is missing.");
    }
    const presXml = await presFile.async("text");
    const presDoc = parser.parseFromString(presXml, "application/xml");

    // Slide Dimensions
    let width = 960;
    let height = 540;
    let emuWidth = 9144000;
    let emuHeight = 5143500;

    const sldSz = presDoc.getElementsByTagNameNS("*", "sldSz")[0];
    if (sldSz) {
      const cx = Number(sldSz.getAttribute("cx"));
      const cy = Number(sldSz.getAttribute("cy"));
      if (cx > 0 && cy > 0) {
        emuWidth = cx;
        emuHeight = cy;
        width = Math.round(cx / 9525);
        height = Math.round(cy / 9525);
      }
    }

    // 3. Parse Presentation Relationships (ppt/_rels/presentation.xml.rels)
    const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    let presRelsDoc: Document;
    const presRelMap: Record<string, string> = {};

    if (presRelsFile) {
      const presRelsXml = await presRelsFile.async("text");
      presRelsDoc = parser.parseFromString(presRelsXml, "application/xml");
      const rels = presRelsDoc.getElementsByTagNameNS("*", "Relationship");
      for (let i = 0; i < rels.length; i++) {
        const id = rels[i].getAttribute("Id");
        const target = rels[i].getAttribute("Target");
        if (id && target) {
          presRelMap[id] = target;
        }
      }
    } else {
      presRelsDoc = parser.parseFromString(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
        "application/xml"
      );
    }

    // 4. Parse Theme (ppt/theme/theme1.xml)
    const themeColors: Record<string, string> = {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#1f497d",
      lt2: "#eeece1",
      accent1: "#4f81bd",
      accent2: "#c0504d",
      accent3: "#9bbb59",
      accent4: "#8064a2",
      accent5: "#4bacc6",
      accent6: "#f79646",
      hlink: "#0000ff",
      folHlink: "#800080",
    };

    const themeFile = zip.file("ppt/theme/theme1.xml");
    if (themeFile) {
      try {
        const themeXml = await themeFile.async("text");
        const themeDoc = parser.parseFromString(themeXml, "application/xml");
        const clrScheme = themeDoc.getElementsByTagNameNS("*", "clrScheme")[0];
        if (clrScheme) {
          for (let i = 0; i < clrScheme.children.length; i++) {
            const child = clrScheme.children[i];
            const name = child.localName;
            const srgb = child.getElementsByTagNameNS("*", "srgbClr")[0];
            const sys = child.getElementsByTagNameNS("*", "sysClr")[0];
            if (srgb && srgb.getAttribute("val")) {
              themeColors[name] = `#${srgb.getAttribute("val")}`;
            } else if (sys && sys.getAttribute("lastClr")) {
              themeColors[name] = `#${sys.getAttribute("lastClr")}`;
            }
          }
        }
      } catch (e) {
        console.warn("Could not parse ppt/theme/theme1.xml:", e);
      }
    }

    const theme: PptxTheme = { colors: themeColors };

    // 5. Identify Slide Order from <p:sldIdLst>
    const sldIdNodes = presDoc.getElementsByTagNameNS("*", "sldId");
    const slideEntries: PptxSlideEntry[] = [];

    if (sldIdNodes.length > 0) {
      for (let i = 0; i < sldIdNodes.length; i++) {
        const rId =
          sldIdNodes[i].getAttribute("r:id") ||
          sldIdNodes[i].getAttributeNS(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "id"
          ) ||
          sldIdNodes[i].getAttribute("id") ||
          "";
        const slideId = parseInt(sldIdNodes[i].getAttribute("id") || "256", 10);
        if (rId && presRelMap[rId]) {
          let target = presRelMap[rId];
          if (!target.startsWith("ppt/")) {
            target = `ppt/${target.replace(/^\//, "")}`;
          }
          const slideDir = target.substring(0, target.lastIndexOf("/"));
          const slideFileName = target.substring(target.lastIndexOf("/") + 1);
          const relsPath = `${slideDir}/_rels/${slideFileName}.rels`;
          slideEntries.push({ rId, slideId, path: target, relsPath });
        }
      }
    } else {
      // Fallback: search zip for slides/slide*.xml
      const files = Object.keys(zip.files).filter(
        (f) => f.startsWith("ppt/slides/slide") && f.endsWith(".xml") && !f.includes("_rels")
      );
      files.sort((a, b) => {
        const numA = parseInt(a.replace(/[^0-9]/g, ""), 10) || 0;
        const numB = parseInt(b.replace(/[^0-9]/g, ""), 10) || 0;
        return numA - numB;
      });
      for (let i = 0; i < files.length; i++) {
        const target = files[i];
        const slideDir = target.substring(0, target.lastIndexOf("/"));
        const slideFileName = target.substring(target.lastIndexOf("/") + 1);
        const relsPath = `${slideDir}/_rels/${slideFileName}.rels`;
        slideEntries.push({
          rId: `slide_${i + 1}`,
          slideId: 256 + i,
          path: target,
          relsPath,
        });
      }
    }

    // 6. Preload Media Files into Data URLs
    const mediaDataUrls = new Map<string, string>();
    const mediaFiles = Object.keys(zip.files).filter(
      (f) => f.startsWith("ppt/media/") && !zip.files[f].dir
    );

    for (const mf of mediaFiles) {
      try {
        const fileObj = zip.file(mf);
        if (fileObj) {
          const ext = mf.split(".").pop()?.toLowerCase() || "png";
          const mime = getMimeType(ext);
          const base64 = await fileObj.async("base64");
          const dataUrl = `data:${mime};base64,${base64}`;
          mediaDataUrls.set(mf, dataUrl);
          mediaDataUrls.set(mf.replace(/^ppt\//, ""), dataUrl);
          mediaDataUrls.set(mf.split("/").pop()!, dataUrl);
        }
      } catch (err) {
        console.warn(`Could not preload media file ${mf}:`, err);
      }
    }

    // 7. Parse Slides
    const slides: PptxSlide[] = [];
    const slideXmlDocs = new Map<string, Document>();
    const slideRelsDocs = new Map<string, Document>();
    const notesXmlDocs = new Map<string, Document>();
    let presentationTitle = "";

    for (let i = 0; i < slideEntries.length; i++) {
      const entry = slideEntries[i];
      const slideFile = zip.file(entry.path);
      if (!slideFile) continue;

      const slideXml = await slideFile.async("text");
      const slideDoc = parser.parseFromString(slideXml, "application/xml");
      slideXmlDocs.set(entry.path, slideDoc);

      // Slide Relationships
      const relsFile = zip.file(entry.relsPath);
      const slideRelMap: Record<string, string> = {};

      if (relsFile) {
        try {
          const relsXml = await relsFile.async("text");
          const relsDoc = parser.parseFromString(relsXml, "application/xml");
          slideRelsDocs.set(entry.relsPath, relsDoc);
          const relNodes = relsDoc.getElementsByTagNameNS("*", "Relationship");
          for (let j = 0; j < relNodes.length; j++) {
            const relId = relNodes[j].getAttribute("Id");
            const target = relNodes[j].getAttribute("Target");
            if (relId && target) {
              slideRelMap[relId] = target;
            }
          }
        } catch (e) {
          console.warn(`Could not parse rels for ${entry.path}:`, e);
        }
      }

      // Speaker Notes (from notesSlide)
      let notes = "";
      for (const [rKey, target] of Object.entries(slideRelMap)) {
        if (target.includes("notesSlide")) {
          entry.notesRId = rKey;
          entry.notesPath = target.startsWith("../")
            ? `ppt/${target.substring(3)}`
            : target.startsWith("ppt/")
            ? target
            : `ppt/notesSlides/${target.split("/").pop()}`;

          const notesFile = zip.file(entry.notesPath);
          if (notesFile) {
            try {
              const notesXml = await notesFile.async("text");
              const notesDoc = parser.parseFromString(notesXml, "application/xml");
              notesXmlDocs.set(entry.notesPath, notesDoc);
              notes = parseSpeakerNotes(notesDoc);
            } catch (err) {
              console.warn(`Error reading speaker notes for slide ${i + 1}:`, err);
            }
          }
          break;
        }
      }

      // Slide Transition
      const transition = parseSlideTransition(slideDoc);

      // Slide Background
      const background = parseSlideBackground(slideDoc, themeColors, mediaDataUrls);

      // Slide Shapes & Elements
      const spTree = slideDoc.getElementsByTagNameNS("*", "spTree")[0];
      const elements: PptxElement[] = [];
      let detectedTitle = "";

      if (spTree) {
        parseShapeTree(
          spTree,
          elements,
          slideRelMap,
          mediaDataUrls,
          themeColors,
          width,
          height,
          (t) => {
            if (!detectedTitle) detectedTitle = t;
          }
        );
      }

      if (i === 0 && detectedTitle && !presentationTitle) {
        presentationTitle = detectedTitle;
      }

      slides.push({
        index: i,
        slideNumber: i + 1,
        id: entry.rId,
        title: detectedTitle || `Slide ${i + 1}`,
        background,
        elements,
        transition,
        notes,
      });
    }

    const presentation: PptxPresentation = {
      title: presentationTitle || "Untitled Presentation",
      width,
      height,
      emuWidth,
      emuHeight,
      slides,
      theme,
    };

    const doc = new PptxDocument(
      zip,
      presentation,
      slideEntries,
      presDoc,
      presRelsDoc,
      contentTypesDoc
    );
    doc.slideXmlDocs = slideXmlDocs;
    doc.slideRelsDocs = slideRelsDocs;
    doc.notesXmlDocs = notesXmlDocs;
    doc.mediaDataUrls = mediaDataUrls;
    return doc;
  }

  getPresentation(): PptxPresentation {
    return this.presentation;
  }

  getSlideCount(): number {
    return this.presentation.slides.length;
  }

  getSlide(index: number): PptxSlide | undefined {
    return this.presentation.slides[index];
  }

  getSlideDimensions(): { width: number; height: number } {
    return {
      width: this.presentation.width,
      height: this.presentation.height,
    };
  }

  getSlideEntries(): ReadonlyArray<PptxSlideEntry> {
    return this.slideEntries;
  }

  isDirty(): boolean {
    return this.dirty;
  }

  setDirty(val: boolean): void {
    this.dirty = val;
  }

  exportMarkdown(fileName: string): string {
    return PptxExport.toMarkdown(this.presentation, fileName);
  }

  // -------------------------------------------------------------
  // CRUD OPERATIONS: TEXT & ELEMENTS
  // -------------------------------------------------------------

  /**
   * Updates text in an existing element (shape, text box, placeholder) on a slide.
   * Modifies both the in-memory domain model and the underlying OOXML slide DOM.
   */
  updateElementText(slideIndex: number, elementId: string, newText: string): boolean {
    const slide = this.presentation.slides[slideIndex];
    const entry = this.slideEntries[slideIndex];
    if (!slide || !entry) return false;

    const slideDoc = this.slideXmlDocs.get(entry.path);
    if (!slideDoc) return false;

    // 1. Locate shape in OOXML DOM
    let targetSp: Element | null = null;
    const sps = slideDoc.getElementsByTagNameNS("*", "sp");
    for (let i = 0; i < sps.length; i++) {
      const cNvPr = sps[i].getElementsByTagNameNS("*", "cNvPr")[0];
      if (
        cNvPr &&
        (cNvPr.getAttribute("id") === elementId ||
          cNvPr.getAttribute("name") === elementId ||
          sps[i].getAttribute("id") === elementId)
      ) {
        targetSp = sps[i];
        break;
      }
    }

    if (!targetSp) {
      // Check tables/graphicFrames
      const frames = slideDoc.getElementsByTagNameNS("*", "graphicFrame");
      for (let i = 0; i < frames.length; i++) {
        const cNvPr = frames[i].getElementsByTagNameNS("*", "cNvPr")[0];
        if (
          cNvPr &&
          (cNvPr.getAttribute("id") === elementId ||
            cNvPr.getAttribute("name") === elementId)
        ) {
          targetSp = frames[i];
          break;
        }
      }
    }

    // 2. Locate element in domain model
    let targetEl: PptxElement | undefined;
    const findInElements = (list: PptxElement[]): PptxElement | undefined => {
      for (const el of list) {
        if (el.id === elementId || el.name === elementId) return el;
        if (el.children) {
          const found = findInElements(el.children);
          if (found) return found;
        }
      }
      return undefined;
    };
    targetEl = findInElements(slide.elements);

    // 3. Update OOXML <p:txBody>
    if (targetSp) {
      let txBody = targetSp.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) {
        txBody = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/presentationml/2006/main",
          "p:txBody"
        );
        const bodyPr = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/drawingml/2006/main",
          "a:bodyPr"
        );
        const lstStyle = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/drawingml/2006/main",
          "a:lstStyle"
        );
        txBody.appendChild(bodyPr);
        txBody.appendChild(lstStyle);
        targetSp.appendChild(txBody);
      }

      // Preserve formatting from original paragraph & run
      const origRPr = txBody.getElementsByTagNameNS("*", "rPr")[0]?.cloneNode(true) as Element | null;
      const origPPr = txBody.getElementsByTagNameNS("*", "pPr")[0]?.cloneNode(true) as Element | null;

      // Remove existing paragraphs
      const existingPs = Array.from(txBody.getElementsByTagNameNS("*", "p"));
      for (const p of existingPs) {
        txBody.removeChild(p);
      }

      // Add new paragraphs for each line
      const lines = newText.split("\n");
      for (const line of lines) {
        const pEl = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/drawingml/2006/main",
          "a:p"
        );
        if (origPPr) {
          pEl.appendChild(origPPr.cloneNode(true));
        }

        const rEl = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/drawingml/2006/main",
          "a:r"
        );
        if (origRPr) {
          rEl.appendChild(origRPr.cloneNode(true));
        } else {
          const rPr = slideDoc.createElementNS(
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "a:rPr"
          );
          rPr.setAttribute("lang", "en-US");
          rPr.setAttribute("sz", "2000");
          rEl.appendChild(rPr);
        }

        const tEl = slideDoc.createElementNS(
          "http://schemas.openxmlformats.org/drawingml/2006/main",
          "a:t"
        );
        tEl.textContent = line;
        tEl.setAttributeNS(
          "http://www.w3.org/XML/1998/namespace",
          "xml:space",
          "preserve"
        );
        rEl.appendChild(tEl);
        pEl.appendChild(rEl);
        txBody.appendChild(pEl);
      }
    }

    // 4. Update domain model
    const lines = newText.split("\n");
    if (targetEl) {
      const firstP = targetEl.paragraphs?.[0];
      const firstR = firstP?.runs?.[0];
      targetEl.paragraphs = lines.map((line) => ({
        align: firstP?.align || "left",
        runs: [
          {
            text: line,
            fontSize: firstR?.fontSize,
            fontFamily: firstR?.fontFamily,
            color: firstR?.color,
            bold: firstR?.bold,
            italic: firstR?.italic,
            underline: firstR?.underline,
          },
        ],
      }));

      // Update slide title if this was the title placeholder
      const isTitle =
        targetEl.name.toLowerCase().includes("title") ||
        targetEl.id === "2" ||
        slide.elements[0]?.id === elementId;
      if (isTitle) {
        slide.title = lines[0]?.trim() || slide.title;
        if (slideIndex === 0) {
          this.presentation.title = slide.title;
        }
      }
    }

    this.dirty = true;
    return true;
  }

  /**
   * Adds a new shape or text box to a slide.
   */
  addElement(
    slideIndex: number,
    type: "textbox" | "rect" | "roundRect" | "ellipse",
    options?: {
      text?: string;
      x?: number;
      y?: number;
      width?: number;
      height?: number;
      color?: string;
      fontSize?: number;
    }
  ): PptxElement | null {
    const slide = this.presentation.slides[slideIndex];
    const entry = this.slideEntries[slideIndex];
    if (!slide || !entry) return null;

    const slideDoc = this.slideXmlDocs.get(entry.path);
    if (!slideDoc) return null;

    const spTree = slideDoc.getElementsByTagNameNS("*", "spTree")[0];
    if (!spTree) return null;

    // Find next unique shape id
    let maxId = 1;
    const cNvPrs = slideDoc.getElementsByTagNameNS("*", "cNvPr");
    for (let i = 0; i < cNvPrs.length; i++) {
      const idVal = parseInt(cNvPrs[i].getAttribute("id") || "0", 10);
      if (!isNaN(idVal) && idVal > maxId) maxId = idVal;
    }
    const shapeId = maxId + 1;

    const widthPx = options?.width ?? (type === "textbox" ? 400 : 220);
    const heightPx = options?.height ?? (type === "textbox" ? 100 : 160);
    const xPx =
      options?.x ??
      Math.max(40, Math.round((this.presentation.width - widthPx) / 2));
    const yPx =
      options?.y ??
      Math.max(40, Math.round((this.presentation.height - heightPx) / 2));
    const xInEmu = xPx * 9525;
    const yInEmu = yPx * 9525;
    const wInEmu = widthPx * 9525;
    const hInEmu = heightPx * 9525;
    const initialText =
      options?.text ??
      (type === "textbox" ? "Click to edit text" : "Shape");
    const color = options?.color || (type === "textbox" ? "#000000" : "#4f81bd");
    const colorHex = color.replace("#", "");
    const fontSizePt = options?.fontSize ?? (type === "textbox" ? 20 : 18);
    const fontSizeHundredths = fontSizePt * 100;

    let shapeXml = "";
    if (type === "textbox") {
      shapeXml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="${shapeId}" name="TextBox ${shapeId}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${xInEmu}" y="${yInEmu}"/>
      <a:ext cx="${wInEmu}" cy="${hInEmu}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr rtlCol="0" wrap="square"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="${fontSizeHundredths}">
          <a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>
        </a:rPr>
        <a:t xml:space="preserve">${escapeXml(initialText)}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;
    } else {
      shapeXml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="${shapeId}" name="Shape ${shapeId}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${xInEmu}" y="${yInEmu}"/>
      <a:ext cx="${wInEmu}" cy="${hInEmu}"/>
    </a:xfrm>
    <a:prstGeom prst="${type}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="${fontSizeHundredths}">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </a:rPr>
        <a:t xml:space="preserve">${escapeXml(initialText)}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;
    }

    const parser = getXmlParser();
    const parsedDoc = parser.parseFromString(shapeXml, "application/xml");
    const importedNode = slideDoc.importNode(parsedDoc.documentElement, true);
    spTree.appendChild(importedNode);

    // Build domain model element
    const newEl: PptxElement = {
      id: String(shapeId),
      name: type === "textbox" ? `TextBox ${shapeId}` : `Shape ${shapeId}`,
      type: type === "textbox" ? "text" : "shape",
      x: xPx,
      y: yPx,
      width: widthPx,
      height: heightPx,
      rotation: 0,
      zIndex: slide.elements.length,
      shapeType: type === "textbox" ? "rect" : type,
      fill:
        type !== "textbox"
          ? { type: "solid", color: `#${colorHex}` }
          : undefined,
      border:
        type !== "textbox"
          ? { color: "#1e293b", width: 1, style: "solid" }
          : undefined,
      paragraphs: [
        {
          align: type === "textbox" ? "left" : "center",
          runs: [
            {
              text: initialText,
              fontSize: fontSizePt,
              color: type === "textbox" ? `#${colorHex}` : "#ffffff",
            },
          ],
        },
      ],
      verticalAlign: type === "textbox" ? "top" : "middle",
    };

    slide.elements.push(newEl);
    this.dirty = true;
    return newEl;
  }

  /**
   * Deletes an element from a slide by element ID.
   */
  deleteElement(slideIndex: number, elementId: string): boolean {
    const slide = this.presentation.slides[slideIndex];
    const entry = this.slideEntries[slideIndex];
    if (!slide || !entry) return false;

    const slideDoc = this.slideXmlDocs.get(entry.path);
    if (!slideDoc) return false;

    // Remove from slide XML
    const cNvPrs = slideDoc.getElementsByTagNameNS("*", "cNvPr");
    for (let i = 0; i < cNvPrs.length; i++) {
      if (cNvPrs[i].getAttribute("id") === elementId) {
        let parent = cNvPrs[i].parentElement;
        while (parent && parent.localName !== "spTree" && parent.localName !== "grpSp") {
          if (["sp", "pic", "graphicFrame", "cxnSp", "grpSp"].includes(parent.localName)) {
            parent.parentElement?.removeChild(parent);
            break;
          }
          parent = parent.parentElement;
        }
        break;
      }
    }

    // Remove from domain model
    const removeRecursive = (elements: PptxElement[]): PptxElement[] => {
      return elements.filter((el) => {
        if (el.id === elementId) return false;
        if (el.children) el.children = removeRecursive(el.children);
        return true;
      });
    };
    slide.elements = removeRecursive(slide.elements);

    this.dirty = true;
    return true;
  }

  // -------------------------------------------------------------
  // CRUD OPERATIONS: SLIDES
  // -------------------------------------------------------------

  /**
   * Adds a new slide at the target position.
   * Generates standard PowerPoint-compliant OOXML parts and relationships.
   */
  addSlide(targetIndex?: number, title?: string): PptxSlide {
    let insertAt =
      targetIndex === undefined ||
      targetIndex < 0 ||
      targetIndex > this.presentation.slides.length
        ? this.presentation.slides.length
        : targetIndex;

    // Allocate next unique slide file number
    let maxNum = 0;
    const allPaths = [
      ...Object.keys(this.zip.files),
      ...this.slideEntries.map((e) => e.path),
    ];
    for (const p of allPaths) {
      const m = p.match(/ppt\/slides\/slide(\d+)\.xml/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
    const nextNum = maxNum + 1;
    const newPath = `ppt/slides/slide${nextNum}.xml`;
    const newRelsPath = `ppt/slides/_rels/slide${nextNum}.xml.rels`;

    // Allocate next slide ID (>= 256)
    let maxSlideId = 255;
    const sldNodes = this.presDoc.getElementsByTagNameNS("*", "sldId");
    for (let i = 0; i < sldNodes.length; i++) {
      const idVal = parseInt(sldNodes[i].getAttribute("id") || "0", 10);
      if (!isNaN(idVal) && idVal > maxSlideId) maxSlideId = idVal;
    }
    for (const e of this.slideEntries) {
      if (e.slideId > maxSlideId) maxSlideId = e.slideId;
    }
    const nextSlideId = maxSlideId + 1;

    // Allocate next relationship ID (rIdX)
    let maxRIdNum = 0;
    const relNodes = this.presRelsDoc.getElementsByTagNameNS("*", "Relationship");
    for (let i = 0; i < relNodes.length; i++) {
      const idStr = relNodes[i].getAttribute("Id") || "";
      const m = idStr.match(/rId(\d+)/);
      if (m) {
        const num = parseInt(m[1], 10);
        if (num > maxRIdNum) maxRIdNum = num;
      }
    }
    for (const e of this.slideEntries) {
      const m = e.rId.match(/rId(\d+)/);
      if (m) {
        const num = parseInt(m[1], 10);
        if (num > maxRIdNum) maxRIdNum = num;
      }
    }
    const nextRId = `rId${maxRIdNum + 1}`;

    // 1. Add <p:sldId> to presentation.xml <p:sldIdLst>
    let sldIdLst = this.presDoc.getElementsByTagNameNS("*", "sldIdLst")[0];
    if (!sldIdLst) {
      sldIdLst = this.presDoc.createElementNS(
        "http://schemas.openxmlformats.org/presentationml/2006/main",
        "p:sldIdLst"
      );
      this.presDoc.documentElement.appendChild(sldIdLst);
    }

    const newSldIdEl = this.presDoc.createElementNS(
      "http://schemas.openxmlformats.org/presentationml/2006/main",
      "p:sldId"
    );
    newSldIdEl.setAttribute("id", String(nextSlideId));
    newSldIdEl.setAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "r:id",
      nextRId
    );

    if (insertAt < this.slideEntries.length) {
      const targetEntry = this.slideEntries[insertAt];
      const existingNodes = Array.from(sldIdLst.getElementsByTagNameNS("*", "sldId"));
      const refNode = existingNodes.find(
        (n) =>
          n.getAttribute("r:id") === targetEntry.rId ||
          n.getAttributeNS(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "id"
          ) === targetEntry.rId
      );
      if (refNode) {
        sldIdLst.insertBefore(newSldIdEl, refNode);
      } else {
        sldIdLst.appendChild(newSldIdEl);
      }
    } else {
      sldIdLst.appendChild(newSldIdEl);
    }

    // 2. Add relationship to ppt/_rels/presentation.xml.rels
    const newRelEl = this.presRelsDoc.createElementNS(
      "http://schemas.openxmlformats.org/package/2006/relationships",
      "Relationship"
    );
    newRelEl.setAttribute("Id", nextRId);
    newRelEl.setAttribute(
      "Type",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    );
    newRelEl.setAttribute("Target", `slides/slide${nextNum}.xml`);
    this.presRelsDoc.documentElement.appendChild(newRelEl);

    // 3. Add override in [Content_Types].xml
    const newOverride = this.contentTypesDoc.createElementNS(
      "http://schemas.openxmlformats.org/package/2006/content-types",
      "Override"
    );
    newOverride.setAttribute("PartName", `/ppt/slides/slide${nextNum}.xml`);
    newOverride.setAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    );
    this.contentTypesDoc.documentElement.appendChild(newOverride);

    // 4. Slide Layout relationship
    let layoutTarget = "../slideLayouts/slideLayout1.xml";
    let hasLayout = false;
    for (const doc of this.slideRelsDocs.values()) {
      const rels = doc.getElementsByTagNameNS("*", "Relationship");
      for (let i = 0; i < rels.length; i++) {
        const t = rels[i].getAttribute("Type");
        if (t?.includes("relationships/slideLayout")) {
          layoutTarget = rels[i].getAttribute("Target") || layoutTarget;
          hasLayout = true;
          break;
        }
      }
      if (hasLayout) break;
    }

    const parser = getXmlParser();
    const relsXml = hasLayout
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutTarget}"/>
</Relationships>`
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    const relsDoc = parser.parseFromString(relsXml, "application/xml");
    this.slideRelsDocs.set(newRelsPath, relsDoc);

    // 5. Build Slide XML
    const presEmuW = this.presentation.emuWidth || 12192000;
    const presEmuH = this.presentation.emuHeight || 6858000;
    const titleX = Math.round(presEmuW * 0.08);
    const titleY = Math.round(presEmuH * 0.1);
    const titleW = Math.round(presEmuW * 0.84);
    const titleH = Math.round(presEmuH * 0.18);
    const bodyX = titleX;
    const bodyY = Math.round(presEmuH * 0.32);
    const bodyW = titleW;
    const bodyH = Math.round(presEmuH * 0.55);

    const titleText = title || `Slide ${insertAt + 1}`;
    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:grpSpPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${titleX}" y="${titleY}"/><a:ext cx="${titleW}" cy="${titleH}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="3600" b="1">
                <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
              </a:rPr>
              <a:t xml:space="preserve">${escapeXml(titleText)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${bodyX}" y="${bodyY}"/><a:ext cx="${bodyW}" cy="${bodyH}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:buChar char="•"/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="2000">
                <a:solidFill><a:srgbClr val="334155"/></a:solidFill>
              </a:rPr>
              <a:t xml:space="preserve">Click to edit text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMap/></p:clrMapOvr>
</p:sld>`;

    const slideDoc = parser.parseFromString(slideXml, "application/xml");
    this.slideXmlDocs.set(newPath, slideDoc);

    // 6. Record slide entry
    const newEntry: PptxSlideEntry = {
      rId: nextRId,
      slideId: nextSlideId,
      path: newPath,
      relsPath: newRelsPath,
    };
    this.slideEntries.splice(insertAt, 0, newEntry);

    // 7. Build domain model
    const newSlide: PptxSlide = {
      index: insertAt,
      slideNumber: insertAt + 1,
      id: nextRId,
      title: titleText,
      background: "#ffffff",
      elements: [
        {
          id: "2",
          name: "Title 1",
          type: "text",
          x: Math.round(titleX / 9525),
          y: Math.round(titleY / 9525),
          width: Math.round(titleW / 9525),
          height: Math.round(titleH / 9525),
          rotation: 0,
          zIndex: 0,
          paragraphs: [
            {
              runs: [
                {
                  text: titleText,
                  fontSize: 36,
                  bold: true,
                  color: "#1e293b",
                },
              ],
            },
          ],
        },
        {
          id: "3",
          name: "Content 2",
          type: "text",
          x: Math.round(bodyX / 9525),
          y: Math.round(bodyY / 9525),
          width: Math.round(bodyW / 9525),
          height: Math.round(bodyH / 9525),
          rotation: 0,
          zIndex: 1,
          paragraphs: [
            {
              bullet: { type: "char", char: "•" },
              runs: [
                {
                  text: "Click to edit text",
                  fontSize: 20,
                  color: "#334155",
                },
              ],
            },
          ],
        },
      ],
      transition: { type: "fade", durationMs: 400, advanceOnClick: true },
      notes: "",
    };

    this.presentation.slides.splice(insertAt, 0, newSlide);

    // Re-index remaining slides
    for (let i = 0; i < this.presentation.slides.length; i++) {
      this.presentation.slides[i].index = i;
      this.presentation.slides[i].slideNumber = i + 1;
    }

    this.dirty = true;
    return newSlide;
  }

  /**
   * Duplicates an existing slide with all its elements, styling, and transitions.
   */
  duplicateSlide(slideIndex: number): PptxSlide {
    const srcEntry = this.slideEntries[slideIndex];
    const srcSlide = this.presentation.slides[slideIndex];
    if (!srcEntry || !srcSlide) {
      throw new Error(`Slide at index ${slideIndex} does not exist`);
    }

    const srcDoc = this.slideXmlDocs.get(srcEntry.path);
    if (!srcDoc) {
      throw new Error(`Slide XML for ${srcEntry.path} not found`);
    }

    const insertAt = slideIndex + 1;

    // Allocate next unique slide file number
    let maxNum = 0;
    const allPaths = [
      ...Object.keys(this.zip.files),
      ...this.slideEntries.map((e) => e.path),
    ];
    for (const p of allPaths) {
      const m = p.match(/ppt\/slides\/slide(\d+)\.xml/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
    const nextNum = maxNum + 1;
    const newPath = `ppt/slides/slide${nextNum}.xml`;
    const newRelsPath = `ppt/slides/_rels/slide${nextNum}.xml.rels`;

    // Allocate next slide ID (>= 256)
    let maxSlideId = 255;
    const sldNodes = this.presDoc.getElementsByTagNameNS("*", "sldId");
    for (let i = 0; i < sldNodes.length; i++) {
      const idVal = parseInt(sldNodes[i].getAttribute("id") || "0", 10);
      if (!isNaN(idVal) && idVal > maxSlideId) maxSlideId = idVal;
    }
    for (const e of this.slideEntries) {
      if (e.slideId > maxSlideId) maxSlideId = e.slideId;
    }
    const nextSlideId = maxSlideId + 1;

    // Allocate next relationship ID (rIdX)
    let maxRIdNum = 0;
    const relNodes = this.presRelsDoc.getElementsByTagNameNS("*", "Relationship");
    for (let i = 0; i < relNodes.length; i++) {
      const idStr = relNodes[i].getAttribute("Id") || "";
      const m = idStr.match(/rId(\d+)/);
      if (m) {
        const num = parseInt(m[1], 10);
        if (num > maxRIdNum) maxRIdNum = num;
      }
    }
    for (const e of this.slideEntries) {
      const m = e.rId.match(/rId(\d+)/);
      if (m) {
        const num = parseInt(m[1], 10);
        if (num > maxRIdNum) maxRIdNum = num;
      }
    }
    const nextRId = `rId${maxRIdNum + 1}`;

    // 1. In presentation.xml: insert <p:sldId> right after srcEntry
    const sldIdLst = this.presDoc.getElementsByTagNameNS("*", "sldIdLst")[0];
    const newSldIdEl = this.presDoc.createElementNS(
      "http://schemas.openxmlformats.org/presentationml/2006/main",
      "p:sldId"
    );
    newSldIdEl.setAttribute("id", String(nextSlideId));
    newSldIdEl.setAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "r:id",
      nextRId
    );

    const existingNodes = Array.from(sldIdLst.getElementsByTagNameNS("*", "sldId"));
    const srcNode = existingNodes.find(
      (n) =>
        n.getAttribute("r:id") === srcEntry.rId ||
        n.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id"
        ) === srcEntry.rId
    );

    if (srcNode && srcNode.nextSibling) {
      sldIdLst.insertBefore(newSldIdEl, srcNode.nextSibling);
    } else {
      sldIdLst.appendChild(newSldIdEl);
    }

    // 2. In ppt/_rels/presentation.xml.rels
    const newRelEl = this.presRelsDoc.createElementNS(
      "http://schemas.openxmlformats.org/package/2006/relationships",
      "Relationship"
    );
    newRelEl.setAttribute("Id", nextRId);
    newRelEl.setAttribute(
      "Type",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    );
    newRelEl.setAttribute("Target", `slides/slide${nextNum}.xml`);
    this.presRelsDoc.documentElement.appendChild(newRelEl);

    // 3. In [Content_Types].xml
    const newOverride = this.contentTypesDoc.createElementNS(
      "http://schemas.openxmlformats.org/package/2006/content-types",
      "Override"
    );
    newOverride.setAttribute("PartName", `/ppt/slides/slide${nextNum}.xml`);
    newOverride.setAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    );
    this.contentTypesDoc.documentElement.appendChild(newOverride);

    // 4. Clone slide XML document
    const parser = getXmlParser();
    const clonedDoc = parser.parseFromString(serializeXml(srcDoc), "application/xml");
    this.slideXmlDocs.set(newPath, clonedDoc);

    // 5. Clone rels XML document (excluding notes relation to avoid collision)
    const srcRels = this.slideRelsDocs.get(srcEntry.relsPath);
    let clonedRelsDoc: Document;
    if (srcRels) {
      clonedRelsDoc = parser.parseFromString(serializeXml(srcRels), "application/xml");
      // Remove notes relationship from clone
      const relNodes = Array.from(
        clonedRelsDoc.getElementsByTagNameNS("*", "Relationship")
      );
      for (const r of relNodes) {
        if (r.getAttribute("Type")?.includes("notesSlide")) {
          r.parentElement?.removeChild(r);
        }
      }
    } else {
      clonedRelsDoc = parser.parseFromString(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
        "application/xml"
      );
    }
    this.slideRelsDocs.set(newRelsPath, clonedRelsDoc);

    // 6. Record slide entry
    const newEntry: PptxSlideEntry = {
      rId: nextRId,
      slideId: nextSlideId,
      path: newPath,
      relsPath: newRelsPath,
    };
    this.slideEntries.splice(insertAt, 0, newEntry);

    // 7. Clone domain model
    const clonedElements: PptxElement[] = JSON.parse(
      JSON.stringify(srcSlide.elements)
    );
    const clonedSlide: PptxSlide = {
      index: insertAt,
      slideNumber: insertAt + 1,
      id: nextRId,
      title: `${srcSlide.title} (Copy)`,
      background: srcSlide.background,
      elements: clonedElements,
      transition: srcSlide.transition ? { ...srcSlide.transition } : undefined,
      notes: srcSlide.notes,
    };

    this.presentation.slides.splice(insertAt, 0, clonedSlide);

    // Re-index remaining slides
    for (let i = 0; i < this.presentation.slides.length; i++) {
      this.presentation.slides[i].index = i;
      this.presentation.slides[i].slideNumber = i + 1;
    }

    this.dirty = true;
    return clonedSlide;
  }

  /**
   * Deletes a slide at the specified index.
   * Completely purges the slide XML, rels, and OPC overrides.
   */
  deleteSlide(slideIndex: number): boolean {
    if (this.presentation.slides.length <= 1) {
      return false; // Cannot delete the only slide
    }

    const entry = this.slideEntries[slideIndex];
    if (!entry) return false;

    // 1. Remove from presentation.xml <p:sldIdLst>
    const sldIdLst = this.presDoc.getElementsByTagNameNS("*", "sldIdLst")[0];
    if (sldIdLst) {
      const sldNodes = Array.from(sldIdLst.getElementsByTagNameNS("*", "sldId"));
      for (const node of sldNodes) {
        const rId =
          node.getAttribute("r:id") ||
          node.getAttributeNS(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "id"
          );
        if (rId === entry.rId) {
          sldIdLst.removeChild(node);
          break;
        }
      }
    }

    // 2. Remove relationship from ppt/_rels/presentation.xml.rels
    const relNodes = Array.from(
      this.presRelsDoc.getElementsByTagNameNS("*", "Relationship")
    );
    for (const rel of relNodes) {
      if (rel.getAttribute("Id") === entry.rId) {
        rel.parentElement?.removeChild(rel);
        break;
      }
    }

    // 3. Remove override from [Content_Types].xml
    const ctOverrides = Array.from(
      this.contentTypesDoc.getElementsByTagNameNS("*", "Override")
    );
    for (const ov of ctOverrides) {
      const pn = ov.getAttribute("PartName");
      if (pn === `/${entry.path}` || pn === entry.path) {
        ov.parentElement?.removeChild(ov);
        break;
      }
    }

    // Clean up notes slide if present and unshared
    if (entry.notesPath) {
      const isShared = this.slideEntries.some(
        (e, idx) => idx !== slideIndex && e.notesPath === entry.notesPath
      );
      if (!isShared) {
        for (const ov of ctOverrides) {
          const pn = ov.getAttribute("PartName");
          if (pn === `/${entry.notesPath}` || pn === entry.notesPath) {
            ov.parentElement?.removeChild(ov);
            break;
          }
        }
        this.deletedFiles.add(entry.notesPath);
        this.notesXmlDocs.delete(entry.notesPath);
      }
    }

    // 4. Mark files for deletion
    this.deletedFiles.add(entry.path);
    this.deletedFiles.add(entry.relsPath);
    this.slideXmlDocs.delete(entry.path);
    this.slideRelsDocs.delete(entry.relsPath);

    // 5. Remove from internal collections
    this.slideEntries.splice(slideIndex, 1);
    this.presentation.slides.splice(slideIndex, 1);

    // 6. Re-index slides
    for (let i = 0; i < this.presentation.slides.length; i++) {
      this.presentation.slides[i].index = i;
      this.presentation.slides[i].slideNumber = i + 1;
    }

    this.dirty = true;
    return true;
  }

  /**
   * Reorders a slide from one position to another.
   */
  reorderSlide(fromIndex: number, toIndex: number): boolean {
    if (
      fromIndex === toIndex ||
      fromIndex < 0 ||
      fromIndex >= this.presentation.slides.length ||
      toIndex < 0 ||
      toIndex >= this.presentation.slides.length
    ) {
      return false;
    }

    const sldIdLst = this.presDoc.getElementsByTagNameNS("*", "sldIdLst")[0];
    if (sldIdLst) {
      const [movedEntry] = this.slideEntries.splice(fromIndex, 1);
      this.slideEntries.splice(toIndex, 0, movedEntry);

      const [movedSlide] = this.presentation.slides.splice(fromIndex, 1);
      this.presentation.slides.splice(toIndex, 0, movedSlide);

      // Re-order DOM elements in <p:sldIdLst>
      const sldNodes = Array.from(
        sldIdLst.getElementsByTagNameNS("*", "sldId")
      );
      for (const n of sldNodes) {
        sldIdLst.removeChild(n);
      }

      for (const entry of this.slideEntries) {
        const matchingNode = sldNodes.find(
          (n) =>
            (n.getAttribute("r:id") ||
              n.getAttributeNS(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "id"
              )) === entry.rId
        );
        if (matchingNode) {
          sldIdLst.appendChild(matchingNode);
        }
      }
    }

    // Re-index slides
    for (let i = 0; i < this.presentation.slides.length; i++) {
      this.presentation.slides[i].index = i;
      this.presentation.slides[i].slideNumber = i + 1;
    }

    this.dirty = true;
    return true;
  }

  // -------------------------------------------------------------
  // CRUD OPERATIONS: SPEAKER NOTES
  // -------------------------------------------------------------

  /**
   * Updates or creates speaker notes for a slide.
   */
  updateSpeakerNotes(slideIndex: number, notesText: string): boolean {
    const slide = this.presentation.slides[slideIndex];
    const entry = this.slideEntries[slideIndex];
    if (!slide || !entry) return false;

    slide.notes = notesText;

    if (entry.notesPath && this.notesXmlDocs.has(entry.notesPath)) {
      // Modify existing notes slide
      const notesDoc = this.notesXmlDocs.get(entry.notesPath)!;
      const sps = notesDoc.getElementsByTagNameNS("*", "sp");
      let notesSp: Element | null = null;
      for (let i = 0; i < sps.length; i++) {
        const ph = sps[i].getElementsByTagNameNS("*", "ph")[0];
        if (ph && ph.getAttribute("type") === "body") {
          notesSp = sps[i];
          break;
        }
      }
      if (!notesSp && sps.length > 0) {
        notesSp = sps[sps.length - 1];
      }

      if (notesSp) {
        let txBody = notesSp.getElementsByTagNameNS("*", "txBody")[0];
        if (!txBody) {
          txBody = notesDoc.createElementNS(
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "p:txBody"
          );
          txBody.appendChild(
            notesDoc.createElementNS(
              "http://schemas.openxmlformats.org/drawingml/2006/main",
              "a:bodyPr"
            )
          );
          txBody.appendChild(
            notesDoc.createElementNS(
              "http://schemas.openxmlformats.org/drawingml/2006/main",
              "a:lstStyle"
            )
          );
          notesSp.appendChild(txBody);
        }

        const ps = Array.from(txBody.getElementsByTagNameNS("*", "p"));
        for (const p of ps) {
          txBody.removeChild(p);
        }

        const lines = notesText.split("\n");
        for (const line of lines) {
          const p = notesDoc.createElementNS(
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "a:p"
          );
          const r = notesDoc.createElementNS(
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "a:r"
          );
          const t = notesDoc.createElementNS(
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "a:t"
          );
          t.textContent = line;
          t.setAttributeNS(
            "http://www.w3.org/XML/1998/namespace",
            "xml:space",
            "preserve"
          );
          r.appendChild(t);
          p.appendChild(r);
          txBody.appendChild(p);
        }
      }
    } else if (notesText.trim()) {
      // Create new notes slide
      let maxNotesNum = 0;
      for (const p of Object.keys(this.zip.files)) {
        const m = p.match(/ppt\/notesSlides\/notesSlide(\d+)\.xml/);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxNotesNum) maxNotesNum = n;
        }
      }
      for (const e of this.slideEntries) {
        if (e.notesPath) {
          const m = e.notesPath.match(/ppt\/notesSlides\/notesSlide(\d+)\.xml/);
          if (m) {
            const n = parseInt(m[1], 10);
            if (n > maxNotesNum) maxNotesNum = n;
          }
        }
      }

      const nextNotesNum = maxNotesNum + 1;
      const newNotesPath = `ppt/notesSlides/notesSlide${nextNotesNum}.xml`;
      const parser = getXmlParser();

      const lines = notesText.split("\n");
      const pTags = lines
        .map(
          (line) =>
            `<a:p><a:r><a:t xml:space="preserve">${escapeXml(line)}</a:t></a:r></a:p>`
        )
        .join("");

      const notesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:grpSpPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${pTags}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;

      const notesDoc = parser.parseFromString(notesXml, "application/xml");
      this.notesXmlDocs.set(newNotesPath, notesDoc);

      // Add relationship to slide rels
      let slideRelsDoc = this.slideRelsDocs.get(entry.relsPath);
      if (!slideRelsDoc) {
        slideRelsDoc = parser.parseFromString(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
          "application/xml"
        );
        this.slideRelsDocs.set(entry.relsPath, slideRelsDoc);
      }

      const notesRId = `rIdNotes${nextNotesNum}`;
      const relNode = slideRelsDoc.createElementNS(
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "Relationship"
      );
      relNode.setAttribute("Id", notesRId);
      relNode.setAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
      );
      relNode.setAttribute(
        "Target",
        `../notesSlides/notesSlide${nextNotesNum}.xml`
      );
      slideRelsDoc.documentElement.appendChild(relNode);

      // Add override in [Content_Types].xml
      const ctOverride = this.contentTypesDoc.createElementNS(
        "http://schemas.openxmlformats.org/package/2006/content-types",
        "Override"
      );
      ctOverride.setAttribute("PartName", `/${newNotesPath}`);
      ctOverride.setAttribute(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
      );
      this.contentTypesDoc.documentElement.appendChild(ctOverride);

      entry.notesPath = newNotesPath;
      entry.notesRId = notesRId;
    }

    this.dirty = true;
    return true;
  }

  // -------------------------------------------------------------
  // SAVE / PERSISTENCE
  // -------------------------------------------------------------

  /**
   * Serializes all modified OOXML parts and generates a binary .pptx zip buffer.
   * Guarantees 100% PowerPoint cross-compatibility and strict XML validity.
   */
  async save(): Promise<Uint8Array> {
    // 1. Serialize presentation.xml
    this.zip.file("ppt/presentation.xml", serializeXml(this.presDoc));

    // 2. Serialize presentation.xml.rels
    this.zip.file(
      "ppt/_rels/presentation.xml.rels",
      serializeXml(this.presRelsDoc)
    );

    // 3. Serialize [Content_Types].xml
    this.zip.file("[Content_Types].xml", serializeXml(this.contentTypesDoc));

    // 4. Serialize slide XML documents
    for (const [path, doc] of this.slideXmlDocs.entries()) {
      if (!this.deletedFiles.has(path)) {
        this.zip.file(path, serializeXml(doc));
      }
    }

    // 5. Serialize slide rels XML documents
    for (const [relsPath, doc] of this.slideRelsDocs.entries()) {
      if (!this.deletedFiles.has(relsPath)) {
        this.zip.file(relsPath, serializeXml(doc));
      }
    }

    // 6. Serialize notes XML documents
    for (const [notesPath, doc] of this.notesXmlDocs.entries()) {
      if (!this.deletedFiles.has(notesPath)) {
        this.zip.file(notesPath, serializeXml(doc));
      }
    }

    // 7. Remove deleted files from zip
    for (const delPath of this.deletedFiles) {
      this.zip.remove(delPath);
    }
    this.deletedFiles.clear();

    this.dirty = false;
    return await this.zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
    });
  }
}

// -------------------------------------------------------------
// HELPER FUNCTIONS FOR OOXML PARSING & SERIALIZATION
// -------------------------------------------------------------

function getXmlParser(): DOMParser {
  return typeof DOMParser !== "undefined"
    ? new DOMParser()
    : new (require("@xmldom/xmldom").DOMParser)();
}

function serializeXml(doc: Document): string {
  const serializer =
    typeof XMLSerializer !== "undefined"
      ? new XMLSerializer()
      : new (require("@xmldom/xmldom").XMLSerializer)();
  let str = serializer.serializeToString(doc);
  if (!str.startsWith("<?xml")) {
    str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + str;
  }
  return str;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getMimeType(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "svg":
      return "image/svg+xml";
    case "webp":
      return "image/webp";
    case "bmp":
      return "image/bmp";
    default:
      return "image/png";
  }
}

function parseSpeakerNotes(doc: Document): string {
  const paragraphs: string[] = [];
  const sps = doc.getElementsByTagNameNS("*", "sp");

  for (let i = 0; i < sps.length; i++) {
    const sp = sps[i];
    const ph = sp.getElementsByTagNameNS("*", "ph")[0];
    const phType = ph?.getAttribute("type");

    // Skip slide number, header, footer, datetime placeholders
    if (
      phType === "sldNum" ||
      phType === "hdr" ||
      phType === "ftr" ||
      phType === "dt"
    ) {
      continue;
    }

    const txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
    if (!txBody) continue;

    const ps = txBody.getElementsByTagNameNS("*", "p");
    for (let j = 0; j < ps.length; j++) {
      const ts = ps[j].getElementsByTagNameNS("*", "t");
      let pText = "";
      for (let k = 0; k < ts.length; k++) {
        pText += ts[k].textContent || "";
      }
      const trimmed = pText.trim();
      if (trimmed) {
        paragraphs.push(trimmed);
      }
    }
  }

  return paragraphs.join("\n\n");
}

function parseSlideTransition(slideDoc: Document): PptxTransition {
  const transNode = slideDoc.getElementsByTagNameNS("*", "transition")[0];
  if (!transNode) {
    return { ...DEFAULT_TRANSITION };
  }

  const speedAttr = transNode.getAttribute("spd") || "med";
  let durationMs = 400;
  if (speedAttr === "slow") durationMs = 800;
  else if (speedAttr === "fast") durationMs = 250;

  const durAttr = transNode.getAttribute("dur");
  if (durAttr) {
    const parsedDur = parseInt(durAttr, 10);
    if (!isNaN(parsedDur) && parsedDur > 0) {
      durationMs = parsedDur;
    }
  }

  let type: TransitionType = "fade";
  let direction: TransitionDirection | undefined = undefined;

  for (let i = 0; i < transNode.children.length; i++) {
    const child = transNode.children[i];
    const local = child.localName.toLowerCase();
    const dir = child.getAttribute("dir") as TransitionDirection | null;

    if (local === "fade") {
      type = "fade";
      break;
    } else if (local === "push") {
      type = "push";
      direction = dir || "l";
      break;
    } else if (local === "wipe") {
      type = "wipe";
      direction = dir || "l";
      break;
    } else if (local === "zoom") {
      type = "zoom";
      break;
    } else if (local === "flip") {
      type = "flip";
      break;
    } else if (local === "dissolve") {
      type = "dissolve";
      break;
    } else if (local === "split") {
      type = "split";
      direction = dir || "in";
      break;
    } else if (local === "wheel") {
      type = "wheel";
      break;
    } else if (local === "cut") {
      type = "cut";
      break;
    } else if (local === "cover") {
      type = "cover";
      direction = dir || "l";
      break;
    } else if (local === "pull") {
      type = "pull";
      direction = dir || "l";
      break;
    }
  }

  return {
    type,
    direction,
    durationMs,
    advanceOnClick: true,
  };
}

function parseSlideBackground(
  slideDoc: Document,
  themeColors: Record<string, string>,
  mediaDataUrls: Map<string, string>
): string | undefined {
  const bg = slideDoc.getElementsByTagNameNS("*", "bg")[0];
  if (!bg) return undefined;

  const bgPr = bg.getElementsByTagNameNS("*", "bgPr")[0];
  if (!bgPr) return undefined;

  // Solid fill
  const solidFill = bgPr.getElementsByTagNameNS("*", "solidFill")[0];
  if (solidFill) {
    return resolveColor(solidFill, themeColors);
  }

  // Gradient fill
  const gradFill = bgPr.getElementsByTagNameNS("*", "gradFill")[0];
  if (gradFill) {
    const gsLst = gradFill.getElementsByTagNameNS("*", "gsLst")[0];
    if (gsLst) {
      const gsNodes = gsLst.getElementsByTagNameNS("*", "gs");
      const stops: string[] = [];
      for (let i = 0; i < gsNodes.length; i++) {
        const pos = Number(gsNodes[i].getAttribute("pos")) / 1000;
        const clr = resolveColor(gsNodes[i], themeColors) || "#ffffff";
        stops.push(`${clr} ${pos}%`);
      }
      if (stops.length > 0) {
        return `linear-gradient(180deg, ${stops.join(", ")})`;
      }
    }
  }

  return undefined;
}

function parseShapeTree(
  spTree: Element,
  elements: PptxElement[],
  slideRelMap: Record<string, string>,
  mediaDataUrls: Map<string, string>,
  themeColors: Record<string, string>,
  slideW: number,
  slideH: number,
  onTitleDetected: (title: string) => void
): void {
  let zIndex = 0;

  for (let i = 0; i < spTree.children.length; i++) {
    const child = spTree.children[i];
    const tag = child.localName;

    // 1. Shapes (<p:sp>)
    if (tag === "sp") {
      const el = parseSpElement(
        child,
        zIndex++,
        slideRelMap,
        themeColors,
        onTitleDetected
      );
      if (el) elements.push(el);
    }
    // 2. Pictures (<p:pic>)
    else if (tag === "pic") {
      const el = parsePicElement(child, zIndex++, slideRelMap, mediaDataUrls);
      if (el) elements.push(el);
    }
    // 3. Graphic Frames (<p:graphicFrame>) -> Tables
    else if (tag === "graphicFrame") {
      const el = parseGraphicFrameElement(child, zIndex++, themeColors);
      if (el) elements.push(el);
    }
    // 4. Group Shapes (<p:grpSp>)
    else if (tag === "grpSp") {
      const el = parseGroupElement(
        child,
        zIndex++,
        slideRelMap,
        mediaDataUrls,
        themeColors,
        slideW,
        slideH,
        onTitleDetected
      );
      if (el) elements.push(el);
    }
    // 5. Connectors (<p:cxnSp>)
    else if (tag === "cxnSp") {
      const el = parseConnectorElement(child, zIndex++, themeColors);
      if (el) elements.push(el);
    }
  }
}

function parseSpElement(
  sp: Element,
  zIndex: number,
  slideRelMap: Record<string, string>,
  themeColors: Record<string, string>,
  onTitleDetected: (title: string) => void
): PptxElement | null {
  const nvSpPr = sp.getElementsByTagNameNS("*", "nvSpPr")[0];
  const cNvPr = nvSpPr?.getElementsByTagNameNS("*", "cNvPr")[0];
  const id = cNvPr?.getAttribute("id") || `sp_${zIndex}`;
  const name = cNvPr?.getAttribute("name") || "";

  // Title detection
  const ph = nvSpPr?.getElementsByTagNameNS("*", "ph")[0];
  const phType = ph?.getAttribute("type");
  const isTitlePh = phType === "title" || phType === "ctrTitle";

  const spPr = sp.getElementsByTagNameNS("*", "spPr")[0];
  const transform = parseTransform(spPr);

  // Shape geometry
  const prstGeom = spPr?.getElementsByTagNameNS("*", "prstGeom")[0];
  const shapeType = prstGeom?.getAttribute("prst") || "rect";

  // Fill
  let fill: PptxElement["fill"] = undefined;
  const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0];
  if (solidFill) {
    const clr = resolveColor(solidFill, themeColors);
    if (clr) fill = { type: "solid", color: clr };
  }

  // Border / Line
  let border: PptxElement["border"] = undefined;
  const ln = spPr?.getElementsByTagNameNS("*", "ln")[0];
  if (ln) {
    const w = ln.getAttribute("w");
    const widthPx = w ? Math.max(1, Math.round(Number(w) / 12700)) : 1;
    const lnFill = ln.getElementsByTagNameNS("*", "solidFill")[0];
    const clr = lnFill ? resolveColor(lnFill, themeColors) : "#000000";
    if (clr) {
      border = { color: clr, width: widthPx, style: "solid" };
    }
  }

  // Text Body
  const txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
  const paragraphs: PptxParagraph[] = [];
  let verticalAlign: "top" | "middle" | "bottom" = "top";
  let padding: PptxElement["padding"] = undefined;

  if (txBody) {
    const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0];
    if (bodyPr) {
      const anchor = bodyPr.getAttribute("anchor");
      if (anchor === "ctr") verticalAlign = "middle";
      else if (anchor === "b") verticalAlign = "bottom";

      const lIns = bodyPr.getAttribute("lIns");
      const tIns = bodyPr.getAttribute("tIns");
      const rIns = bodyPr.getAttribute("rIns");
      const bIns = bodyPr.getAttribute("bIns");
      padding = {
        left: lIns ? Math.round(Number(lIns) / 12700) : 7,
        top: tIns ? Math.round(Number(tIns) / 12700) : 4,
        right: rIns ? Math.round(Number(rIns) / 12700) : 7,
        bottom: bIns ? Math.round(Number(bIns) / 12700) : 4,
      };
    }

    const pNodes = txBody.getElementsByTagNameNS("*", "p");
    for (let i = 0; i < pNodes.length; i++) {
      const p = parseParagraph(pNodes[i], themeColors);
      paragraphs.push(p);

      if (isTitlePh) {
        const text = p.runs
          .map((r) => r.text)
          .join("")
          .trim();
        if (text) onTitleDetected(text);
      }
    }
  }

  return {
    id,
    name,
    type: paragraphs.length > 0 && !fill && !border ? "text" : "shape",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    rotation: transform.rotation,
    zIndex,
    shapeType,
    fill,
    border,
    paragraphs: paragraphs.length > 0 ? paragraphs : undefined,
    verticalAlign,
    padding,
  };
}

function parsePicElement(
  pic: Element,
  zIndex: number,
  slideRelMap: Record<string, string>,
  mediaDataUrls: Map<string, string>
): PptxElement | null {
  const nvPicPr = pic.getElementsByTagNameNS("*", "nvPicPr")[0];
  const cNvPr = nvPicPr?.getElementsByTagNameNS("*", "cNvPr")[0];
  const id = cNvPr?.getAttribute("id") || `pic_${zIndex}`;
  const name = cNvPr?.getAttribute("name") || "Image";

  const spPr = pic.getElementsByTagNameNS("*", "spPr")[0];
  const transform = parseTransform(spPr);

  const blipFill = pic.getElementsByTagNameNS("*", "blipFill")[0];
  const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0];
  const embedId = blip?.getAttribute("r:embed");

  let imageSrc = "";
  if (embedId && slideRelMap[embedId]) {
    const target = slideRelMap[embedId];
    const fileName = target.split("/").pop() || "";
    imageSrc =
      mediaDataUrls.get(target) ||
      mediaDataUrls.get(target.replace(/^ppt\//, "")) ||
      mediaDataUrls.get(fileName) ||
      "";
  }

  return {
    id,
    name,
    type: "image",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    rotation: transform.rotation,
    zIndex,
    imageSrc,
    imageAlt: name,
  };
}

function parseGraphicFrameElement(
  frame: Element,
  zIndex: number,
  themeColors: Record<string, string>
): PptxElement | null {
  const nvGraphicFramePr = frame.getElementsByTagNameNS(
    "*",
    "nvGraphicFramePr"
  )[0];
  const cNvPr = nvGraphicFramePr?.getElementsByTagNameNS("*", "cNvPr")[0];
  const id = cNvPr?.getAttribute("id") || `table_${zIndex}`;
  const name = cNvPr?.getAttribute("name") || "Table";

  const xfrm = frame.getElementsByTagNameNS("*", "xfrm")[0];
  const transform = parseTransformFromXfrm(xfrm);

  const tbl = frame.getElementsByTagNameNS("*", "tbl")[0];
  if (!tbl) return null;

  const cols: number[] = [];
  const gridCols = tbl.getElementsByTagNameNS("*", "gridCol");
  for (let i = 0; i < gridCols.length; i++) {
    const w = Number(gridCols[i].getAttribute("w")) || 0;
    cols.push(Math.round(w / 9525));
  }

  const rows: PptxTableRow[] = [];
  const trNodes = tbl.getElementsByTagNameNS("*", "tr");

  for (let i = 0; i < trNodes.length; i++) {
    const tr = trNodes[i];
    const h = Number(tr.getAttribute("h")) || 0;
    const rowHeight = Math.round(h / 9525);
    const cells: PptxTableCell[] = [];
    const tcNodes = tr.getElementsByTagNameNS("*", "tc");

    for (let j = 0; j < tcNodes.length; j++) {
      const tc = tcNodes[j];
      const gridSpan = Number(tc.getAttribute("gridSpan")) || 1;
      const rowSpan = Number(tc.getAttribute("rowSpan")) || 1;

      let backgroundColor: string | undefined = undefined;
      const solidFill = tc.getElementsByTagNameNS("*", "solidFill")[0];
      if (solidFill) {
        backgroundColor = resolveColor(solidFill, themeColors);
      }

      const cellParagraphs: PptxParagraph[] = [];
      const txBody = tc.getElementsByTagNameNS("*", "txBody")[0];
      if (txBody) {
        const pNodes = txBody.getElementsByTagNameNS("*", "p");
        for (let k = 0; k < pNodes.length; k++) {
          cellParagraphs.push(parseParagraph(pNodes[k], themeColors));
        }
      }

      const text = cellParagraphs
        .map((p) => p.runs.map((r) => r.text).join(""))
        .join("\n");

      cells.push({
        text,
        paragraphs: cellParagraphs,
        backgroundColor,
        colSpan: gridSpan > 1 ? gridSpan : undefined,
        rowSpan: rowSpan > 1 ? rowSpan : undefined,
      });
    }

    rows.push({
      height: rowHeight,
      cells,
    });
  }

  const table: PptxTable = {
    columns: cols,
    rows,
  };

  return {
    id,
    name,
    type: "table",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    rotation: transform.rotation,
    zIndex,
    table,
  };
}

function parseGroupElement(
  grpSp: Element,
  zIndex: number,
  slideRelMap: Record<string, string>,
  mediaDataUrls: Map<string, string>,
  themeColors: Record<string, string>,
  slideW: number,
  slideH: number,
  onTitleDetected: (title: string) => void
): PptxElement | null {
  const grpSpPr = grpSp.getElementsByTagNameNS("*", "grpSpPr")[0];
  const transform = parseTransform(grpSpPr);

  const children: PptxElement[] = [];
  parseShapeTree(
    grpSp,
    children,
    slideRelMap,
    mediaDataUrls,
    themeColors,
    slideW,
    slideH,
    onTitleDetected
  );

  return {
    id: `grp_${zIndex}`,
    name: "Group",
    type: "group",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    rotation: transform.rotation,
    zIndex,
    children,
  };
}

function parseConnectorElement(
  cxnSp: Element,
  zIndex: number,
  themeColors: Record<string, string>
): PptxElement | null {
  const spPr = cxnSp.getElementsByTagNameNS("*", "spPr")[0];
  const transform = parseTransform(spPr);

  const ln = spPr?.getElementsByTagNameNS("*", "ln")[0];
  let border: PptxElement["border"] = undefined;
  if (ln) {
    const w = ln.getAttribute("w");
    const widthPx = w ? Math.max(1, Math.round(Number(w) / 12700)) : 2;
    const lnFill = ln.getElementsByTagNameNS("*", "solidFill")[0];
    const clr = lnFill ? resolveColor(lnFill, themeColors) : "#000000";
    if (clr) {
      border = { color: clr, width: widthPx, style: "solid" };
    }
  }

  return {
    id: `cxn_${zIndex}`,
    name: "Line",
    type: "line",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: Math.max(transform.height, 2),
    rotation: transform.rotation,
    zIndex,
    shapeType: "line",
    border,
  };
}

function parseParagraph(
  pNode: Element,
  themeColors: Record<string, string>
): PptxParagraph {
  const pPr = pNode.getElementsByTagNameNS("*", "pPr")[0];
  let align: PptxParagraph["align"] = "left";
  let bullet: PptxParagraph["bullet"] = undefined;
  let indent: number | undefined = undefined;

  if (pPr) {
    const algn = pPr.getAttribute("algn");
    if (algn === "ctr") align = "center";
    else if (algn === "r") align = "right";
    else if (algn === "just") align = "justify";

    const marL = pPr.getAttribute("marL");
    if (marL) {
      indent = Math.round(Number(marL) / 12700);
    }

    const buChar = pPr.getElementsByTagNameNS("*", "buChar")[0];
    const buAutoNum = pPr.getElementsByTagNameNS("*", "buAutoNum")[0];
    const buNone = pPr.getElementsByTagNameNS("*", "buNone")[0];

    if (buNone) {
      bullet = { type: "none" };
    } else if (buChar) {
      bullet = {
        type: "char",
        char: buChar.getAttribute("char") || "•",
      };
    } else if (buAutoNum) {
      bullet = {
        type: "number",
        number: 1,
      };
    }
  }

  const runs: PptxRun[] = [];

  for (let i = 0; i < pNode.children.length; i++) {
    const child = pNode.children[i];
    const tag = child.localName;

    if (tag === "r") {
      const run = parseRun(child, themeColors);
      runs.push(run);
    } else if (tag === "fld") {
      const run = parseField(child, themeColors);
      runs.push(run);
    } else if (tag === "br") {
      runs.push({ text: "\n" });
    }
  }

  return {
    align,
    bullet,
    indent,
    runs,
  };
}

function parseRun(rNode: Element, themeColors: Record<string, string>): PptxRun {
  const t = rNode.getElementsByTagNameNS("*", "t")[0]?.textContent || "";
  const rPr = rNode.getElementsByTagNameNS("*", "rPr")[0];

  let fontSize: number | undefined = undefined;
  let fontFamily: string | undefined = undefined;
  let color: string | undefined = undefined;
  let bold = false;
  let italic = false;
  let underline = false;
  let strikethrough = false;
  let hyperlink: PptxRun["hyperlink"] = undefined;

  if (rPr) {
    const sz = rPr.getAttribute("sz");
    if (sz) {
      fontSize = Math.round(Number(sz) / 100);
    }

    bold = rPr.getAttribute("b") === "1";
    italic = rPr.getAttribute("i") === "1";
    underline = rPr.getAttribute("u") === "sng";
    strikethrough = rPr.getAttribute("strike") === "sngStrike";

    const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
    if (latin) {
      fontFamily = latin.getAttribute("typeface") || undefined;
    }

    const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
    if (solidFill) {
      color = resolveColor(solidFill, themeColors);
    }

    const hlinkClick = rPr.getElementsByTagNameNS("*", "hlinkClick")[0];
    if (hlinkClick) {
      const url = hlinkClick.getAttribute("action") || "";
      const tooltip = hlinkClick.getAttribute("tooltip") || undefined;
      hyperlink = { url, tooltip };
    }
  }

  return {
    text: t,
    fontSize,
    fontFamily,
    color,
    bold,
    italic,
    underline,
    strikethrough,
    hyperlink,
  };
}

function parseField(
  fldNode: Element,
  themeColors: Record<string, string>
): PptxRun {
  const fldType = fldNode.getAttribute("type");
  const isSlideNumber = fldType === "slNum";
  const t =
    fldNode.getElementsByTagNameNS("*", "t")[0]?.textContent ||
    (isSlideNumber ? "1" : "");
  const rPr = fldNode.getElementsByTagNameNS("*", "rPr")[0];

  let fontSize: number | undefined = undefined;
  let color: string | undefined = undefined;
  let bold = false;

  if (rPr) {
    const sz = rPr.getAttribute("sz");
    if (sz) fontSize = Math.round(Number(sz) / 100);
    bold = rPr.getAttribute("b") === "1";
    const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
    if (solidFill) color = resolveColor(solidFill, themeColors);
  }

  return {
    text: t,
    fontSize,
    color,
    bold,
    isSlideNumber,
  };
}

function parseTransform(node: Element | null | undefined): {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
} {
  const xfrm = node?.getElementsByTagNameNS("*", "xfrm")[0];
  return parseTransformFromXfrm(xfrm);
}

function parseTransformFromXfrm(xfrm: Element | null | undefined): {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
} {
  if (!xfrm) {
    return { x: 0, y: 0, width: 200, height: 100, rotation: 0 };
  }

  const off = xfrm.getElementsByTagNameNS("*", "off")[0];
  const ext = xfrm.getElementsByTagNameNS("*", "ext")[0];
  const rotAttr = xfrm.getAttribute("rot");

  const offX = off ? Number(off.getAttribute("x")) || 0 : 0;
  const offY = off ? Number(off.getAttribute("y")) || 0 : 0;
  const extCx = ext ? Number(ext.getAttribute("cx")) || 0 : 0;
  const extCy = ext ? Number(ext.getAttribute("cy")) || 0 : 0;

  return {
    x: Math.round(offX / 9525),
    y: Math.round(offY / 9525),
    width: Math.round(extCx / 9525),
    height: Math.round(extCy / 9525),
    rotation: rotAttr ? Number(rotAttr) / 60000 : 0,
  };
}

function resolveColor(
  container: Element,
  themeColors: Record<string, string>
): string | undefined {
  const srgb = container.getElementsByTagNameNS("*", "srgbClr")[0];
  if (srgb) {
    const val = srgb.getAttribute("val");
    if (val) return `#${val}`;
  }

  const schemeClr = container.getElementsByTagNameNS("*", "schemeClr")[0];
  if (schemeClr) {
    const val = schemeClr.getAttribute("val");
    if (val) {
      if (themeColors[val]) return themeColors[val];
      const mapping: Record<string, string> = {
        bg1: "lt1",
        bg2: "lt2",
        tx1: "dk1",
        tx2: "dk2",
      };
      if (mapping[val] && themeColors[mapping[val]]) {
        return themeColors[mapping[val]];
      }
    }
  }

  const sysClr = container.getElementsByTagNameNS("*", "sysClr")[0];
  if (sysClr) {
    const lastClr = sysClr.getAttribute("lastClr");
    if (lastClr) return `#${lastClr}`;
  }

  return undefined;
}
