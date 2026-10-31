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

export class PptxDocument {
  private zip: JSZip;
  private presentation: PptxPresentation;
  private slideXmlDocs: Map<number, Document> = new Map();
  private mediaDataUrls: Map<string, string> = new Map();

  private constructor(zip: JSZip, presentation: PptxPresentation) {
    this.zip = zip;
    this.presentation = presentation;
  }

  static async load(bytes: ArrayBuffer): Promise<PptxDocument> {
    const zip = await JSZip.loadAsync(bytes);
    const parser = typeof DOMParser !== "undefined"
      ? new DOMParser()
      : new (require("@xmldom/xmldom").DOMParser)();

    // 1. Parse ppt/presentation.xml
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

    // 2. Parse Presentation Relationships (ppt/_rels/presentation.xml.rels)
    const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    const presRelMap: Record<string, string> = {};
    if (presRelsFile) {
      const presRelsXml = await presRelsFile.async("text");
      const presRelsDoc = parser.parseFromString(presRelsXml, "application/xml");
      const rels = presRelsDoc.getElementsByTagNameNS("*", "Relationship");
      for (let i = 0; i < rels.length; i++) {
        const id = rels[i].getAttribute("Id");
        const target = rels[i].getAttribute("Target");
        if (id && target) {
          presRelMap[id] = target;
        }
      }
    }

    // 3. Parse Theme (ppt/theme/theme1.xml)
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

    // 4. Identify Slide Order from <p:sldIdLst>
    const sldIdNodes = presDoc.getElementsByTagNameNS("*", "sldId");
    const slidePaths: Array<{ rId: string; path: string }> = [];

    if (sldIdNodes.length > 0) {
      for (let i = 0; i < sldIdNodes.length; i++) {
        const rId = sldIdNodes[i].getAttribute("r:id");
        if (rId && presRelMap[rId]) {
          let target = presRelMap[rId];
          if (!target.startsWith("ppt/")) {
            target = `ppt/${target.replace(/^\//, "")}`;
          }
          slidePaths.push({ rId, path: target });
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
        slidePaths.push({ rId: `slide_${i + 1}`, path: files[i] });
      }
    }

    // 5. Preload Media Files into Data URLs
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
          // Also set without 'ppt/' prefix
          mediaDataUrls.set(mf.replace(/^ppt\//, ""), dataUrl);
          mediaDataUrls.set(mf.split("/").pop()!, dataUrl);
        }
      } catch (err) {
        console.warn(`Could not preload media file ${mf}:`, err);
      }
    }

    // 6. Parse Slides
    const slides: PptxSlide[] = [];
    const slideXmlDocs = new Map<number, Document>();
    let presentationTitle = "";

    for (let i = 0; i < slidePaths.length; i++) {
      const { rId, path } = slidePaths[i];
      const slideFile = zip.file(path);
      if (!slideFile) continue;

      const slideXml = await slideFile.async("text");
      const slideDoc = parser.parseFromString(slideXml, "application/xml");
      slideXmlDocs.set(i, slideDoc);

      // Slide Relationships (ppt/slides/_rels/slide{N}.xml.rels)
      const slideDir = path.substring(0, path.lastIndexOf("/"));
      const slideFileName = path.substring(path.lastIndexOf("/") + 1);
      const relsPath = `${slideDir}/_rels/${slideFileName}.rels`;
      const relsFile = zip.file(relsPath);
      const slideRelMap: Record<string, string> = {};

      if (relsFile) {
        try {
          const relsXml = await relsFile.async("text");
          const relsDoc = parser.parseFromString(relsXml, "application/xml");
          const relNodes = relsDoc.getElementsByTagNameNS("*", "Relationship");
          for (let j = 0; j < relNodes.length; j++) {
            const relId = relNodes[j].getAttribute("Id");
            const target = relNodes[j].getAttribute("Target");
            if (relId && target) {
              slideRelMap[relId] = target;
            }
          }
        } catch (e) {
          console.warn(`Could not parse rels for ${path}:`, e);
        }
      }

      // Speaker Notes (from notesSlide)
      let notes = "";
      for (const [rKey, target] of Object.entries(slideRelMap)) {
        if (target.includes("notesSlide")) {
          let resolvedNotesPath = target.startsWith("../")
            ? `ppt/${target.substring(3)}`
            : target.startsWith("ppt/")
            ? target
            : `ppt/notesSlides/${target.split("/").pop()}`;

          const notesFile = zip.file(resolvedNotesPath);
          if (notesFile) {
            try {
              const notesXml = await notesFile.async("text");
              const notesDoc = parser.parseFromString(notesXml, "application/xml");
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
        id: rId,
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

    const doc = new PptxDocument(zip, presentation);
    doc.slideXmlDocs = slideXmlDocs;
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

  exportMarkdown(fileName: string): string {
    return PptxExport.toMarkdown(this.presentation, fileName);
  }

  async save(): Promise<Uint8Array> {
    return await this.zip.generateAsync({ type: "uint8array" });
  }
}

// -------------------------------------------------------------
// HELPER FUNCTIONS FOR OOXML PARSING
// -------------------------------------------------------------

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
    if (phType === "sldNum" || phType === "hdr" || phType === "ftr" || phType === "dt") {
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
        const pos = Number(gsNodes[i].getAttribute("pos")) / 1000; // 0 to 100
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

      // Check if this is slide title
      if (isTitlePh) {
        const text = p.runs.map((r) => r.text).join("").trim();
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
  const nvGraphicFramePr = frame.getElementsByTagNameNS("*", "nvGraphicFramePr")[0];
  const cNvPr = nvGraphicFramePr?.getElementsByTagNameNS("*", "cNvPr")[0];
  const id = cNvPr?.getAttribute("id") || `table_${zIndex}`;
  const name = cNvPr?.getAttribute("name") || "Table";

  const xfrm = frame.getElementsByTagNameNS("*", "xfrm")[0];
  const transform = parseTransformFromXfrm(xfrm);

  const tbl = frame.getElementsByTagNameNS("*", "tbl")[0];
  if (!tbl) return null;

  // Table columns
  const cols: number[] = [];
  const gridCols = tbl.getElementsByTagNameNS("*", "gridCol");
  for (let i = 0; i < gridCols.length; i++) {
    const w = Number(gridCols[i].getAttribute("w")) || 0;
    cols.push(Math.round(w / 9525));
  }

  // Table rows
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

      // Fill
      let backgroundColor: string | undefined = undefined;
      const solidFill = tc.getElementsByTagNameNS("*", "solidFill")[0];
      if (solidFill) {
        backgroundColor = resolveColor(solidFill, themeColors);
      }

      // Paragraphs
      const cellParagraphs: PptxParagraph[] = [];
      const txBody = tc.getElementsByTagNameNS("*", "txBody")[0];
      if (txBody) {
        const pNodes = txBody.getElementsByTagNameNS("*", "p");
        for (let k = 0; k < pNodes.length; k++) {
          cellParagraphs.push(parseParagraph(pNodes[k], themeColors));
        }
      }

      const text = cellParagraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");

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

function parseParagraph(pNode: Element, themeColors: Record<string, string>): PptxParagraph {
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
      // Slide number field: <a:fld id="..." type="slNum">
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
      fontSize = Math.round(Number(sz) / 100); // 100ths of pt -> pt
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

function parseField(fldNode: Element, themeColors: Record<string, string>): PptxRun {
  const fldType = fldNode.getAttribute("type");
  const isSlideNumber = fldType === "slNum";
  const t = fldNode.getElementsByTagNameNS("*", "t")[0]?.textContent || (isSlideNumber ? "1" : "");
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

function resolveColor(container: Element, themeColors: Record<string, string>): string | undefined {
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
      // Mapping for bg1, bg2, tx1, tx2
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
