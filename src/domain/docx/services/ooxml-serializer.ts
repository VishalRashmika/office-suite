import { Node as PMNode, Mark } from "prosemirror-model";
import { docxSchema as schema } from "../models/schema";
import { LEVEL_TO_HEADING_STYLE } from "../models/style";
import { NS, childNS } from "../../../core/utils/xml";

export interface IDocxSerializeContext {
  documentXmlDoc: Document;
  createEl(localName: string): Element;
  setAttr(el: Element, localName: string, value: string): void;
  setXmlAttr(el: Element, localName: string, value: string): void;
  getSectPr(): Element | null;
  getOrCreateNumId?(kind: "bullet" | "ordered"): string;
  getListFormat?(numId: string): "bullet" | "ordered";
}

/**
 * Builds <w:rPr> matching OOXML ECMA-376 schema sequence:
 * 1. rStyle
 * 2. rFonts
 * 3. b
 * 4. i
 * 5. strike
 * 6. vertAlign (subscript / superscript)
 * 7. color
 * 8. sz
 * 9. highlight
 * 10. u
 */
function buildRunProps(context: IDocxSerializeContext, marks: readonly Mark[]): Element | null {
  const has = (name: string) => marks.find((m) => m.type.name === name);
  const bold = has("bold");
  const italic = has("italic");
  const underline = has("underline");
  const strike = has("strike");
  const subscript = has("subscript");
  const superscript = has("superscript");
  const code = has("code");
  const fontFamily = has("fontFamily");
  const color = has("color");
  const highlight = has("highlight");
  const fontSize = has("fontSize");

  if (
    !bold &&
    !italic &&
    !underline &&
    !strike &&
    !subscript &&
    !superscript &&
    !code &&
    !fontFamily &&
    !color &&
    !highlight &&
    !fontSize
  ) {
    return null;
  }

  const rPr = context.createEl("rPr");

  if (code) {
    const rs = context.createEl("rStyle");
    context.setAttr(rs, "val", "CodeChar");
    rPr.appendChild(rs);
  }

  if (fontFamily) {
    const rf = context.createEl("rFonts");
    const fontName = String(fontFamily.attrs.font || "Calibri");
    context.setAttr(rf, "ascii", fontName);
    context.setAttr(rf, "hAnsi", fontName);
    rPr.appendChild(rf);
  } else if (code) {
    const rf = context.createEl("rFonts");
    context.setAttr(rf, "ascii", "Courier New");
    context.setAttr(rf, "hAnsi", "Courier New");
    rPr.appendChild(rf);
  }

  if (bold) rPr.appendChild(context.createEl("b"));
  if (italic) rPr.appendChild(context.createEl("i"));
  if (strike) rPr.appendChild(context.createEl("strike"));

  if (subscript) {
    const va = context.createEl("vertAlign");
    context.setAttr(va, "val", "subscript");
    rPr.appendChild(va);
  } else if (superscript) {
    const va = context.createEl("vertAlign");
    context.setAttr(va, "val", "superscript");
    rPr.appendChild(va);
  }

  if (color) {
    const c = context.createEl("color");
    context.setAttr(c, "val", String(color.attrs.color).replace("#", ""));
    rPr.appendChild(c);
  }

  if (fontSize) {
    const sz = context.createEl("sz");
    context.setAttr(sz, "val", String(Math.round(Number(fontSize.attrs.pt) * 2)));
    rPr.appendChild(sz);
  }

  if (highlight) {
    const h = context.createEl("highlight");
    context.setAttr(h, "val", String(highlight.attrs.color));
    rPr.appendChild(h);
  }

  if (underline) {
    const u = context.createEl("u");
    context.setAttr(u, "val", "single");
    rPr.appendChild(u);
  }

  return rPr;
}

function textRun(context: IDocxSerializeContext, text: string, marks: readonly Mark[]): Element {
  const r = context.createEl("r");
  const rPr = buildRunProps(context, marks);
  if (rPr) r.appendChild(rPr);
  const t = context.createEl("t");
  context.setXmlAttr(t, "space", "preserve");
  t.textContent = text;
  r.appendChild(t);
  return r;
}

function breakRun(context: IDocxSerializeContext, marks: readonly Mark[]): Element {
  const r = context.createEl("r");
  const rPr = buildRunProps(context, marks);
  if (rPr) r.appendChild(rPr);
  r.appendChild(context.createEl("br"));
  return r;
}

function buildDrawingXml(rId: string, widthPx: number, heightPx: number, alt = "Picture"): string {
  const cx = Math.round(widthPx * 9525);
  const cy = Math.round(heightPx * 9525);
  const docPrId = Math.floor(Math.random() * 1000000) + 1;
  return `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${docPrId}" name="${alt}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${alt}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

function imageRun(context: IDocxSerializeContext, node: PMNode): Element {
  const r = context.createEl("r");
  const rId = (node.attrs.rId as string) || null;
  const width = node.attrs.width ? Number(node.attrs.width) : 300;
  const height = node.attrs.height ? Number(node.attrs.height) : 200;
  const alt = (node.attrs.alt as string) || "Picture";

  let xml = (node.attrs.drawingXml as string) || "";
  if (!xml && rId) {
    xml = buildDrawingXml(rId, width, height, alt);
  }

  if (!xml) return r;

  const nsDecls = Object.entries(NS)
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
    .join(" ");
  const xmlToParse = `<w:r ${nsDecls}>${xml}</w:r>`;
  const parser = new DOMParser();
  const parsed = parser.parseFromString(xmlToParse, "application/xml");
  if (!parsed.getElementsByTagName("parsererror").length) {
    const drawing = parsed.documentElement.firstElementChild;
    if (drawing) {
      const cx = Math.round(width * 9525);
      const cy = Math.round(height * 9525);

      const extElements = drawing.getElementsByTagNameNS(NS.wp, "extent");
      for (let i = 0; i < extElements.length; i++) {
        extElements[i].setAttribute("cx", String(cx));
        extElements[i].setAttribute("cy", String(cy));
      }

      const aExtElements = drawing.getElementsByTagNameNS(NS.a, "ext");
      for (let i = 0; i < aExtElements.length; i++) {
        aExtElements[i].setAttribute("cx", String(cx));
        aExtElements[i].setAttribute("cy", String(cy));
      }

      if (rId) {
        const blips = drawing.getElementsByTagNameNS(NS.a, "blip");
        for (let i = 0; i < blips.length; i++) {
          blips[i].setAttributeNS(NS.r, "embed", rId);
          blips[i].setAttribute("r:embed", rId);
        }
      }

      const imported = context.documentXmlDoc.importNode(drawing, true);
      r.appendChild(imported);
    }
  } else {
    console.warn("Failed to parse drawingXml in imageRun:", xml);
  }
  return r;
}

function inlineContentToRuns(context: IDocxSerializeContext, node: PMNode): Element[] {
  const runs: Element[] = [];
  node.forEach((child) => {
    if (child.type === schema.nodes.image) {
      runs.push(imageRun(context, child));
    } else if (child.type === schema.nodes.hard_break) {
      runs.push(breakRun(context, child.marks));
    } else if (child.isText) {
      runs.push(textRun(context, child.text ?? "", child.marks));
    }
  });
  return runs;
}

function buildParagraph(
  context: IDocxSerializeContext,
  node: PMNode,
  listOverride?: { numId: string; ilvl: string }
): Element {
  const p = context.createEl("p");
  const align = (node.attrs.align as string) || null;
  const styleId = (node.attrs.styleId as string) || null;
  const level = node.type === schema.nodes.heading ? (node.attrs.level as number) : null;
  const effectiveStyleId = level ? LEVEL_TO_HEADING_STYLE[level] : styleId;
  const numId = listOverride?.numId ?? ((node.attrs.numId as string) || null);
  const ilvl = listOverride?.ilvl ?? ((node.attrs.ilvl as string) || "0");
  const lineSpacing = (node.attrs.lineSpacing as string) || null;
  const indent = (node.attrs.indent as number | string) || null;
  const sectPrXml = (node.attrs.sectPrXml as string) || null;

  if (effectiveStyleId || align || numId || lineSpacing || indent || sectPrXml) {
    const pPr = context.createEl("pPr");

    if (effectiveStyleId) {
      const pStyle = context.createEl("pStyle");
      context.setAttr(pStyle, "val", effectiveStyleId);
      pPr.appendChild(pStyle);
    }

    if (numId) {
      const numPr = context.createEl("numPr");
      const ilvlEl = context.createEl("ilvl");
      context.setAttr(ilvlEl, "val", String(ilvl));
      const numIdEl = context.createEl("numId");
      context.setAttr(numIdEl, "val", String(numId));
      numPr.appendChild(ilvlEl);
      numPr.appendChild(numIdEl);
      pPr.appendChild(numPr);
    }

    if (lineSpacing) {
      const sp = context.createEl("spacing");
      const mult = parseFloat(lineSpacing) || 1;
      context.setAttr(sp, "line", String(Math.round(mult * 240)));
      context.setAttr(sp, "lineRule", "auto");
      pPr.appendChild(sp);
    }

    if (indent && Number(indent) > 0) {
      const ind = context.createEl("ind");
      context.setAttr(ind, "left", String(Number(indent) * 720));
      pPr.appendChild(ind);
    }

    if (align) {
      const jc = context.createEl("jc");
      context.setAttr(jc, "val", align === "justify" ? "both" : align);
      pPr.appendChild(jc);
    }

    if (node.attrs.background) {
      const shd = context.createEl("shd");
      context.setAttr(shd, "val", "clear");
      context.setAttr(shd, "color", "auto");
      context.setAttr(shd, "fill", String(node.attrs.background).replace("#", "").toUpperCase());
      pPr.appendChild(shd);
    }

    if (node.attrs.border) {
      const pBdr = context.createEl("pBdr");
      const left = context.createEl("left");
      context.setAttr(left, "val", "single");
      context.setAttr(left, "sz", "24");
      context.setAttr(left, "space", "15");
      context.setAttr(left, "color", String(node.attrs.border).replace("#", "").toUpperCase());
      pBdr.appendChild(left);
      pPr.appendChild(pBdr);
    }

    if (sectPrXml) {
      const nsDecls = Object.entries(NS)
        .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
        .join(" ");
      const parser = new DOMParser();
      const parsed = parser.parseFromString(`<w:root ${nsDecls}>${sectPrXml}</w:root>`, "application/xml");
      if (!parsed.getElementsByTagName("parsererror").length) {
        const sectPr = parsed.documentElement.firstElementChild;
        if (sectPr) {
          const imported = context.documentXmlDoc.importNode(sectPr, true);
          pPr.appendChild(imported);
        }
      }
    }

    p.appendChild(pPr);
  }

  for (const run of inlineContentToRuns(context, node)) p.appendChild(run);
  return p;
}

function buildTable(context: IDocxSerializeContext, node: PMNode): Element {
  const tbl = context.createEl("tbl");
  const tblPr = context.createEl("tblPr");
  const tblStyle = context.createEl("tblStyle");
  context.setAttr(tblStyle, "val", "TableGrid");
  tblPr.appendChild(tblStyle);

  if (node.attrs.align) {
    const jc = context.createEl("jc");
    context.setAttr(jc, "val", String(node.attrs.align));
    tblPr.appendChild(jc);
  }

  const tblW = context.createEl("tblW");
  context.setAttr(tblW, "w", "0");
  context.setAttr(tblW, "type", "auto");
  tblPr.appendChild(tblW);
  tbl.appendChild(tblPr);

  let maxCols = 1;
  node.forEach((row) => {
    let rowCols = 0;
    row.forEach((cell) => {
      rowCols += cell.attrs.colspan || 1;
    });
    if (rowCols > maxCols) maxCols = rowCols;
  });

  const tblGrid = context.createEl("tblGrid");
  for (let i = 0; i < maxCols; i++) {
    const gridCol = context.createEl("gridCol");
    context.setAttr(gridCol, "w", "2000");
    tblGrid.appendChild(gridCol);
  }
  tbl.appendChild(tblGrid);

  node.forEach((row) => {
    const tr = context.createEl("tr");
    row.forEach((cell) => {
      const tc = context.createEl("tc");
      const tcPr = context.createEl("tcPr");
      if (cell.attrs.colspan > 1) {
        const gridSpan = context.createEl("gridSpan");
        context.setAttr(gridSpan, "val", String(cell.attrs.colspan));
        tcPr.appendChild(gridSpan);
      }
      if (cell.attrs.background) {
        const shd = context.createEl("shd");
        context.setAttr(shd, "val", "clear");
        context.setAttr(shd, "color", "auto");
        context.setAttr(shd, "fill", String(cell.attrs.background).replace("#", "").toUpperCase());
        tcPr.appendChild(shd);
      }
      tc.appendChild(tcPr);
      if (cell.childCount === 0) {
        tc.appendChild(context.createEl("p"));
      } else {
        cell.forEach((para) => tc.appendChild(buildParagraph(context, para)));
      }
      tr.appendChild(tc);
    });
    tbl.appendChild(tr);
  });
  return tbl;
}

function buildPageBreak(context: IDocxSerializeContext): Element {
  const p = context.createEl("p");
  const r = context.createEl("r");
  const br = context.createEl("br");
  context.setAttr(br, "type", "page");
  r.appendChild(br);
  p.appendChild(r);
  return p;
}

function buildHorizontalRule(context: IDocxSerializeContext): Element {
  const p = context.createEl("p");
  const pPr = context.createEl("pPr");
  const pBdr = context.createEl("pBdr");
  const bottom = context.createEl("bottom");
  context.setAttr(bottom, "val", "single");
  context.setAttr(bottom, "sz", "6");
  context.setAttr(bottom, "space", "1");
  context.setAttr(bottom, "color", "auto");
  pBdr.appendChild(bottom);
  pPr.appendChild(pBdr);
  p.appendChild(pPr);
  return p;
}

function flattenBlock(context: IDocxSerializeContext, node: PMNode, out: Element[]) {
  if (node.type === schema.nodes.paragraph || node.type === schema.nodes.heading) {
    out.push(buildParagraph(context, node));
  } else if (node.type === schema.nodes.page_break) {
    out.push(buildPageBreak(context));
  } else if (node.type === schema.nodes.horizontal_rule) {
    out.push(buildHorizontalRule(context));
  } else if (node.type === schema.nodes.blockquote) {
    node.forEach((child) => {
      if (child.type === schema.nodes.paragraph) {
        const p = buildParagraph(context, child);
        let pPr = p.firstElementChild;
        if (!pPr || pPr.localName !== "pPr") {
          pPr = context.createEl("pPr");
          p.insertBefore(pPr, p.firstChild);
        }
        let pStyle = childNS(pPr, NS.w, "pStyle");
        if (!pStyle) {
          pStyle = context.createEl("pStyle");
          context.setAttr(pStyle, "val", "Quote");
          pPr.insertBefore(pStyle, pPr.firstChild);
        }
        let ind = childNS(pPr, NS.w, "ind");
        if (!ind) {
          ind = context.createEl("ind");
          context.setAttr(ind, "left", "720");
          pPr.appendChild(ind);
        }
        out.push(p);
      } else {
        flattenBlock(context, child, out);
      }
    });
  } else if (node.type === schema.nodes.table) {
    out.push(buildTable(context, node));
  } else if (node.type === schema.nodes.bullet_list) {
    const existingNumId = (node.attrs.numId as string) || null;
    let numId = "1";
    if (existingNumId && context.getListFormat && context.getListFormat(existingNumId) === "bullet") {
      numId = existingNumId;
    } else if (context.getOrCreateNumId) {
      numId = context.getOrCreateNumId("bullet");
    }
    node.forEach((listItem) => {
      listItem.forEach((child) => {
        if (child.type === schema.nodes.paragraph) {
          out.push(buildParagraph(context, child, { numId, ilvl: "0" }));
        } else {
          flattenBlock(context, child, out);
        }
      });
    });
  } else if (node.type === schema.nodes.ordered_list) {
    const existingNumId = (node.attrs.numId as string) || null;
    let numId = "2";
    if (existingNumId && context.getListFormat && context.getListFormat(existingNumId) === "ordered") {
      numId = existingNumId;
    } else if (context.getOrCreateNumId) {
      numId = context.getOrCreateNumId("ordered");
    }
    node.forEach((listItem) => {
      listItem.forEach((child) => {
        if (child.type === schema.nodes.paragraph) {
          out.push(buildParagraph(context, child, { numId, ilvl: "0" }));
        } else {
          flattenBlock(context, child, out);
        }
      });
    });
  }
}

export function proseMirrorDocToDocxBody(context: IDocxSerializeContext, pmDoc: PMNode): Element[] {
  const out: Element[] = [];
  pmDoc.forEach((block) => flattenBlock(context, block, out));
  const sectPr = context.getSectPr();
  if (sectPr) out.push(sectPr);
  return out;
}
