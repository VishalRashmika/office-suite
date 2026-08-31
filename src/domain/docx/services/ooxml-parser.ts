import { Node as PMNode, Mark } from "prosemirror-model";
import { docxSchema as schema } from "../models/schema";
import { HEADING_STYLE_TO_LEVEL, normalizeDocxColor, DocxStyleInfo } from "../models/style";
import { NS, childNS, childrenNS } from "../../../core/utils/xml";

export interface IDocxParseContext {
  body: Element;
  getMediaDataUrl(rId: string): Promise<string | null>;
  getStyle(styleId: string): DocxStyleInfo | null;
  getListFormat?(numId: string): "bullet" | "ordered";
}

function parseRunProps(rPr: Element | null): Mark[] {
  if (!rPr) return [];
  const marks: Mark[] = [];

  const b = childNS(rPr, NS.w, "b");
  if (b && b.getAttributeNS(NS.w, "val") !== "0" && b.getAttribute("w:val") !== "0") {
    marks.push(schema.marks.bold.create());
  }
  const i = childNS(rPr, NS.w, "i");
  if (i && i.getAttributeNS(NS.w, "val") !== "0" && i.getAttribute("w:val") !== "0") {
    marks.push(schema.marks.italic.create());
  }
  const u = childNS(rPr, NS.w, "u");
  const uVal = u?.getAttributeNS(NS.w, "val") ?? u?.getAttribute("w:val");
  if (u && uVal !== "none") {
    marks.push(schema.marks.underline.create());
  }
  const strike = childNS(rPr, NS.w, "strike");
  if (strike && strike.getAttributeNS(NS.w, "val") !== "0" && strike.getAttribute("w:val") !== "0") {
    marks.push(schema.marks.strike.create());
  }
  const vertAlign = childNS(rPr, NS.w, "vertAlign");
  const vertVal = vertAlign?.getAttributeNS(NS.w, "val") ?? vertAlign?.getAttribute("w:val");
  if (vertVal === "subscript") {
    marks.push(schema.marks.subscript.create());
  } else if (vertVal === "superscript") {
    marks.push(schema.marks.superscript.create());
  }
  const rFonts = childNS(rPr, NS.w, "rFonts");
  const font =
    rFonts?.getAttributeNS(NS.w, "ascii") ??
    rFonts?.getAttribute("w:ascii") ??
    rFonts?.getAttributeNS(NS.w, "hAnsi") ??
    rFonts?.getAttribute("w:hAnsi");
  if (font && font !== "Calibri") {
    marks.push(schema.marks.fontFamily.create({ font }));
  }
  const rStyle = childNS(rPr, NS.w, "rStyle");
  const rStyleVal = rStyle?.getAttributeNS(NS.w, "val") ?? rStyle?.getAttribute("w:val");
  if (rStyleVal === "VerbatimChar" || rStyleVal === "CodeChar" || rStyleVal === "SourceCode") {
    marks.push(schema.marks.code.create());
  }

  // Text Color
  const color = childNS(rPr, NS.w, "color");
  const colorVal = color?.getAttributeNS(NS.w, "val") ?? color?.getAttribute("w:val");
  const normColor = normalizeDocxColor(colorVal, color);
  if (normColor) {
    marks.push(schema.marks.color.create({ color: normColor }));
  }

  // Highlight or Run Shading
  const highlight = childNS(rPr, NS.w, "highlight");
  const hlVal = highlight?.getAttributeNS(NS.w, "val") ?? highlight?.getAttribute("w:val");
  const normHl = normalizeDocxColor(hlVal, highlight);

  const shd = childNS(rPr, NS.w, "shd");
  const shdFill =
    shd?.getAttributeNS(NS.w, "fill") ??
    shd?.getAttribute("w:fill") ??
    shd?.getAttributeNS(NS.w, "color") ??
    shd?.getAttribute("w:color");
  const normShd = normalizeDocxColor(shdFill, shd);

  const effectiveBg = normHl || normShd;
  if (effectiveBg) {
    marks.push(schema.marks.highlight.create({ color: effectiveBg }));
  }

  const sz = childNS(rPr, NS.w, "sz");
  const szVal = sz?.getAttributeNS(NS.w, "val") ?? sz?.getAttribute("w:val");
  if (szVal) {
    marks.push(schema.marks.fontSize.create({ pt: parseInt(szVal, 10) / 2 }));
  }
  return marks;
}

async function parseDrawing(drawing: Element, context: IDocxParseContext): Promise<PMNode | null> {
  const blip = drawing.getElementsByTagNameNS(NS.a, "blip")[0] ?? drawing.querySelector("blip");
  const rId = blip?.getAttributeNS(NS.r, "embed") ?? blip?.getAttribute("r:embed") ?? null;
  let src = "";
  if (rId) {
    src = (await context.getMediaDataUrl(rId)) ?? "";
  }
  const ext =
    drawing.getElementsByTagNameNS(NS.wp, "extent")[0] ??
    drawing.getElementsByTagNameNS(NS.a, "ext")[0] ??
    drawing.querySelector("extent, ext");
  const cx = ext?.getAttribute("cx");
  const cy = ext?.getAttribute("cy");
  const width = cx ? Math.round(parseInt(cx, 10) / 9525) : null;
  const height = cy ? Math.round(parseInt(cy, 10) / 9525) : null;

  const serializer = new XMLSerializer();
  return schema.nodes.image.create({
    rId,
    drawingXml: serializer.serializeToString(drawing),
    src,
    width,
    height,
  });
}

type InlineOrBreak = { kind: "inline"; node: PMNode } | { kind: "break" };

function hasPageBreakBefore(pPr: Element | null): boolean {
  if (!pPr) return false;
  const el = childNS(pPr, NS.w, "pageBreakBefore") ?? pPr.querySelector("pageBreakBefore");
  if (!el) return false;
  const val =
    el.getAttributeNS(NS.w, "val") ??
    el.getAttribute("w:val") ??
    el.getAttribute("val");
  if (val === "0" || val === "false" || val === "off" || val === "none") return false;
  return true;
}

function hasRowPageBreakBefore(tr: Element): boolean {
  const trPr = childNS(tr, NS.w, "trPr") ?? tr.querySelector("trPr");
  if (!trPr) return false;
  const pbb = childNS(trPr, NS.w, "pageBreakBefore") ?? trPr.querySelector("pageBreakBefore");
  if (pbb) {
    const val = pbb.getAttributeNS(NS.w, "val") ?? pbb.getAttribute("w:val") ?? pbb.getAttribute("val");
    if (val !== "0" && val !== "false" && val !== "off" && val !== "none") return true;
  }
  const sectPr = childNS(trPr, NS.w, "sectPr") ?? trPr.querySelector("sectPr");
  if (sectPr && isSectionBreakPageBreak(sectPr)) return true;
  return false;
}

function isSectionBreakPageBreak(sectPr: Element | null): boolean {
  if (!sectPr) return false;
  const typeEl = childNS(sectPr, NS.w, "type") ?? sectPr.querySelector("type");
  if (!typeEl) {
    // If <w:type> is omitted in OOXML, default is "nextPage"
    return true;
  }
  const val =
    typeEl.getAttributeNS(NS.w, "val") ??
    typeEl.getAttribute("w:val") ??
    typeEl.getAttribute("val") ??
    "nextPage";
  const lower = val.toLowerCase();
  return lower !== "continuous";
}

async function parseRunItems(run: Element, context: IDocxParseContext): Promise<InlineOrBreak[]> {
  const rPr = childNS(run, NS.w, "rPr");
  const marks = parseRunProps(rPr);
  const out: InlineOrBreak[] = [];

  for (const child of Array.from(run.children)) {
    const local = child.localName;
    const isW = child.namespaceURI === NS.w || !child.namespaceURI;

    if (local === "t" && isW) {
      const text = child.textContent ?? "";
      if (text.length) out.push({ kind: "inline", node: schema.text(text, marks) });
    } else if (local === "tab" && isW) {
      out.push({ kind: "inline", node: schema.text("\t", marks) });
    } else if (local === "noBreakHyphen" && isW) {
      out.push({ kind: "inline", node: schema.text("\u2011", marks) });
    } else if (local === "softHyphen" && isW) {
      out.push({ kind: "inline", node: schema.text("\u00AD", marks) });
    } else if (local === "br" && isW) {
      const brType =
        child.getAttributeNS(NS.w, "type") ??
        child.getAttribute("w:type") ??
        child.getAttribute("type");
      if (brType && (brType.toLowerCase() === "page" || brType.toLowerCase() === "column")) {
        out.push({ kind: "break" });
      } else {
        out.push({ kind: "inline", node: schema.nodes.hard_break.create() });
      }
    } else if (local === "lastRenderedPageBreak" && isW) {
      out.push({ kind: "break" });
    } else if (local === "cr" && isW) {
      out.push({ kind: "inline", node: schema.nodes.hard_break.create() });
    } else if (local === "drawing" && isW) {
      const img = await parseDrawing(child, context);
      if (img) out.push({ kind: "inline", node: img });
    } else if (local === "AlternateContent") {
      const drawing = child.getElementsByTagNameNS(NS.w, "drawing")[0] ?? child.querySelector("drawing");
      if (drawing) {
        const img = await parseDrawing(drawing, context);
        if (img) out.push({ kind: "inline", node: img });
      }
    }
  }

  return out;
}

async function collectParagraphItems(container: Element, context: IDocxParseContext): Promise<InlineOrBreak[]> {
  const items: InlineOrBreak[] = [];
  for (const child of Array.from(container.children)) {
    const local = child.localName;
    const isW = child.namespaceURI === NS.w || !child.namespaceURI;
    if (!isW && local !== "AlternateContent") continue;

    if (local === "r") {
      items.push(...(await parseRunItems(child, context)));
    } else if (local === "br") {
      const brType =
        child.getAttributeNS(NS.w, "type") ??
        child.getAttribute("w:type") ??
        child.getAttribute("type");
      if (brType && (brType.toLowerCase() === "page" || brType.toLowerCase() === "column")) {
        items.push({ kind: "break" });
      } else {
        items.push({ kind: "inline", node: schema.nodes.hard_break.create() });
      }
    } else if (local === "lastRenderedPageBreak") {
      items.push({ kind: "break" });
    } else if (local === "hyperlink") {
      items.push(...(await collectParagraphItems(child, context)));
    } else if (local === "sdt") {
      const sdtContent = childNS(child, NS.w, "sdtContent") ?? child.querySelector("sdtContent");
      if (sdtContent) {
        items.push(...(await collectParagraphItems(sdtContent, context)));
      }
    } else if (local === "ins" || local === "smartTag") {
      items.push(...(await collectParagraphItems(child, context)));
    }
  }
  return items;
}

function paragraphListInfo(pPr: Element | null): { numId: string | null; ilvl: string | null } {
  if (!pPr) return { numId: null, ilvl: null };
  const numPr = childNS(pPr, NS.w, "numPr");
  if (!numPr) return { numId: null, ilvl: null };
  const numId =
    childNS(numPr, NS.w, "numId")?.getAttributeNS(NS.w, "val") ??
    childNS(numPr, NS.w, "numId")?.getAttribute("w:val") ??
    null;
  const ilvl =
    childNS(numPr, NS.w, "ilvl")?.getAttributeNS(NS.w, "val") ??
    childNS(numPr, NS.w, "ilvl")?.getAttribute("w:val") ??
    null;
  return { numId, ilvl };
}

function parseLineSpacing(pPr: Element | null): string | null {
  if (!pPr) return null;
  const sp = childNS(pPr, NS.w, "spacing");
  const line = sp?.getAttributeNS(NS.w, "line") ?? sp?.getAttribute("w:line");
  if (!line) return null;
  const num = parseInt(line, 10);
  if (isNaN(num)) return null;
  if (Math.abs(num - 240) < 10) return "1";
  if (Math.abs(num - 276) < 10) return "1.15";
  if (Math.abs(num - 360) < 10) return "1.5";
  if (Math.abs(num - 480) < 10) return "2";
  return (num / 240).toFixed(2);
}

function parseIndent(pPr: Element | null): string | null {
  if (!pPr) return null;
  const ind = childNS(pPr, NS.w, "ind");
  const left = ind?.getAttributeNS(NS.w, "left") ?? ind?.getAttribute("w:left");
  if (!left) return null;
  const num = parseInt(left, 10);
  if (isNaN(num) || num <= 0) return null;
  return String(Math.round(num / 720));
}

async function parseParagraph(p: Element, context: IDocxParseContext): Promise<PMNode[]> {
  const pPr = childNS(p, NS.w, "pPr");
  const styleId =
    childNS(pPr ?? p, NS.w, "pStyle")?.getAttributeNS(NS.w, "val") ??
    childNS(pPr ?? p, NS.w, "pStyle")?.getAttribute("w:val") ??
    null;
  const jc =
    childNS(pPr ?? p, NS.w, "jc")?.getAttributeNS(NS.w, "val") ??
    childNS(pPr ?? p, NS.w, "jc")?.getAttribute("w:val") ??
    null;
  const align =
    jc === "both"
      ? "justify"
      : jc === "left" || jc === "center" || jc === "right" || jc === "justify"
      ? jc
      : null;
  const { numId, ilvl } = paragraphListInfo(pPr);
  const lineSpacing = parseLineSpacing(pPr);
  const indent = parseIndent(pPr);

  const sectPr = pPr ? childNS(pPr, NS.w, "sectPr") : null;
  const sectPrXml = sectPr ? new XMLSerializer().serializeToString(sectPr) : null;

  // Shading / Callout Background
  const pShd = pPr ? childNS(pPr, NS.w, "shd") : null;
  const pFill =
    pShd?.getAttributeNS(NS.w, "fill") ??
    pShd?.getAttribute("w:fill") ??
    pShd?.getAttributeNS(NS.w, "color") ??
    pShd?.getAttribute("w:color");
  const background = normalizeDocxColor(pFill, pShd);

  // Borders
  const pBdr = pPr ? childNS(pPr, NS.w, "pBdr") : null;
  let border: string | null = null;
  if (pBdr) {
    const leftBdr = childNS(pBdr, NS.w, "left");
    if (leftBdr) {
      const bdrCol = leftBdr.getAttributeNS(NS.w, "color") ?? leftBdr.getAttribute("w:color");
      border = normalizeDocxColor(bdrCol, leftBdr) || "#007acc";
    }
  }

  // Check if paragraph is just a horizontal divider
  if (pBdr && childNS(pBdr, NS.w, "bottom") && !childNS(pBdr, NS.w, "left")) {
    return [schema.nodes.horizontal_rule.create()];
  }

  const pageBreakBefore = hasPageBreakBefore(pPr);
  const items = await collectParagraphItems(p, context);
  const out: PMNode[] = [];

  if (pageBreakBefore) {
    out.push(schema.nodes.page_break.create());
  }

  const level = styleId ? HEADING_STYLE_TO_LEVEL[styleId] : undefined;
  const styleInfo = styleId ? context.getStyle(styleId) : null;
  let headingColor = styleInfo?.color || null;
  if (!headingColor && level) {
    if (level === 1 || level === 2) headingColor = "#2E74B5";
    else if (level === 3) headingColor = "#1F4D78";
    else if (level === 4) headingColor = "#2E74B5";
  }

  const createBlock = (inlines: PMNode[]) => {
    if (level) {
      return schema.nodes.heading.create(
        { level, align, color: headingColor, styleId, lineSpacing, indent, sectPrXml },
        inlines
      );
    } else if (styleId === "Quote" || styleId === "Blockquote") {
      const para = schema.nodes.paragraph.create(
        { align, styleId, numId, ilvl, lineSpacing, indent, background, border, sectPrXml },
        inlines
      );
      return schema.nodes.blockquote.create(null, [para]);
    } else {
      return schema.nodes.paragraph.create(
        { align, styleId, numId, ilvl, lineSpacing, indent, background, border, sectPrXml },
        inlines
      );
    }
  };

  let currentInlines: PMNode[] = [];
  for (const item of items) {
    if (item.kind === "break") {
      if (currentInlines.length > 0) {
        out.push(createBlock(currentInlines));
        currentInlines = [];
      }
      out.push(schema.nodes.page_break.create());
    } else {
      currentInlines.push(item.node);
    }
  }

  if (currentInlines.length > 0) {
    out.push(createBlock(currentInlines));
  }

  if (sectPr && isSectionBreakPageBreak(sectPr)) {
    out.push(schema.nodes.page_break.create());
  }

  if (out.length === 0) {
    out.push(createBlock([]));
  }

  return out;
}

async function parseTable(tbl: Element, context: IDocxParseContext): Promise<PMNode[]> {
  const tblPr = childNS(tbl, NS.w, "tblPr");
  const tblJc =
    childNS(tblPr ?? tbl, NS.w, "jc")?.getAttributeNS(NS.w, "val") ??
    childNS(tblPr ?? tbl, NS.w, "jc")?.getAttribute("w:val");
  const tblAlign = tblJc === "center" ? "center" : tblJc === "right" ? "right" : null;

  const tblBorders = tblPr ? childNS(tblPr, NS.w, "tblBorders") : null;
  let defaultTblBorder: string | null = null;
  if (tblBorders) {
    const bdr =
      childNS(tblBorders, NS.w, "bottom") ||
      childNS(tblBorders, NS.w, "insideH") ||
      childNS(tblBorders, NS.w, "top");
    const col = bdr?.getAttributeNS(NS.w, "color") ?? bdr?.getAttribute("w:color");
    defaultTblBorder = normalizeDocxColor(col, bdr);
  }

  const outBlocks: PMNode[] = [];
  let currentTableRows: PMNode[] = [];

  const flushTable = () => {
    if (currentTableRows.length > 0) {
      outBlocks.push(schema.nodes.table.create({ align: tblAlign }, currentTableRows));
      currentTableRows = [];
    }
  };

  const trElements = childrenNS(tbl, NS.w, "tr");
  for (let rIdx = 0; rIdx < trElements.length; rIdx++) {
    const tr = trElements[rIdx];

    if (hasRowPageBreakBefore(tr)) {
      flushTable();
      outBlocks.push(schema.nodes.page_break.create());
    }

    let rowHasPageBreakBefore = false;
    let rowHasPageBreakAfter = false;

    const cells: PMNode[] = [];
    const tcElements = childrenNS(tr, NS.w, "tc");

    for (const tc of tcElements) {
      const tcPr = childNS(tc, NS.w, "tcPr");
      const gridSpan =
        childNS(tcPr ?? tc, NS.w, "gridSpan")?.getAttributeNS(NS.w, "val") ??
        childNS(tcPr ?? tc, NS.w, "gridSpan")?.getAttribute("w:val");

      const tcShd = tcPr ? childNS(tcPr, NS.w, "shd") : null;
      const tcFill =
        tcShd?.getAttributeNS(NS.w, "fill") ??
        tcShd?.getAttribute("w:fill") ??
        tcShd?.getAttributeNS(NS.w, "color") ??
        tcShd?.getAttribute("w:color");
      const cellBg = normalizeDocxColor(tcFill, tcShd);

      const tcBorders = tcPr ? childNS(tcPr, NS.w, "tcBorders") : null;
      let borderColor: string | null = defaultTblBorder;
      if (tcBorders) {
        const bdr = childNS(tcBorders, NS.w, "bottom") || childNS(tcBorders, NS.w, "top");
        const col = bdr?.getAttributeNS(NS.w, "color") ?? bdr?.getAttribute("w:color");
        borderColor = normalizeDocxColor(col, bdr) || borderColor;
      }

      const tcSectPr = tcPr ? childNS(tcPr, NS.w, "sectPr") : null;
      if (tcSectPr && isSectionBreakPageBreak(tcSectPr)) {
        rowHasPageBreakAfter = true;
      }

      const cellParas: PMNode[] = [];
      const pElements = childrenNS(tc, NS.w, "p");

      for (let pIdx = 0; pIdx < pElements.length; pIdx++) {
        const p = pElements[pIdx];
        const parsed = await parseParagraph(p, context);

        for (let bIdx = 0; bIdx < parsed.length; bIdx++) {
          const block = parsed[bIdx];
          if (block.type === schema.nodes.page_break) {
            if (pIdx === 0 && cellParas.length === 0) {
              rowHasPageBreakBefore = true;
            } else {
              rowHasPageBreakAfter = true;
            }
          } else if (block.type === schema.nodes.paragraph) {
            cellParas.push(block);
          } else {
            cellParas.push(schema.nodes.paragraph.create(block.attrs, block.content));
          }
        }
      }

      if (!cellParas.length) cellParas.push(schema.nodes.paragraph.create());

      cells.push(
        schema.nodes.table_cell.create(
          {
            colspan: gridSpan ? parseInt(gridSpan, 10) : 1,
            background: cellBg,
            borderColor,
          },
          cellParas
        )
      );
    }

    if (rowHasPageBreakBefore && currentTableRows.length > 0) {
      flushTable();
      outBlocks.push(schema.nodes.page_break.create());
    }

    currentTableRows.push(schema.nodes.table_row.create(null, cells));

    if (rowHasPageBreakAfter) {
      flushTable();
      outBlocks.push(schema.nodes.page_break.create());
    }
  }

  flushTable();
  return outBlocks.length > 0 ? outBlocks : [schema.nodes.table.create({ align: tblAlign }, [])];
}

function groupLists(blocks: PMNode[], numFmtOf: (numId: string) => "bullet" | "ordered"): PMNode[] {
  const out: PMNode[] = [];
  let i = 0;
  while (i < blocks.length) {
    const b = blocks[i];
    if (b.type === schema.nodes.paragraph && b.attrs.numId) {
      const numId = String(b.attrs.numId);
      const items: PMNode[] = [];
      while (i < blocks.length && blocks[i].type === schema.nodes.paragraph && String(blocks[i].attrs.numId) === numId) {
        items.push(schema.nodes.list_item.create(null, blocks[i]));
        i++;
      }
      const kind = numFmtOf(numId);
      const listType = kind === "bullet" ? schema.nodes.bullet_list : schema.nodes.ordered_list;
      out.push(listType.create({ numId }, items));
    } else {
      out.push(b);
      i++;
    }
  }
  return out;
}

function deduplicatePageBreaks(blocks: PMNode[]): PMNode[] {
  const result: PMNode[] = [];
  let prevWasBreak = false;
  for (const block of blocks) {
    if (block.type === schema.nodes.page_break) {
      if (!prevWasBreak) {
        result.push(block);
        prevWasBreak = true;
      }
    } else {
      result.push(block);
      prevWasBreak = false;
    }
  }
  return result;
}

export async function docxToProseMirrorDoc(context: IDocxParseContext): Promise<PMNode> {
  const body = context.body;
  const blocks: PMNode[] = [];
  for (const child of Array.from(body.children)) {
    if (child.namespaceURI !== NS.w && child.namespaceURI !== null) continue;
    if (child.localName === "p") {
      blocks.push(...(await parseParagraph(child, context)));
    } else if (child.localName === "tbl") {
      blocks.push(...(await parseTable(child, context)));
    } else if (child.localName === "sdt") {
      const sdtContent = childNS(child, NS.w, "sdtContent");
      if (sdtContent) {
        for (const sdtChild of Array.from(sdtContent.children)) {
          if (sdtChild.localName === "p") {
            blocks.push(...(await parseParagraph(sdtChild, context)));
          } else if (sdtChild.localName === "tbl") {
            blocks.push(...(await parseTable(sdtChild, context)));
          }
        }
      }
    }
  }

  const grouped = groupLists(blocks, (numId) =>
    context.getListFormat ? context.getListFormat(numId) : "bullet"
  );
  const deduplicated = deduplicatePageBreaks(grouped);
  const content = deduplicated.length ? deduplicated : [schema.nodes.paragraph.create()];
  return schema.nodes.doc.create(null, content);
}
