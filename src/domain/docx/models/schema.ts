import { Schema, Node as PMNode } from "prosemirror-model";

/**
 * A ProseMirror schema whose nodes/marks map cleanly onto OOXML (WordprocessingML)
 * concepts, enabling round-tripping: docx -> doc JSON -> edit -> doc JSON -> docx.
 */
export const docxSchema = new Schema({
  nodes: {
    doc: { content: "block+" },

    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        align: { default: null }, // "left" | "center" | "right" | "both" | "justify"
        styleId: { default: null }, // original w:pStyle id, preserved if untouched
        numId: { default: null }, // list numbering id (w:numPr/w:numId)
        ilvl: { default: null }, // list indent level
        lineSpacing: { default: null }, // e.g. "1", "1.15", "1.5", "2"
        indent: { default: null }, // indent level (0, 1, 2, ...)
        background: { default: null }, // paragraph shading / callout background color
        border: { default: null }, // callout border
        sectPrXml: { default: null }, // preserved section properties XML for intermediate section breaks
      },
      parseDOM: [{ tag: "p" }],
      toDOM(node) {
        const styles: string[] = [];
        if (node.attrs.align) styles.push(`text-align:${node.attrs.align}`);
        if (node.attrs.lineSpacing) styles.push(`line-height:${node.attrs.lineSpacing}`);
        if (node.attrs.indent) styles.push(`margin-left:${Number(node.attrs.indent) * 36}pt`);
        if (node.attrs.background) {
          styles.push(`background-color:${node.attrs.background}`);
          styles.push("padding:8px 12px;border-radius:4px;");
        }
        if (node.attrs.border) {
          styles.push(`border-left:4px solid ${node.attrs.border}`);
        }
        return ["p", { style: styles.join(";") }, 0];
      },
    },

    heading: {
      group: "block",
      content: "inline*",
      attrs: {
        level: { default: 1 },
        align: { default: null },
        color: { default: null },
        styleId: { default: null },
        lineSpacing: { default: null },
        indent: { default: null },
        sectPrXml: { default: null },
      },
      parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({ tag: `h${level}`, attrs: { level } })),
      toDOM(node) {
        const styles: string[] = [];
        if (node.attrs.align) styles.push(`text-align:${node.attrs.align}`);
        if (node.attrs.color) styles.push(`color:${node.attrs.color}`);
        if (node.attrs.lineSpacing) styles.push(`line-height:${node.attrs.lineSpacing}`);
        if (node.attrs.indent) styles.push(`margin-left:${Number(node.attrs.indent) * 36}pt`);
        return [`h${node.attrs.level}`, { style: styles.join(";") }, 0];
      },
    },

    blockquote: {
      group: "block",
      content: "block+",
      parseDOM: [{ tag: "blockquote" }],
      toDOM() {
        return ["blockquote", 0];
      },
    },

    page_break: {
      group: "block",
      atom: true,
      selectable: true,
      parseDOM: [{ tag: "div[data-docx-page-break]" }],
      toDOM() {
        return [
          "div",
          { "data-docx-page-break": "1", class: "docx-page-break", contenteditable: "false" },
          ["span", { class: "docx-page-break-label" }, "— Page Break —"],
        ];
      },
    },

    horizontal_rule: {
      group: "block",
      atom: true,
      selectable: true,
      parseDOM: [{ tag: "hr" }],
      toDOM() {
        return ["hr"];
      },
    },

    // A list item's paragraphs keep their own numId/ilvl (mirrors how Word actually
    // stores lists as flat paragraphs with numPr, not nested list nodes). We still
    // give the editor a nested bullet/ordered_list UI via prosemirror-schema-list
    // wrapper nodes, converting between the two representations on load/save.
    bullet_list: {
      group: "block",
      content: "list_item+",
      attrs: { numId: { default: null } },
      parseDOM: [{ tag: "ul" }],
      toDOM() {
        return ["ul", 0];
      },
    },
    ordered_list: {
      group: "block",
      content: "list_item+",
      attrs: { numId: { default: null }, order: { default: 1 } },
      parseDOM: [{ tag: "ol" }],
      toDOM() {
        return ["ol", 0];
      },
    },
    list_item: {
      content: "paragraph block*",
      defining: true,
      parseDOM: [{ tag: "li" }],
      toDOM() {
        return ["li", 0];
      },
    },

    table: {
      group: "block",
      content: "table_row+",
      attrs: { align: { default: null } },
      parseDOM: [{ tag: "table" }],
      toDOM(node) {
        const styles: string[] = [];
        if (node.attrs.align === "center") styles.push("margin-left:auto;margin-right:auto");
        else if (node.attrs.align === "right") styles.push("margin-left:auto;margin-right:0");
        return ["table", { style: styles.join(";") }, ["tbody", 0]];
      },
    },
    table_row: {
      content: "table_cell+",
      parseDOM: [{ tag: "tr" }],
      toDOM() {
        return ["tr", 0];
      },
    },
    table_cell: {
      content: "paragraph+",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        background: { default: null },
        align: { default: null },
        borderColor: { default: null },
      },
      parseDOM: [{ tag: "td" }],
      toDOM(node) {
        const styles: string[] = [];
        if (node.attrs.background) styles.push(`background-color:${node.attrs.background}`);
        if (node.attrs.align) styles.push(`text-align:${node.attrs.align}`);
        if (node.attrs.borderColor) styles.push(`border-color:${node.attrs.borderColor}`);
        return [
          "td",
          {
            colspan: (node.attrs.colspan as number) || 1,
            rowspan: (node.attrs.rowspan as number) || 1,
            style: styles.join(";"),
          },
          0,
        ];
      },
    },

    // Images are kept as an atom with width, height, alignment, alt text,
    // plus the raw <w:drawing> XML and relationship id for lossless round-tripping.
    image: {
      group: "inline",
      inline: true,
      atom: true,
      draggable: true,
      selectable: true,
      attrs: {
        rId: { default: null },
        drawingXml: { default: "" },
        src: { default: "" }, // data URL for on-screen preview
        width: { default: null },
        height: { default: null },
        naturalWidth: { default: null },
        naturalHeight: { default: null },
        align: { default: "center" }, // "left" | "center" | "right" | "inline"
        alt: { default: "" },
      },
      parseDOM: [
        {
          tag: "img[data-docx-drawing]",
          getAttrs(dom) {
            if (typeof dom === "string") return false;
            const el = dom;
            return {
              src: el.getAttribute("src") || "",
              width: el.getAttribute("width") ? parseInt(el.getAttribute("width")!, 10) : null,
              height: el.getAttribute("height") ? parseInt(el.getAttribute("height")!, 10) : null,
              alt: el.getAttribute("alt") || "",
              align: el.getAttribute("data-align") || "center",
            };
          },
        },
      ],
      toDOM(node) {
        return [
          "img",
          {
            "data-docx-drawing": "1",
            src: (node.attrs.src as string) || "",
            width: (node.attrs.width as number) || undefined,
            height: (node.attrs.height as number) || undefined,
            alt: (node.attrs.alt as string) || undefined,
            "data-align": (node.attrs.align as string) || "center",
            class: `docx-image-node docx-image-align-${node.attrs.align || "center"}`,
          },
        ];
      },
    },

    text: { group: "inline" },

    hard_break: {
      group: "inline",
      inline: true,
      selectable: false,
      parseDOM: [{ tag: "br" }],
      toDOM() {
        return ["br"];
      },
    },
  },

  marks: {
    bold: {
      parseDOM: [{ tag: "strong" }, { tag: "b" }],
      toDOM() {
        return ["strong", 0];
      },
    },
    italic: {
      parseDOM: [{ tag: "em" }, { tag: "i" }],
      toDOM() {
        return ["em", 0];
      },
    },
    underline: {
      parseDOM: [{ tag: "u" }],
      toDOM() {
        return ["u", 0];
      },
    },
    strike: {
      parseDOM: [{ tag: "s" }],
      toDOM() {
        return ["s", 0];
      },
    },
    subscript: {
      parseDOM: [{ tag: "sub" }],
      toDOM() {
        return ["sub", 0];
      },
    },
    superscript: {
      parseDOM: [{ tag: "sup" }],
      toDOM() {
        return ["sup", 0];
      },
    },
    code: {
      parseDOM: [{ tag: "code" }],
      toDOM() {
        return ["code", 0];
      },
    },
    fontFamily: {
      attrs: { font: { default: "Calibri" } },
      parseDOM: [{ style: "font-family", getAttrs: (v: string) => ({ font: v.replace(/['"]/g, "") }) }],
      toDOM(mark) {
        return ["span", { style: `font-family:${mark.attrs.font}` }, 0];
      },
    },
    color: {
      attrs: { color: {} },
      parseDOM: [{ style: "color", getAttrs: (v: string) => ({ color: v }) }],
      toDOM(mark) {
        return ["span", { style: `color:${mark.attrs.color}` }, 0];
      },
    },
    highlight: {
      attrs: { color: {} },
      parseDOM: [{ style: "background-color", getAttrs: (v: string) => ({ color: v }) }],
      toDOM(mark) {
        return [
          "mark",
          { style: `background-color:${mark.attrs.color};color:inherit;padding:1px 3px;border-radius:2px;` },
          0,
        ];
      },
    },
    fontSize: {
      attrs: { pt: {} }, // half-points in OOXML (w:sz), stored here as points
      parseDOM: [{ style: "font-size", getAttrs: (v: string) => ({ pt: parseFloat(v) }) }],
      toDOM(mark) {
        return ["span", { style: `font-size:${mark.attrs.pt}pt` }, 0];
      },
    },
  },
});

export type DocxNode = PMNode;
