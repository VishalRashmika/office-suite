import { EditorView } from "prosemirror-view";
import { EditorState, Transaction } from "prosemirror-state";
import { MarkType, NodeType, Node as PMNode } from "prosemirror-model";
import { wrapInList, wrapRangeInList, liftListItem } from "prosemirror-schema-list";
import { wrapIn, lift } from "prosemirror-commands";
import { docxSchema as schema } from "../../domain/docx";

import { DocxDocument } from "../../infrastructure/docx/docx-document";

export class DocxCommands {
  static toggleList(view: EditorView, listType: NodeType): boolean {
    const { state, dispatch } = view;
    const { $from, $to } = state.selection;
    const listItemType = schema.nodes.list_item;

    // Check if selection is already inside a list
    let insideListType: NodeType | null = null;
    let listPos = -1;

    for (let d = $from.depth; d > 0; d--) {
      const node = $from.node(d);
      if (node.type === schema.nodes.bullet_list || node.type === schema.nodes.ordered_list) {
        insideListType = node.type;
        listPos = $from.before(d);
        break;
      }
    }

    // 1. If already inside the same list type, lift out of list (toggle off)
    if (insideListType === listType) {
      if (liftListItem(listItemType)(state, dispatch)) {
        view.focus();
        return true;
      }
    }

    // 2. If inside a different list type, change the list type
    if (insideListType && insideListType !== listType && listPos >= 0) {
      const tr = state.tr.setNodeMarkup(listPos, listType);
      dispatch(tr);
      view.focus();
      return true;
    }

    // 3. Otherwise, convert selected lines to the target list
    const range = $from.blockRange($to);
    if (!range) {
      if (wrapInList(listType)(state, dispatch)) {
        view.focus();
        return true;
      }
      return false;
    }

    let tr = state.tr;
    let modified = false;
    const fromPos = range.start;
    const toPos = range.end;

    state.doc.nodesBetween(fromPos, toPos, (node, pos) => {
      if (node.isTextblock && node.type !== schema.nodes.paragraph) {
        tr.setBlockType(pos, pos + node.nodeSize, schema.nodes.paragraph);
        modified = true;
      }
    });

    const targetState = modified ? state.apply(tr) : state;
    const newRange = targetState.selection.$from.blockRange(targetState.selection.$to);
    if (newRange) {
      const finalTr = targetState.tr;
      if (wrapRangeInList(finalTr, newRange, listType)) {
        dispatch(finalTr.scrollIntoView());
        view.focus();
        return true;
      }
    }

    if (wrapInList(listType)(state, dispatch)) {
      view.focus();
      return true;
    }
    return false;
  }

  static toggleBlockquote(view: EditorView): boolean {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.blockquote) {
        if (lift(state, dispatch)) {
          view.focus();
          return true;
        }
      }
    }
    if (wrapIn(schema.nodes.blockquote)(state, dispatch)) {
      view.focus();
      return true;
    }
    return false;
  }
  static insertPageBreak(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
    if (dispatch) {
      const breakNode = schema.nodes.page_break.create();
      const tr = state.tr.replaceSelectionWith(breakNode).scrollIntoView();
      dispatch(tr);
    }
    return true;
  }

  static insertHorizontalRule(view: EditorView): void {
    const hr = schema.nodes.horizontal_rule.create();
    const tr = view.state.tr.replaceSelectionWith(hr).scrollIntoView();
    view.dispatch(tr);
  }

  static applyMark(view: EditorView, markType: MarkType, attrs?: Record<string, any>): void {
    const { state, dispatch } = view;
    const { from, to, empty } = state.selection;
    const mark = markType.create(attrs);
    if (empty) {
      dispatch(state.tr.addStoredMark(mark));
    } else {
      dispatch(state.tr.addMark(from, to, mark));
    }
  }

  static clearFormatting(view: EditorView): void {
    const { from, to, empty } = view.state.selection;
    if (empty) return;
    let tr = view.state.tr;
    for (const name in schema.marks) {
      tr = tr.removeMark(from, to, schema.marks[name]);
    }
    view.dispatch(tr);
  }

  static setAlignment(view: EditorView, align: "left" | "center" | "right" | "justify"): void {
    const { $from } = view.state.selection;
    const pos = $from.before($from.depth);
    const tr = view.state.tr.setNodeAttribute(pos, "align", align);
    view.dispatch(tr);
  }

  static setLineSpacing(view: EditorView, lineSpacing: string): void {
    const { $from } = view.state.selection;
    const pos = $from.before($from.depth);
    const tr = view.state.tr.setNodeAttribute(pos, "lineSpacing", lineSpacing);
    view.dispatch(tr);
  }

  static adjustIndent(view: EditorView, delta: number): void {
    const { $from } = view.state.selection;
    const pos = $from.before($from.depth);
    const node = view.state.doc.nodeAt(pos);
    const cur = Number(node?.attrs.indent || 0);
    const next = Math.max(0, cur + delta);
    view.dispatch(view.state.tr.setNodeAttribute(pos, "indent", String(next)));
  }

  static insertTable(view: EditorView, rows = 3, cols = 3): void {
    const { state, dispatch } = view;
    const tableRows: PMNode[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: PMNode[] = [];
      for (let c = 0; c < cols; c++) {
        const para = schema.nodes.paragraph.create();
        const isHeader = r === 0;
        cells.push(
          schema.nodes.table_cell.create(
            {
              colspan: 1,
              rowspan: 1,
              background: isHeader ? "#F2F4F8" : null,
            },
            [para]
          )
        );
      }
      tableRows.push(schema.nodes.table_row.create(null, cells));
    }
    const table = schema.nodes.table.create(null, tableRows);
    const tr = state.tr.replaceSelectionWith(table).scrollIntoView();
    dispatch(tr);
  }

  static addTableRow(view: EditorView, below = true): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table_row) {
        const rowNode = $from.node(d);
        const targetPos = below ? $from.after(d) : $from.before(d);
        const cells: PMNode[] = [];
        for (let c = 0; c < rowNode.childCount; c++) {
          cells.push(schema.nodes.table_cell.create({ colspan: 1, rowspan: 1 }, [schema.nodes.paragraph.create()]));
        }
        const newRow = schema.nodes.table_row.create(null, cells);
        dispatch(state.tr.insert(targetPos, newRow));
        return;
      }
    }
  }

  static addTableColumn(view: EditorView, right = true): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    let tableDepth = -1;
    let cellIndex = -1;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table_cell) {
        cellIndex = $from.index(d - 1);
      }
      if ($from.node(d).type === schema.nodes.table) {
        tableDepth = d;
        break;
      }
    }
    if (tableDepth === -1 || cellIndex === -1) return;

    const targetIndex = right ? cellIndex + 1 : cellIndex;
    const tableNode = $from.node(tableDepth);
    const tablePos = $from.before(tableDepth);

    const newRows: PMNode[] = [];
    tableNode.forEach((row) => {
      const cells: PMNode[] = [];
      let i = 0;
      row.forEach((cell) => {
        if (i === targetIndex && !right) {
          cells.push(schema.nodes.table_cell.create({ colspan: 1, rowspan: 1 }, [schema.nodes.paragraph.create()]));
        }
        cells.push(cell);
        if (i === targetIndex - 1 && right) {
          cells.push(schema.nodes.table_cell.create({ colspan: 1, rowspan: 1 }, [schema.nodes.paragraph.create()]));
        }
        i++;
      });
      if (targetIndex >= row.childCount && right) {
        cells.push(schema.nodes.table_cell.create({ colspan: 1, rowspan: 1 }, [schema.nodes.paragraph.create()]));
      }
      newRows.push(schema.nodes.table_row.create(null, cells));
    });

    const newTable = schema.nodes.table.create(tableNode.attrs, newRows);
    dispatch(state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, newTable));
  }

  static deleteTableRow(view: EditorView): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table_row) {
        const from = $from.before(d);
        const to = $from.after(d);
        dispatch(state.tr.delete(from, to));
        return;
      }
    }
  }

  static deleteTableColumn(view: EditorView): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    let tableDepth = -1;
    let cellIndex = -1;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table_cell) {
        cellIndex = $from.index(d - 1);
      }
      if ($from.node(d).type === schema.nodes.table) {
        tableDepth = d;
        break;
      }
    }
    if (tableDepth === -1 || cellIndex === -1) return;

    const tableNode = $from.node(tableDepth);
    const tablePos = $from.before(tableDepth);

    const newRows: PMNode[] = [];
    tableNode.forEach((row) => {
      if (row.childCount <= 1) return;
      const cells: PMNode[] = [];
      let i = 0;
      row.forEach((cell) => {
        if (i !== cellIndex) cells.push(cell);
        i++;
      });
      newRows.push(schema.nodes.table_row.create(null, cells));
    });

    if (newRows.length === 0) {
      dispatch(state.tr.delete(tablePos, tablePos + tableNode.nodeSize));
    } else {
      const newTable = schema.nodes.table.create(tableNode.attrs, newRows);
      dispatch(state.tr.replaceWith(tablePos, tablePos + tableNode.nodeSize, newTable));
    }
  }

  static deleteTable(view: EditorView): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table) {
        const from = $from.before(d);
        const to = $from.after(d);
        dispatch(state.tr.delete(from, to));
        return;
      }
    }
  }

  static setCellBackground(view: EditorView, color: string): void {
    const { state, dispatch } = view;
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type === schema.nodes.table_cell) {
        const pos = $from.before(d);
        dispatch(state.tr.setNodeAttribute(pos, "background", color || null));
        return;
      }
    }
  }

  static async insertImageFile(view: EditorView, docxDoc: DocxDocument | null, file: File): Promise<void> {
    if (!docxDoc) return;
    const buffer = await file.arrayBuffer();
    const ext = file.name.split(".").pop() || "png";
    const { rId, dataUrl } = await docxDoc.addMediaFile(buffer, ext);

    const tempImg = new Image();
    tempImg.onload = () => {
      const nw = tempImg.naturalWidth || 400;
      const nh = tempImg.naturalHeight || 300;
      const targetWidth = Math.min(nw, 640);
      const targetHeight = Math.round((targetWidth / nw) * nh);

      const imgNode = schema.nodes.image.create({
        rId,
        src: dataUrl,
        drawingXml: "",
        width: targetWidth,
        height: targetHeight,
        naturalWidth: nw,
        naturalHeight: nh,
        align: "center",
      });

      const tr = view.state.tr.replaceSelectionWith(imgNode).scrollIntoView();
      view.dispatch(tr);
      view.focus();
    };
    tempImg.src = dataUrl;
  }

  static promptInsertImage(view: EditorView, docxDoc: DocxDocument | null): void {
    if (!docxDoc) return;
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.style.display = "none";
    input.onchange = () => {
      const file = input.files?.[0];
      if (file) {
        DocxCommands.insertImageFile(view, docxDoc, file);
      }
    };
    input.click();
  }
}
