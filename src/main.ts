import { Plugin, WorkspaceLeaf } from "obsidian";
import { DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { SpreadsheetView, VIEW_TYPE_SHEET } from "./presentation/spreadsheet/spreadsheet-view";

export default class DocxEditorPlugin extends Plugin {
  async onload() {
    // 1. Docx Editor View
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf));
    this.registerExtensions(["docx"], VIEW_TYPE_DOCX);

    // 2. Spreadsheet View (XLSX, CSV, TSV)
    this.registerView(VIEW_TYPE_SHEET, (leaf: WorkspaceLeaf) => new SpreadsheetView(leaf));
    this.registerExtensions(["xlsx"], VIEW_TYPE_SHEET);
    this.registerExtensions(["csv", "tsv"], VIEW_TYPE_SHEET);

    // Flush any pending debounced save when Obsidian is closing/reloading.
    this.registerEvent(
      this.app.workspace.on("quit" as any, async () => {
        for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
          const view = leaf.view as DocxView;
          await view.requestSave?.();
        }
        for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_SHEET)) {
          const view = leaf.view as SpreadsheetView;
          await view.requestSave?.();
        }
      })
    );
  }

  onunload() {
    // Views clean up in onUnloadFile.
  }
}
