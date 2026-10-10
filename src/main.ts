import { Plugin, WorkspaceLeaf, Notice, TFile } from "obsidian";
import { DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { SpreadsheetView, VIEW_TYPE_SHEET } from "./presentation/spreadsheet/spreadsheet-view";
import { PdfView, VIEW_TYPE_PDF } from "./presentation/pdf/pdf-view";

export default class OfficeSuitePlugin extends Plugin {
  async onload() {
    // 1. Docx Editor View
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf));
    this.registerExtensions(["docx"], VIEW_TYPE_DOCX);

    // 2. Spreadsheet View (XLSX, CSV, TSV)
    this.registerView(VIEW_TYPE_SHEET, (leaf: WorkspaceLeaf) => new SpreadsheetView(leaf));
    this.registerExtensions(["xlsx"], VIEW_TYPE_SHEET);
    this.registerExtensions(["csv", "tsv"], VIEW_TYPE_SHEET);

    // 3. PDF Editor & Annotation View
    this.registerView(VIEW_TYPE_PDF, (leaf: WorkspaceLeaf) => new PdfView(leaf));
    try {
      (this.app as any).viewRegistry?.unregisterExtensions?.(["pdf"]);
    } catch (e) {
      console.warn("Could not unregister built-in pdf extension:", e);
    }
    try {
      this.registerExtensions(["pdf"], VIEW_TYPE_PDF);
    } catch (e) {
      console.warn("Could not register pdf extension:", e);
    }

    // Commands for PDF
    this.addCommand({
      id: "pdf-export-annotations",
      name: "PDF: Export Annotations to Markdown Note",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PdfView);
        if (activeLeaf) {
          if (!checking) {
            void activeLeaf.onExportMarkdown();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pdf-copy-reference-link",
      name: "PDF: Copy Obsidian Deep Link to Current Page",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PdfView);
        if (activeLeaf) {
          if (!checking) {
            void activeLeaf.onCopyReference();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pdf-rotate-page",
      name: "PDF: Rotate Current Page 90° Clockwise",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PdfView);
        if (activeLeaf) {
          if (!checking) {
            void activeLeaf.onRotatePage();
          }
          return true;
        }
        return false;
      },
    });

    // Flush any pending debounced save when Obsidian is closing/reloading.
    this.registerEvent(
      (this.app.workspace as { on(name: string, cb: () => Promise<void>): import("obsidian").EventRef }).on(
        "quit",
        async () => {
          for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
            const view = leaf.view as DocxView;
            await view.requestSave?.();
          }
          for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_SHEET)) {
            const view = leaf.view as SpreadsheetView;
            await view.requestSave?.();
          }
          for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_PDF)) {
            const view = leaf.view as PdfView;
            await view.requestSave?.();
          }
        }
      )
    );
  }

  onunload() {
    try {
      (this.app as any).viewRegistry?.unregisterExtensions?.(["pdf"]);
      (this.app as any).viewRegistry?.registerExtensions?.(["pdf"], "pdf");
    } catch (e) {
      // Ignore cleanup error on exit
    }
  }
}
