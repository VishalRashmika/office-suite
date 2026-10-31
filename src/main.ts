import { Plugin, WorkspaceLeaf, Notice, TFile } from "obsidian";
import { DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { SpreadsheetView, VIEW_TYPE_SHEET } from "./presentation/spreadsheet/spreadsheet-view";
import { PdfView, VIEW_TYPE_PDF } from "./presentation/pdf/pdf-view";
import { PptxView, VIEW_TYPE_PPTX } from "./presentation/pptx/pptx-view";
import { OfficeSuiteSettings, DEFAULT_OFFICE_SETTINGS, normalizeSettings } from "./settings/office-settings";
import { OfficeSuiteSettingTab } from "./presentation/settings/office-setting-tab";
import { PdfColorSchemeProfile, cloneProfile, ZOTERO_DEFAULT_PROFILE } from "./domain/pdf";

export default class OfficeSuitePlugin extends Plugin {
  settings: OfficeSuiteSettings = DEFAULT_OFFICE_SETTINGS;
  private colorSchemeListeners: Array<() => void> = [];

  async onload() {
    await this.loadSettings();
    this.addSettingTab(new OfficeSuiteSettingTab(this.app, this));

    // 1. Docx Editor View
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf));
    this.registerExtensions(["docx"], VIEW_TYPE_DOCX);

    // 2. Spreadsheet View (XLSX, CSV, TSV)
    this.registerView(VIEW_TYPE_SHEET, (leaf: WorkspaceLeaf) => new SpreadsheetView(leaf));
    this.registerExtensions(["xlsx"], VIEW_TYPE_SHEET);
    this.registerExtensions(["csv", "tsv"], VIEW_TYPE_SHEET);

    // 3. PDF Editor & Annotation View
    this.registerView(VIEW_TYPE_PDF, (leaf: WorkspaceLeaf) => new PdfView(leaf, this));
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

    // 4. PowerPoint Presentation View (PPTX)
    this.registerView(VIEW_TYPE_PPTX, (leaf: WorkspaceLeaf) => new PptxView(leaf, this));
    this.registerExtensions(["pptx"], VIEW_TYPE_PPTX);

    // Commands for PPTX
    this.addCommand({
      id: "pptx-start-slideshow",
      name: "PowerPoint: Start Fullscreen Slideshow (F5)",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onStartSlideshow();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pptx-next-slide",
      name: "PowerPoint: Next Slide",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onNextSlide();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pptx-prev-slide",
      name: "PowerPoint: Previous Slide",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onPrevSlide();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pptx-toggle-notes",
      name: "PowerPoint: Toggle Speaker Notes Drawer",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onToggleNotes();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pptx-toggle-sidebar",
      name: "PowerPoint: Toggle Slide Thumbnails",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onToggleSidebar();
          }
          return true;
        }
        return false;
      },
    });

    this.addCommand({
      id: "pptx-export-markdown",
      name: "PowerPoint: Export Presentation to Markdown Note",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
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
      id: "pptx-copy-slide-link",
      name: "PowerPoint: Copy Deep Link to Current Slide",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PptxView);
        if (activeLeaf) {
          if (!checking) {
            void activeLeaf.onCopySlideLink();
          }
          return true;
        }
        return false;
      },
    });

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
      id: "pdf-customize-color-scheme",
      name: "PDF: Customize Color Scheme & Tags for Current PDF",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PdfView);
        if (activeLeaf) {
          if (!checking) {
            activeLeaf.onOpenColorSchemeModal();
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
      id: "pdf-save-document",
      name: "PDF: Save Changes to PDF Document",
      checkCallback: (checking: boolean) => {
        const activeLeaf = this.app.workspace.getActiveViewOfType(PdfView);
        if (activeLeaf) {
          if (!checking) {
            void activeLeaf.requestSave();
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

  async loadSettings(): Promise<void> {
    this.settings = normalizeSettings(await this.loadData());
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  onColorSchemeChanged(listener: () => void): () => void {
    this.colorSchemeListeners.push(listener);
    return () => {
      this.colorSchemeListeners = this.colorSchemeListeners.filter((l) => l !== listener);
    };
  }

  notifyColorSchemeChanged(): void {
    for (const listener of this.colorSchemeListeners) {
      try {
        listener();
      } catch (err) {
        console.error("Error in color scheme listener:", err);
      }
    }
  }

  getColorSchemeForFile(filePath: string): PdfColorSchemeProfile {
    if (this.settings.pdfFileSchemes[filePath]) {
      return this.settings.pdfFileSchemes[filePath];
    }
    const defaultProfile =
      this.settings.pdfProfiles.find((p) => p.id === this.settings.pdfDefaultProfileId) ||
      this.settings.pdfProfiles[0] ||
      ZOTERO_DEFAULT_PROFILE;
    return cloneProfile(defaultProfile);
  }

  async setColorSchemeForFile(filePath: string, scheme: PdfColorSchemeProfile): Promise<void> {
    this.settings.pdfFileSchemes[filePath] = scheme;
    await this.saveSettings();
    this.notifyColorSchemeChanged();
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
