import { Plugin, WorkspaceLeaf, Tasks } from "obsidian";
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
      name: "Powerpoint: Start fullscreen slideshow",
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
      name: "Powerpoint: Next slide",
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
      name: "Powerpoint: Previous slide",
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
      name: "Powerpoint: Toggle speaker notes drawer",
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
      name: "Powerpoint: Toggle slide thumbnails",
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
      name: "Powerpoint: Export presentation to Markdown note",
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
      name: "Powerpoint: Copy deep link to current slide",
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
      name: "PDF: Export annotations to Markdown note",
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
      name: "PDF: Customize color scheme & tags for current PDF",
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
      name: "PDF: Copy Obsidian deep link to current page",
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
      name: "PDF: Save changes to PDF document",
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
      name: "PDF: Rotate current page 90° clockwise",
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
      this.app.workspace.on("quit", (tasks: Tasks) => {
        tasks.addPromise(
          (async () => {
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
          })()
        );
      })
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
