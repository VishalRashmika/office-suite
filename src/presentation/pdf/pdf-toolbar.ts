/**
 * Comprehensive PDF Toolbar matching native PDF editors like Adobe Acrobat.
 * Includes tools (Select, Highlight, Note, Draw, Text Box, Shapes, Eraser),
 * Color presets, Zoom, Page navigation, Page editing, Citations, and Export.
 */

import { setIcon } from "obsidian";
import {
  PdfAnnotationType,
  PDF_COLOR_PRESETS,
  PdfRgbColor,
  hexToRgb,
  rgbToHex,
} from "../../domain/pdf";

export type PdfActiveTool =
  | "hand" // Default pan / text select
  | "highlight" // Text highlight
  | "underline" // Text underline
  | "note" // Sticky note
  | "draw" // Freehand ink
  | "eraser" // Erase ink/annotations
  | "freetext" // Text box directly on PDF
  | "square" // Rectangle shape
  | "circle" // Circle shape
  | "line"; // Line shape

export interface PdfToolbarDelegate {
  onToolChanged(tool: PdfActiveTool): void;
  onColorChanged(color: PdfRgbColor): void;
  onStrokeWidthChanged(width: number): void;
  onZoomChanged(scale: number | "fit-width" | "fit-page"): void;
  onPageChanged(pageNumber: number): void;
  onRotatePage(): void;
  onDuplicatePage(): void;
  onDeletePage(): void;
  onInsertBlankPage(): void;
  onToggleSidebar(): void;
  onCite(): void;
  onCopyReference(): void;
  onExportMarkdown(): void;
  onUndo(): void;
  onRedo(): void;
  onToggleTheme(): void;
  onThemeChanged?(theme: "light" | "sepia" | "dark"): void;
  onSave(): void;
}

export class PdfToolbar {
  private containerEl: HTMLElement | null = null;
  private delegate: PdfToolbarDelegate | null = null;
  private activeTool: PdfActiveTool = "hand";
  private activeColor: PdfRgbColor = PDF_COLOR_PRESETS[0].rgb;
  private strokeWidth = 2;
  private currentPage = 1;
  private totalPages = 1;
  private zoomLevel = 1.0;
  private saveStatusEl: HTMLElement | null = null;
  private pageInputEl: HTMLInputElement | null = null;
  private sidebarBtnEl: HTMLElement | null = null;
  private toolButtons: Map<PdfActiveTool, HTMLElement> = new Map();

  build(parentEl: HTMLElement, delegate: PdfToolbarDelegate): void {
    this.delegate = delegate;
    this.containerEl = parentEl.createDiv({ cls: "pdf-editor-toolbar" });

    // Group 1: Sidebar Toggle
    const groupSidebar = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });
    this.sidebarBtnEl = this.createIconButton(groupSidebar, "panel-left", "Toggle Sidebar (Thumbnails & Annotations)", () => {
      this.delegate?.onToggleSidebar();
    });
    this.setSidebarActive(true);

    this.createSeparator();

    // Group 2: Tools (Select, Highlight, Underline, Sticky Note, Text Box, Draw, Eraser, Shapes)
    const groupTools = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });

    this.addToolButton(groupTools, "hand", "mouse-pointer", "Select & Pan");
    this.addToolButton(groupTools, "highlight", "highlighter", "Highlight Text");
    this.addToolButton(groupTools, "underline", "underline", "Underline Text");
    this.addToolButton(groupTools, "note", "message-square", "Sticky Note");
    this.addToolButton(groupTools, "freetext", "type", "Text Box");
    this.addToolButton(groupTools, "draw", "pen-tool", "Draw / Ink");
    this.addToolButton(groupTools, "eraser", "eraser", "Eraser");

    this.createSeparator();

    // Group 3: Zotero Color Palette & Stroke Width
    const groupStyle = this.containerEl.createDiv({ cls: "pdf-toolbar-group pdf-toolbar-style-group" });
    const colorPickerWrap = groupStyle.createDiv({ cls: "pdf-toolbar-color-palette" });

    for (const preset of PDF_COLOR_PRESETS) {
      const swatch = colorPickerWrap.createDiv({ cls: "pdf-toolbar-swatch" });
      swatch.style.backgroundColor = preset.hex;
      swatch.title = preset.name;
      if (rgbToHex(this.activeColor).toLowerCase() === preset.hex.toLowerCase()) {
        swatch.addClass("is-active");
      }
      swatch.onclick = () => {
        this.activeColor = preset.rgb;
        colorPickerWrap.querySelectorAll(".pdf-toolbar-swatch").forEach((s) => s.removeClass("is-active"));
        swatch.addClass("is-active");
        this.delegate?.onColorChanged(this.activeColor);
      };
    }

    // Stroke width select
    const widthSelect = groupStyle.createEl("select", { cls: "pdf-toolbar-select" });
    for (const w of [1, 2, 4, 6, 8]) {
      const opt = widthSelect.createEl("option", { text: `${w}px`, value: `${w}` });
      if (w === this.strokeWidth) opt.selected = true;
    }
    widthSelect.onchange = () => {
      this.strokeWidth = parseInt(widthSelect.value, 10);
      this.delegate?.onStrokeWidthChanged(this.strokeWidth);
    };

    this.createSeparator();

    // Group 4: Page Navigation
    const groupNav = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });
    this.createIconButton(groupNav, "chevron-left", "Previous Page", () => {
      if (this.currentPage > 1) {
        this.currentPage--;
        this.updatePageIndicator();
        this.delegate?.onPageChanged(this.currentPage);
      }
    });

    const pageWrap = groupNav.createDiv({ cls: "pdf-toolbar-page-wrap" });
    this.pageInputEl = pageWrap.createEl("input", {
      cls: "pdf-toolbar-page-input",
      type: "number",
      value: "1",
      attr: { min: "1", max: `${this.totalPages}` },
    });
    this.pageInputEl.onchange = () => {
      let val = parseInt(this.pageInputEl?.value || "1", 10);
      if (isNaN(val)) val = 1;
      val = Math.max(1, Math.min(this.totalPages, val));
      this.currentPage = val;
      this.updatePageIndicator();
      this.delegate?.onPageChanged(this.currentPage);
    };
    pageWrap.createSpan({ cls: "pdf-toolbar-page-total", text: ` / ${this.totalPages}` });

    this.createIconButton(groupNav, "chevron-right", "Next Page", () => {
      if (this.currentPage < this.totalPages) {
        this.currentPage++;
        this.updatePageIndicator();
        this.delegate?.onPageChanged(this.currentPage);
      }
    });

    this.createSeparator();

    // Group 5: Zoom Controls
    const groupZoom = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });
    this.createIconButton(groupZoom, "zoom-out", "Zoom Out", () => {
      this.delegate?.onZoomChanged(Math.max(0.4, this.zoomLevel - 0.2));
    });

    const zoomSelect = groupZoom.createEl("select", { cls: "pdf-toolbar-select" });
    const zoomOptions = [
      { label: "Fit Width", value: "fit-width" },
      { label: "Fit Page", value: "fit-page" },
      { label: "50%", value: "0.5" },
      { label: "75%", value: "0.75" },
      { label: "100%", value: "1.0" },
      { label: "125%", value: "1.25" },
      { label: "150%", value: "1.5" },
      { label: "200%", value: "2.0" },
    ];
    for (const opt of zoomOptions) {
      const el = zoomSelect.createEl("option", { text: opt.label, value: opt.value });
      if (opt.value === "1.0") el.selected = true;
    }
    zoomSelect.onchange = () => {
      const val = zoomSelect.value;
      if (val === "fit-width" || val === "fit-page") {
        this.delegate?.onZoomChanged(val);
      } else {
        const num = parseFloat(val);
        this.delegate?.onZoomChanged(num);
      }
    };

    this.createIconButton(groupZoom, "zoom-in", "Zoom In", () => {
      this.delegate?.onZoomChanged(Math.min(3.0, this.zoomLevel + 0.2));
    });

    this.createSeparator();

    // Group 6: Native PDF Editor Page Operations (Rotate, Duplicate, Insert, Delete)
    const groupPageEdit = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });
    this.createIconButton(groupPageEdit, "rotate-cw", "Rotate Page 90° Clockwise", () => {
      this.delegate?.onRotatePage();
    });
    this.createIconButton(groupPageEdit, "copy", "Duplicate Current Page", () => {
      this.delegate?.onDuplicatePage();
    });
    this.createIconButton(groupPageEdit, "plus-square", "Insert Blank Page", () => {
      this.delegate?.onInsertBlankPage();
    });
    this.createIconButton(groupPageEdit, "trash", "Delete Current Page", () => {
      this.delegate?.onDeletePage();
    });

    this.createSeparator();

    // Group 7: Deep Linking, Citations & Export to Markdown
    const groupObsidian = this.containerEl.createDiv({ cls: "pdf-toolbar-group" });
    const citeBtn = this.createIconButton(groupObsidian, "quote", "Cite / Reference (APA, MLA, BibTeX, Markdown)", () => {
      this.delegate?.onCite();
    });
    const linkBtn = this.createIconButton(groupObsidian, "link", "Copy Obsidian Deep Link to Page", () => {
      this.delegate?.onCopyReference();
    });
    const exportBtn = groupObsidian.createEl("button", {
      cls: "pdf-toolbar-btn pdf-toolbar-export-btn",
      attr: { title: "Export all annotations and highlights into an Obsidian Markdown Note" },
    });
    setIcon(exportBtn, "file-output");
    exportBtn.createSpan({ text: "Export to Note" });
    exportBtn.onclick = () => {
      this.delegate?.onExportMarkdown();
    };

    this.createSeparator();

    // Group 8: Undo / Redo & Reader Theme & Save status
    const groupEnd = this.containerEl.createDiv({ cls: "pdf-toolbar-group pdf-toolbar-end-group" });
    this.createIconButton(groupEnd, "undo-2", "Undo", () => this.delegate?.onUndo());
    this.createIconButton(groupEnd, "redo-2", "Redo", () => this.delegate?.onRedo());

    // Reader Color Theme Selector (Light, Sepia, Dark)
    const themeSelect = groupEnd.createEl("select", {
      cls: "pdf-toolbar-select pdf-toolbar-theme-select",
      attr: { title: "Reader Color Theme (Light, Sepia, Dark)" },
    });
    themeSelect.createEl("option", { text: "Light", value: "light" });
    themeSelect.createEl("option", { text: "Sepia", value: "sepia" });
    themeSelect.createEl("option", { text: "Dark", value: "dark" });
    themeSelect.onchange = () => {
      const theme = themeSelect.value as "light" | "sepia" | "dark";
      this.delegate?.onThemeChanged?.(theme);
    };

    // Save Status Indicator
    this.saveStatusEl = groupEnd.createSpan({ cls: "pdf-toolbar-save-status", text: "Saved" });

    // Set default active tool
    this.setActiveTool("hand");
  }

  setSidebarActive(active: boolean): void {
    if (this.sidebarBtnEl) {
      if (active) {
        this.sidebarBtnEl.addClass("is-active");
      } else {
        this.sidebarBtnEl.removeClass("is-active");
      }
    }
  }

  private addToolButton(
    group: HTMLElement,
    tool: PdfActiveTool,
    iconId: string,
    title: string
  ): void {
    const btn = this.createIconButton(group, iconId, title, () => {
      this.setActiveTool(tool);
    });
    this.toolButtons.set(tool, btn);
  }

  setActiveTool(tool: PdfActiveTool): void {
    this.activeTool = tool;
    for (const [t, btn] of this.toolButtons.entries()) {
      if (t === tool) {
        btn.addClass("is-active");
      } else {
        btn.removeClass("is-active");
      }
    }
    this.delegate?.onToolChanged(tool);
  }

  setPageData(current: number, total: number): void {
    this.currentPage = current;
    this.totalPages = total;
    this.updatePageIndicator();
  }

  setZoom(zoom: number): void {
    this.zoomLevel = zoom;
  }

  updateSaveStatus(dirty: boolean, saving = false): void {
    if (!this.saveStatusEl) return;
    if (saving) {
      this.saveStatusEl.setText("Saving...");
      this.saveStatusEl.className = "pdf-toolbar-save-status is-saving";
    } else if (dirty) {
      this.saveStatusEl.setText("Unsaved changes");
      this.saveStatusEl.className = "pdf-toolbar-save-status is-dirty";
    } else {
      this.saveStatusEl.setText("Saved");
      this.saveStatusEl.className = "pdf-toolbar-save-status is-saved";
    }
  }

  private updatePageIndicator(): void {
    if (this.pageInputEl) {
      this.pageInputEl.value = `${this.currentPage}`;
      this.pageInputEl.max = `${this.totalPages}`;
      const totalSpan = this.containerEl?.querySelector(".pdf-toolbar-page-total");
      if (totalSpan) totalSpan.textContent = ` / ${this.totalPages}`;
    }
  }

  private createIconButton(
    parent: HTMLElement,
    iconId: string,
    title: string,
    onClick: () => void
  ): HTMLElement {
    const btn = parent.createEl("button", {
      cls: "pdf-toolbar-btn",
      title,
    });
    setIcon(btn, iconId);
    btn.onclick = (e) => {
      e.stopPropagation();
      onClick();
    };
    return btn;
  }

  private createSeparator(): void {
    this.containerEl?.createDiv({ cls: "pdf-toolbar-sep" });
  }

  dispose(): void {
    this.containerEl?.remove();
    this.containerEl = null;
    this.toolButtons.clear();
    this.delegate = null;
  }
}
