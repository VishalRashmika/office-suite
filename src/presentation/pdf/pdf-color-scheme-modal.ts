/**
 * Modal dialog for customizing the color scheme and semantic tags for an individual PDF.
 * Allows users to edit tags (e.g. Green = "New Info", Red = "Important Points", Yellow = "Processes"),
 * import colors from templates, add/remove colors, and save the scheme as a reusable profile.
 */

import { App, Modal, Notice, setIcon } from "obsidian";
import {
  PdfColorSchemeProfile,
  PdfTaggedColor,
  cloneProfile,
  normalizeHexColor,
  hexToRgb,
  ZOTERO_HEX_PALETTE,
} from "../../domain/pdf";

export interface PdfColorSchemeModalOptions {
  fileName: string;
  currentScheme: PdfColorSchemeProfile;
  availableProfiles: PdfColorSchemeProfile[];
  onApply: (scheme: PdfColorSchemeProfile) => void;
  onSaveAsTemplate: (newProfile: PdfColorSchemeProfile) => void;
}

export class PdfColorSchemeModal extends Modal {
  private options: PdfColorSchemeModalOptions;
  private editingColors: PdfTaggedColor[];
  private currentSchemeName: string;
  private colorsContainerEl: HTMLElement | null = null;

  constructor(app: App, options: PdfColorSchemeModalOptions) {
    super(app);
    this.options = options;
    this.currentSchemeName = options.currentScheme.name || "Custom Scheme";
    // Deep clone colors for editing
    this.editingColors = options.currentScheme.colors.map((c) => ({
      ...c,
      rgb: [...c.rgb],
    }));
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("pdf-color-scheme-modal");

    // Title & Context
    contentEl.createEl("h2", {
      text: "Customize Color Scheme for PDF",
      cls: "pdf-scheme-modal-title",
    });

    const subheader = contentEl.createDiv({ cls: "pdf-scheme-modal-subtitle" });
    subheader.createSpan({ text: "Document: " });
    subheader.createSpan({ cls: "pdf-scheme-file-badge", text: this.options.fileName });

    // Quick Import Row
    const importRow = contentEl.createDiv({ cls: "pdf-scheme-import-row" });
    importRow.createSpan({ cls: "pdf-scheme-row-label", text: "Import from Template:" });

    const profileSelect = importRow.createEl("select", { cls: "pdf-scheme-select" });
    profileSelect.createEl("option", { text: "-- Select a Template Profile --", value: "" });

    for (const p of this.options.availableProfiles) {
      profileSelect.createEl("option", { text: p.name, value: p.id });
    }

    const importBtn = importRow.createEl("button", { cls: "pdf-scheme-btn", text: "Import" });
    importBtn.onclick = () => {
      const selectedId = profileSelect.value;
      if (!selectedId) {
        new Notice("Please select a profile to import.");
        return;
      }
      const found = this.options.availableProfiles.find((p) => p.id === selectedId);
      if (found) {
        this.editingColors = found.colors.map((c) => ({
          ...c,
          rgb: [...c.rgb],
        }));
        this.currentSchemeName = found.name;
        this.renderColorRows();
        new Notice(`Imported colors from "${found.name}"`);
      }
    };

    // Color Rows Container
    this.colorsContainerEl = contentEl.createDiv({ cls: "pdf-scheme-colors-list" });
    this.renderColorRows();

    // Toolbar buttons: Add Color, Reset to Zotero
    const toolsRow = contentEl.createDiv({ cls: "pdf-scheme-tools-row" });

    const addColorBtn = toolsRow.createEl("button", { cls: "pdf-scheme-btn" });
    setIcon(addColorBtn, "plus");
    addColorBtn.createSpan({ text: "Add Color" });
    addColorBtn.onclick = () => {
      this.editingColors.push({
        id: `custom_${Date.now()}`,
        hex: "#3b82f6",
        name: "Custom",
        tag: "Notes",
        rgb: [0.23, 0.51, 0.96],
      });
      this.renderColorRows();
    };

    const resetBtn = toolsRow.createEl("button", { cls: "pdf-scheme-btn", text: "Reset to Zotero Default" });
    resetBtn.onclick = () => {
      const zoteroPreset = this.options.availableProfiles.find((p) => p.id === "zotero-default");
      if (zoteroPreset) {
        this.editingColors = zoteroPreset.colors.map((c) => ({ ...c, rgb: [...c.rgb] }));
        this.currentSchemeName = zoteroPreset.name;
        this.renderColorRows();
        new Notice("Reset to standard Zotero palette");
      }
    };

    // Modal Footer Actions
    const footer = contentEl.createDiv({ cls: "pdf-scheme-modal-footer" });

    const leftActions = footer.createDiv({ cls: "pdf-scheme-footer-left" });
    const saveAsTemplateBtn = leftActions.createEl("button", {
      cls: "pdf-scheme-btn",
      text: "Save as Reusable Template Profile...",
    });
    saveAsTemplateBtn.onclick = () => {
      this.promptSaveAsTemplate();
    };

    const rightActions = footer.createDiv({ cls: "pdf-scheme-footer-right" });
    const cancelBtn = rightActions.createEl("button", { cls: "pdf-scheme-btn", text: "Cancel" });
    cancelBtn.onclick = () => this.close();

    const applyBtn = rightActions.createEl("button", {
      cls: "pdf-scheme-btn mod-cta",
      text: "Apply to this PDF",
    });
    applyBtn.onclick = () => {
      this.applyChanges();
    };
  }

  private renderColorRows(): void {
    if (!this.colorsContainerEl) return;
    this.colorsContainerEl.empty();

    if (this.editingColors.length === 0) {
      this.colorsContainerEl.createDiv({
        cls: "pdf-scheme-empty",
        text: "No colors defined. Click 'Add Color' or import a template profile above.",
      });
      return;
    }

    const header = this.colorsContainerEl.createDiv({ cls: "pdf-scheme-row-header" });
    header.createSpan({ cls: "col-swatch-hdr", text: "Color" });
    header.createSpan({ cls: "col-name-hdr", text: "Color Name" });
    header.createSpan({ cls: "col-tag-hdr", text: "Semantic Tag / Meaning" });
    header.createSpan({ cls: "col-action-hdr", text: "" });

    this.editingColors.forEach((item, index) => {
      const row = this.colorsContainerEl!.createDiv({ cls: "pdf-scheme-color-row" });

      // Native Color Input + Swatch
      const swatchWrap = row.createDiv({ cls: "pdf-scheme-swatch-wrap" });
      const colorInput = swatchWrap.createEl("input", {
        type: "color",
        cls: "pdf-scheme-color-picker",
        value: normalizeHexColor(item.hex),
      });
      colorInput.onchange = () => {
        item.hex = normalizeHexColor(colorInput.value);
        item.rgb = hexToRgb(item.hex);
      };

      // Color Name input
      const nameInput = row.createEl("input", {
        type: "text",
        cls: "pdf-scheme-input col-name-input",
        value: item.name,
        attr: { placeholder: "e.g. Green" },
      });
      nameInput.oninput = () => {
        item.name = nameInput.value;
      };

      // Semantic Tag input (e.g. "New Info", "Important Points", "Processes")
      const tagInput = row.createEl("input", {
        type: "text",
        cls: "pdf-scheme-input col-tag-input",
        value: item.tag,
        attr: { placeholder: "e.g. New Info, Important Points..." },
      });
      tagInput.oninput = () => {
        item.tag = tagInput.value;
      };

      // Delete Button
      const delBtn = row.createEl("button", {
        cls: "pdf-scheme-icon-btn mod-danger",
        attr: { title: "Remove Color" },
      });
      setIcon(delBtn, "trash");
      delBtn.onclick = () => {
        if (this.editingColors.length <= 1) {
          new Notice("You must keep at least one color in the scheme.");
          return;
        }
        this.editingColors.splice(index, 1);
        this.renderColorRows();
      };
    });
  }

  private applyChanges(): void {
    if (this.editingColors.length === 0) {
      new Notice("Cannot apply empty color scheme.");
      return;
    }

    const updatedScheme: PdfColorSchemeProfile = {
      id: `file_scheme_${Date.now()}`,
      name: `Custom (${this.options.fileName})`,
      colors: this.editingColors.map((c) => ({
        id: c.id || c.hex.replace("#", ""),
        hex: normalizeHexColor(c.hex),
        name: c.name.trim() || "Color",
        tag: c.tag.trim(),
        rgb: c.rgb || hexToRgb(c.hex),
      })),
      isBuiltIn: false,
      modifiedAt: Date.now(),
    };

    this.options.onApply(updatedScheme);
    this.close();
    new Notice("Color scheme and tags applied to this PDF!");
  }

  private promptSaveAsTemplate(): void {
    const templateName = prompt("Enter a name for this new reusable template profile:", this.currentSchemeName);
    if (!templateName || !templateName.trim()) return;

    const newProfile: PdfColorSchemeProfile = {
      id: `profile_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      name: templateName.trim(),
      description: `User-defined color scheme profile created from ${this.options.fileName}`,
      colors: this.editingColors.map((c) => ({
        id: c.id || c.hex.replace("#", ""),
        hex: normalizeHexColor(c.hex),
        name: c.name.trim() || "Color",
        tag: c.tag.trim(),
        rgb: c.rgb || hexToRgb(c.hex),
      })),
      isBuiltIn: false,
      modifiedAt: Date.now(),
    };

    this.options.onSaveAsTemplate(newProfile);
    new Notice(`Saved template profile "${newProfile.name}"! Available in all PDFs.`);
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
