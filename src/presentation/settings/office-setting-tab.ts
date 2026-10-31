/**
 * Obsidian Settings Tab for Office Suite.
 * Allows users to manage PDF Color Scheme Templates/Profiles, define semantic color tags
 * (e.g., Green = "New Info", Red = "Important Points", Yellow = "Processes"),
 * create custom templates, set the default template profile, and edit color palettes.
 */

import { App, PluginSettingTab, Setting, Notice, setIcon } from "obsidian";
import type OfficeSuitePlugin from "../../main";
import {
  PdfColorSchemeProfile,
  PdfTaggedColor,
  cloneProfile,
  normalizeHexColor,
  hexToRgb,
  ZOTERO_HEX_PALETTE,
} from "../../domain/pdf";

export class OfficeSuiteSettingTab extends PluginSettingTab {
  private plugin: OfficeSuitePlugin;
  private activeEditingProfileId: string | null = null;

  constructor(app: App, plugin: OfficeSuitePlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.addClass("office-suite-settings-root");

    containerEl.createEl("h2", { text: "Office Suite Settings" });

    // ==================== Section 1: Default Profile ====================
    new Setting(containerEl)
      .setName("Default PDF Color Scheme Template")
      .setDesc(
        "Select the default color scheme and semantic tags used when opening any PDF document."
      )
      .addDropdown((dropdown) => {
        for (const profile of this.plugin.settings.pdfProfiles) {
          dropdown.addOption(profile.id, profile.name);
        }
        dropdown.setValue(this.plugin.settings.pdfDefaultProfileId);
        dropdown.onChange(async (val) => {
          this.plugin.settings.pdfDefaultProfileId = val;
          await this.plugin.saveSettings();
          this.plugin.notifyColorSchemeChanged();
          this.display(); // re-render to update badges
        });
      });

    // ==================== Section 2: Template Profiles Manager ====================
    const profilesHeader = containerEl.createDiv({ cls: "office-settings-section-header" });
    profilesHeader.createEl("h3", { text: "PDF Color Scheme Templates & Profiles" });
    profilesHeader.createEl("p", {
      cls: "setting-item-description",
      text: "Define and customize color scheme templates with semantic tags (like Green for 'New Info', Red for 'Important Points', Yellow for 'Processes'). Profiles can be quickly imported and applied to any PDF.",
    });

    const addProfileSetting = new Setting(containerEl)
      .setName("Create New Template Profile")
      .setDesc("Create a new reusable color scheme template based on the official Zotero color palette.")
      .addButton((btn) => {
        btn
          .setButtonText("Add New Profile")
          .setCta()
          .onClick(async () => {
            const profileName = prompt("Enter a name for the new Color Scheme Template:", "Custom Research Scheme");
            if (!profileName || !profileName.trim()) return;

            const basePreset =
              this.plugin.settings.pdfProfiles.find((p) => p.id === "zotero-default") ||
              this.plugin.settings.pdfProfiles[0];
            const newProfile = cloneProfile(basePreset, undefined, profileName.trim());
            newProfile.description = "User-created color scheme template";

            this.plugin.settings.pdfProfiles.push(newProfile);
            this.activeEditingProfileId = newProfile.id;
            await this.plugin.saveSettings();
            this.plugin.notifyColorSchemeChanged();
            this.display();
            new Notice(`Created template profile "${newProfile.name}"!`);
          });
      });

    // List of Profiles
    const profilesContainer = containerEl.createDiv({ cls: "office-settings-profiles-container" });

    for (const profile of this.plugin.settings.pdfProfiles) {
      this.renderProfileCard(profilesContainer, profile);
    }
  }

  private renderProfileCard(container: HTMLElement, profile: PdfColorSchemeProfile): void {
    const isDefault = profile.id === this.plugin.settings.pdfDefaultProfileId;
    const isEditing = this.activeEditingProfileId === profile.id;

    const card = container.createDiv({
      cls: `office-profile-card ${isEditing ? "is-editing" : ""} ${isDefault ? "is-default-card" : ""}`,
    });

    // Header: Name, Badges, Actions
    const cardHeader = card.createDiv({ cls: "office-profile-header" });
    const headerLeft = cardHeader.createDiv({ cls: "office-profile-header-left" });

    headerLeft.createEl("h4", { cls: "office-profile-name", text: profile.name });

    if (isDefault) {
      headerLeft.createSpan({ cls: "office-profile-badge mod-default", text: "Default" });
    }
    if (profile.isBuiltIn) {
      headerLeft.createSpan({ cls: "office-profile-badge mod-builtin", text: "Built-In" });
    }

    const headerRight = cardHeader.createDiv({ cls: "office-profile-actions" });

    // Set Default button
    if (!isDefault) {
      const setDefaultBtn = headerRight.createEl("button", {
        cls: "office-profile-btn",
        text: "Set Default",
      });
      setDefaultBtn.onclick = async () => {
        this.plugin.settings.pdfDefaultProfileId = profile.id;
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
        this.display();
        new Notice(`"${profile.name}" is now the default PDF color scheme.`);
      };
    }

    // Duplicate button
    const dupBtn = headerRight.createEl("button", {
      cls: "office-profile-btn",
      title: "Duplicate profile",
    });
    setIcon(dupBtn, "copy");
    dupBtn.createSpan({ text: " Duplicate" });
    dupBtn.onclick = async () => {
      const copy = cloneProfile(profile, undefined, `${profile.name} (Copy)`);
      this.plugin.settings.pdfProfiles.push(copy);
      this.activeEditingProfileId = copy.id;
      await this.plugin.saveSettings();
      this.plugin.notifyColorSchemeChanged();
      this.display();
      new Notice(`Duplicated profile as "${copy.name}"`);
    };

    // Edit / Collapse button
    const editBtn = headerRight.createEl("button", {
      cls: `office-profile-btn ${isEditing ? "mod-active" : ""}`,
    });
    setIcon(editBtn, isEditing ? "chevron-up" : "pencil");
    editBtn.createSpan({ text: isEditing ? " Done" : " Edit" });
    editBtn.onclick = () => {
      this.activeEditingProfileId = isEditing ? null : profile.id;
      this.display();
    };

    // Delete button (can delete custom profiles, or non-default built-ins if cloned)
    if (!profile.isBuiltIn && !isDefault) {
      const delBtn = headerRight.createEl("button", {
        cls: "office-profile-btn mod-danger",
        title: "Delete profile",
      });
      setIcon(delBtn, "trash");
      delBtn.onclick = async () => {
        if (!confirm(`Are you sure you want to delete profile "${profile.name}"?`)) return;
        this.plugin.settings.pdfProfiles = this.plugin.settings.pdfProfiles.filter((p) => p.id !== profile.id);
        if (this.activeEditingProfileId === profile.id) {
          this.activeEditingProfileId = null;
        }
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
        this.display();
        new Notice(`Deleted profile "${profile.name}".`);
      };
    }

    if (profile.description) {
      card.createDiv({ cls: "office-profile-desc", text: profile.description });
    }

    // Color Swatches Preview with Tags
    const previewRow = card.createDiv({ cls: "office-profile-swatches-preview" });
    for (const color of profile.colors) {
      const chip = previewRow.createDiv({ cls: "office-profile-color-chip" });
      const dot = chip.createSpan({ cls: "office-chip-dot" });
      dot.style.backgroundColor = color.hex;
      chip.createSpan({ cls: "office-chip-label", text: `${color.name}: ${color.tag || "Untagged"}` });
      chip.title = `${color.name} (${color.hex}): ${color.tag || "No tag"}`;
    }

    // Expanded Edit Form
    if (isEditing) {
      this.renderProfileEditor(card, profile);
    }
  }

  private renderProfileEditor(card: HTMLElement, profile: PdfColorSchemeProfile): void {
    const editor = card.createDiv({ cls: "office-profile-editor" });

    // Name & Description Inputs
    const metaRow = editor.createDiv({ cls: "office-profile-meta-row" });
    metaRow.createSpan({ cls: "office-profile-meta-label", text: "Profile Name:" });
    const nameInput = metaRow.createEl("input", {
      type: "text",
      cls: "office-profile-name-input",
      value: profile.name,
    });
    nameInput.oninput = async () => {
      profile.name = nameInput.value;
      await this.plugin.saveSettings();
      this.plugin.notifyColorSchemeChanged();
    };

    // Color Tags Table
    const tableHeader = editor.createDiv({ cls: "office-profile-table-header" });
    tableHeader.createSpan({ cls: "col-swatch", text: "Color" });
    tableHeader.createSpan({ cls: "col-name", text: "Color Name" });
    tableHeader.createSpan({ cls: "col-tag", text: "Semantic Tag (e.g. New Info, Important, Processes)" });
    tableHeader.createSpan({ cls: "col-del", text: "" });

    const rowsWrap = editor.createDiv({ cls: "office-profile-rows-wrap" });

    profile.colors.forEach((col, idx) => {
      const row = rowsWrap.createDiv({ cls: "office-profile-color-row" });

      // Native Color Input
      const swatchWrap = row.createDiv({ cls: "col-swatch-wrap" });
      const colorInput = swatchWrap.createEl("input", {
        type: "color",
        cls: "office-color-picker",
        value: normalizeHexColor(col.hex),
      });
      colorInput.onchange = async () => {
        col.hex = normalizeHexColor(colorInput.value);
        col.rgb = hexToRgb(col.hex);
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
      };

      // Color Name input
      const nameIn = row.createEl("input", {
        type: "text",
        cls: "office-profile-input col-name-in",
        value: col.name,
        attr: { placeholder: "Color name" },
      });
      nameIn.oninput = async () => {
        col.name = nameIn.value;
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
      };

      // Semantic Tag input
      const tagIn = row.createEl("input", {
        type: "text",
        cls: "office-profile-input col-tag-in",
        value: col.tag,
        attr: { placeholder: "e.g. New Info, Important Points, Processes..." },
      });
      tagIn.oninput = async () => {
        col.tag = tagIn.value;
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
      };

      // Delete color
      const delColBtn = row.createEl("button", {
        cls: "office-icon-btn mod-danger",
        title: "Delete this color",
      });
      setIcon(delColBtn, "trash");
      delColBtn.onclick = async () => {
        if (profile.colors.length <= 1) {
          new Notice("A color profile must have at least one color.");
          return;
        }
        profile.colors.splice(idx, 1);
        await this.plugin.saveSettings();
        this.plugin.notifyColorSchemeChanged();
        this.display();
      };
    });

    // Editor actions row: Add Color, Reset to Zotero
    const actionsRow = editor.createDiv({ cls: "office-profile-editor-actions" });

    const addColorBtn = actionsRow.createEl("button", { cls: "office-profile-btn" });
    setIcon(addColorBtn, "plus");
    addColorBtn.createSpan({ text: " Add Color" });
    addColorBtn.onclick = async () => {
      profile.colors.push({
        id: `c_${Date.now()}`,
        hex: "#3b82f6",
        name: "Custom",
        tag: "General",
        rgb: [0.23, 0.51, 0.96],
      });
      await this.plugin.saveSettings();
      this.plugin.notifyColorSchemeChanged();
      this.display();
    };

    const resetColorsBtn = actionsRow.createEl("button", {
      cls: "office-profile-btn",
      text: "Reset to Standard Zotero Colors",
    });
    resetColorsBtn.onclick = async () => {
      profile.colors = [
        { id: "green", hex: ZOTERO_HEX_PALETTE.green, name: "Green", tag: "New Info", rgb: [0.37, 0.70, 0.22] },
        { id: "red", hex: ZOTERO_HEX_PALETTE.red, name: "Red", tag: "Important Points", rgb: [1.0, 0.40, 0.40] },
        { id: "yellow", hex: ZOTERO_HEX_PALETTE.yellow, name: "Yellow", tag: "Processes", rgb: [1.0, 0.83, 0.0] },
        { id: "blue", hex: ZOTERO_HEX_PALETTE.blue, name: "Blue", tag: "Definitions & Questions", rgb: [0.18, 0.66, 0.90] },
        { id: "purple", hex: ZOTERO_HEX_PALETTE.purple, name: "Purple", tag: "Theories & Claims", rgb: [0.64, 0.54, 0.90] },
        { id: "magenta", hex: ZOTERO_HEX_PALETTE.magenta, name: "Magenta", tag: "Methodology & Proofs", rgb: [0.90, 0.43, 0.93] },
        { id: "orange", hex: ZOTERO_HEX_PALETTE.orange, name: "Orange", tag: "Caveats & Objections", rgb: [0.95, 0.60, 0.22] },
        { id: "gray", hex: ZOTERO_HEX_PALETTE.gray, name: "Gray", tag: "References & Context", rgb: [0.67, 0.67, 0.67] },
      ];
      await this.plugin.saveSettings();
      this.plugin.notifyColorSchemeChanged();
      this.display();
      new Notice("Reset to standard Zotero palette!");
    };
    // ==================== Section 3: PowerPoint (PPTX) Settings ====================
    const pptxHeader = containerEl.createDiv({ cls: "office-settings-section-header" });
    pptxHeader.createEl("h3", { text: "PowerPoint (.pptx) Presentation Settings" });
    pptxHeader.createEl("p", {
      cls: "setting-item-description",
      text: "Configure default presentation view mode, slide transition effects, and viewport options.",
    });

    new Setting(containerEl)
      .setName("Default View Mode")
      .setDesc("Choose whether PPTX decks open in Single Slide view, Continuous Scroll, or Slide Sorter Grid.")
      .addDropdown((dropdown) => {
        dropdown.addOption("single", "Single Slide");
        dropdown.addOption("scroll", "Continuous Scroll");
        dropdown.addOption("grid", "Slide Sorter Grid");
        dropdown.setValue(this.plugin.settings.pptxDefaultViewMode);
        dropdown.onChange(async (val: "single" | "scroll" | "grid") => {
          this.plugin.settings.pptxDefaultViewMode = val;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName("Default Transition Effect")
      .setDesc("Default transition animation used when navigating slides (or 'Auto' to read from file).")
      .addDropdown((dropdown) => {
        dropdown.addOption("auto", "Auto (From Presentation)");
        dropdown.addOption("fade", "Fade");
        dropdown.addOption("push", "Push / Slide");
        dropdown.addOption("wipe", "Wipe");
        dropdown.addOption("zoom", "Zoom");
        dropdown.addOption("flip", "3D Flip");
        dropdown.addOption("cut", "Cut (Instant)");
        dropdown.setValue(this.plugin.settings.pptxDefaultTransition);
        dropdown.onChange(async (val) => {
          this.plugin.settings.pptxDefaultTransition = val;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName("Transition Speed")
      .setDesc("Controls the duration of slide transitions.")
      .addDropdown((dropdown) => {
        dropdown.addOption("fast", "Fast (250ms)");
        dropdown.addOption("medium", "Medium (500ms)");
        dropdown.addOption("slow", "Slow (800ms)");
        dropdown.setValue(this.plugin.settings.pptxTransitionSpeed);
        dropdown.onChange(async (val: "fast" | "medium" | "slow") => {
          this.plugin.settings.pptxTransitionSpeed = val;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName("Show Slide Number Badge")
      .setDesc("Display an overlay badge indicating the current slide number (e.g. 3 / 12) on slides.")
      .addToggle((toggle) => {
        toggle.setValue(this.plugin.settings.pptxShowSlideNumberBadge);
        toggle.onChange(async (val) => {
          this.plugin.settings.pptxShowSlideNumberBadge = val;
          await this.plugin.saveSettings();
        });
      });
  }
}
