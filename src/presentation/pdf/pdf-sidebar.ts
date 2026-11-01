/**
 * Zotero-style Collapsible PDF Left Side Panel.
 * Contains:
 * 1. Annotations Tab:
 *    - Full-text search with page number matching
 *    - Sort by Color (grouped by semantic tags like Green = "New Info", Red = "Important Points")
 *    - Sort by Page (document reading order)
 *    - Sort by Date (newest / oldest)
 *    - Sort by Annotation Type
 *    - Color Palette quick filter swatches with tag labels and counts
 *    - Type filter chips (Highlights, Underlines, Sticky Notes, Text Boxes, Drawings, Shapes)
 *    - 1:1 Zotero card presentation: vertical left color bar, semantic tag badge, clickable page link,
 *      underlined/highlighted quotes, attached comments, Obsidian linked note links, quick color changer
 *    - 1-click Export to Obsidian Markdown note
 * 2. Outline Tab:
 *    - Hierarchical PDF Table of Contents / Bookmarks extracted via PDF.js with jump-to-page navigation
 * 3. Page Thumbnails Tab:
 *    - High-fidelity mini page previews, active page indicator, page rotation/duplicate/deletion
 */

import { setIcon, Notice } from "obsidian";
import {
  PdfAnnotation,
  PdfAnnotationType,
  PdfRgbColor,
  CitationBuilder,
  PdfColorSchemeProfile,
  PDF_COLOR_PRESETS,
  rgbToHex,
  hexToRgb,
  normalizeHexColor,
  sortAnnotations,
  groupAnnotationsByColor,
  groupAnnotationsByType,
  filterAnnotations,
  getAnnotationColorInfo,
  formatAnnotationDate,
  formatTypeLabel,
  AnnotSortField,
  AnnotSortDirection,
} from "../../domain/pdf";
import { PdfDocumentAdapter } from "../../infrastructure/pdf";

export interface PdfSidebarDelegate {
  onSelectPage(pageIndex: number): void;
  onSelectAnnotation(annotation: PdfAnnotation): void;
  onEditAnnotation(annotation: PdfAnnotation): void;
  onDeleteAnnotation(id: string): void;
  onRotatePage(pageIndex: number): void;
  onDuplicatePage(pageIndex: number): void;
  onDeletePage(pageIndex: number): void;
  onExportMarkdown?(): void;
  onChangeAnnotationColor?(id: string, color: PdfRgbColor): void;
  onOpenLinkedNote?(vaultPath: string): void;
}

export type SidebarTab = "thumbnails" | "outline" | "annotations";

export class PdfSidebar {
  private containerEl: HTMLElement | null = null;
  private contentEl: HTMLElement | null = null;
  private activeTab: SidebarTab = "annotations";

  // Sorting & Filtering State (cloning Zotero)
  private sortField: AnnotSortField = "page";
  private sortDirection: AnnotSortDirection = "asc";
  private filterType: "all" | PdfAnnotationType | "shape" = "all";
  private filterColorHex = "all";
  private filterTag = "all";
  private searchQuery = "";
  private selectedAnnotId: string | null = null;
  private collapsedColorGroups: Set<string> = new Set();
  private collapsedTypeGroups: Set<string> = new Set();
  private currentPageIndex = 0;

  private delegate: PdfSidebarDelegate | null = null;
  private pdfAdapter: PdfDocumentAdapter | null = null;
  private pdfjsDoc: any = null;
  private fileName = "";
  private isVisible = false;
  private colorScheme?: PdfColorSchemeProfile;

  // Header UI elements
  private annotTabBtn: HTMLElement | null = null;
  private outlineTabBtn: HTMLElement | null = null;
  private thumbTabBtn: HTMLElement | null = null;

  public onVisibilityChange?: (visible: boolean) => void;

  setColorScheme(scheme: PdfColorSchemeProfile): void {
    this.colorScheme = scheme;
    if (this.activeTab === "annotations" && this.contentEl) {
      this.renderActiveTab();
    }
  }

  setCurrentPage(pageIndex: number): void {
    this.currentPageIndex = pageIndex;
    if (this.activeTab === "thumbnails" && this.contentEl) {
      this.updateThumbnailActiveState();
    }
  }

  getIsVisible(): boolean {
    return this.isVisible;
  }

  build(
    parentEl: HTMLElement,
    pdfAdapter: PdfDocumentAdapter,
    pdfjsDoc: any,
    fileName: string,
    delegate: PdfSidebarDelegate,
    colorScheme?: PdfColorSchemeProfile,
    initialVisible = false
  ): void {
    this.pdfAdapter = pdfAdapter;
    this.pdfjsDoc = pdfjsDoc;
    this.fileName = fileName;
    this.delegate = delegate;
    this.colorScheme = colorScheme;
    this.isVisible = initialVisible;

    this.containerEl = parentEl.createDiv({
      cls: `pdf-sidebar-container${initialVisible ? "" : " is-hidden"}`,
    });
    if (!initialVisible) {
      this.containerEl.style.display = "none";
    }

    // 1. Resizer drag handle
    const resizer = this.containerEl.createDiv({ cls: "pdf-sidebar-resizer" });
    let startX = 0;
    let startWidth = 0;

    resizer.addEventListener("pointerdown", (e: PointerEvent) => {
      e.preventDefault();
      startX = e.clientX;
      startWidth = this.containerEl!.getBoundingClientRect().width;

      const onPointerMove = (me: PointerEvent) => {
        const newW = Math.max(260, Math.min(600, startWidth + (me.clientX - startX)));
        if (this.containerEl) {
          this.containerEl.style.width = `${newW}px`;
        }
      };

      const onPointerUp = () => {
        window.removeEventListener("pointermove", onPointerMove);
        window.removeEventListener("pointerup", onPointerUp);
      };

      window.addEventListener("pointermove", onPointerMove);
      window.addEventListener("pointerup", onPointerUp);
    });

    // 2. Sidebar Header with 3 Zotero Tabs & Collapse button
    const headerEl = this.containerEl.createDiv({ cls: "pdf-sidebar-header" });
    const tabsEl = headerEl.createDiv({ cls: "pdf-sidebar-tabs" });

    // Tab 1: Thumbnails (Pages)
    this.thumbTabBtn = tabsEl.createEl("button", {
      cls: `pdf-sidebar-tab ${this.activeTab === "thumbnails" ? "is-active" : ""}`,
      attr: { title: "Page Thumbnails" },
    });
    setIcon(this.thumbTabBtn, "layout-grid");
    this.thumbTabBtn.createSpan({ cls: "pdf-sidebar-tab-label", text: "Pages" });
    this.thumbTabBtn.onclick = () => this.switchTab("thumbnails");

    // Tab 2: Outline (Table of Contents)
    this.outlineTabBtn = tabsEl.createEl("button", {
      cls: `pdf-sidebar-tab ${this.activeTab === "outline" ? "is-active" : ""}`,
      attr: { title: "Document Outline / Table of Contents" },
    });
    setIcon(this.outlineTabBtn, "list-tree");
    this.outlineTabBtn.createSpan({ cls: "pdf-sidebar-tab-label", text: "Outline" });
    this.outlineTabBtn.onclick = () => this.switchTab("outline");

    // Tab 3: Annotations (Zotero core feature)
    const annotCount = pdfAdapter.getAllAnnotations().length;
    this.annotTabBtn = tabsEl.createEl("button", {
      cls: `pdf-sidebar-tab ${this.activeTab === "annotations" ? "is-active" : ""}`,
      attr: { title: "Annotations & Highlights List" },
    });
    setIcon(this.annotTabBtn, "highlighter");
    this.annotTabBtn.createSpan({
      cls: "pdf-sidebar-tab-label",
      text: annotCount > 0 ? `Annotations (${annotCount})` : "Annotations",
    });
    this.annotTabBtn.onclick = () => this.switchTab("annotations");

    // Collapse Button
    const closeBtn = headerEl.createEl("button", {
      cls: "pdf-sidebar-close-btn",
      attr: { title: "Collapse side panel" },
    });
    setIcon(closeBtn, "chevrons-left");
    closeBtn.onclick = () => this.toggle();

    // 3. Content container
    this.contentEl = this.containerEl.createDiv({ cls: "pdf-sidebar-content" });

    this.renderActiveTab();
  }

  toggle(): boolean {
    if (!this.containerEl) return false;
    this.isVisible = !this.isVisible;
    if (this.isVisible) {
      this.containerEl.style.display = "";
      this.containerEl.removeClass("is-hidden");
    } else {
      this.containerEl.addClass("is-hidden");
      this.containerEl.style.display = "none";
    }
    this.onVisibilityChange?.(this.isVisible);
    return this.isVisible;
  }

  show(): void {
    this.isVisible = true;
    if (this.containerEl) {
      this.containerEl.style.display = "";
      this.containerEl.removeClass("is-hidden");
    }
    this.onVisibilityChange?.(true);
  }

  hide(): void {
    this.isVisible = false;
    if (this.containerEl) {
      this.containerEl.addClass("is-hidden");
      this.containerEl.style.display = "none";
    }
    this.onVisibilityChange?.(false);
  }

  switchTab(tab: SidebarTab): void {
    this.activeTab = tab;
    if (!this.containerEl) return;

    this.thumbTabBtn?.toggleClass("is-active", tab === "thumbnails");
    this.outlineTabBtn?.toggleClass("is-active", tab === "outline");
    this.annotTabBtn?.toggleClass("is-active", tab === "annotations");

    this.renderActiveTab();
  }

  updateData(pdfAdapter: PdfDocumentAdapter, pdfjsDoc: any): void {
    this.pdfAdapter = pdfAdapter;
    this.pdfjsDoc = pdfjsDoc;

    if (this.annotTabBtn) {
      const count = pdfAdapter.getAllAnnotations().length;
      const label = this.annotTabBtn.querySelector(".pdf-sidebar-tab-label");
      if (label) {
        label.setText(count > 0 ? `Annotations (${count})` : "Annotations");
      }
    }

    this.renderActiveTab();
  }

  private renderActiveTab(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    this.contentEl.empty();

    if (this.activeTab === "thumbnails") {
      this.renderThumbnails();
    } else if (this.activeTab === "outline") {
      this.renderOutline();
    } else {
      this.renderAnnotationsView();
    }
  }

  // =========================================================================
  // ==================== 1. ANNOTATIONS TAB (ZOTERO 1:1) ====================
  // =========================================================================

  private renderAnnotationsView(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    const viewWrap = this.contentEl.createDiv({ cls: "pdf-annot-view-wrap" });

    // 1. Search Bar with Instant Clear Button
    const searchRow = viewWrap.createDiv({ cls: "pdf-annot-search-row" });
    const searchIconWrap = searchRow.createSpan({ cls: "pdf-annot-search-icon" });
    setIcon(searchIconWrap, "search");

    const searchInput = searchRow.createEl("input", {
      cls: "pdf-annot-search-input",
      type: "text",
      value: this.searchQuery,
      attr: { placeholder: "Search annotations, quotes, tags..." },
    });

    const clearSearchBtn = searchRow.createSpan({
      cls: `pdf-annot-search-clear ${this.searchQuery ? "" : "is-hidden"}`,
      attr: { title: "Clear search" },
    });
    setIcon(clearSearchBtn, "x");
    clearSearchBtn.onclick = () => {
      this.searchQuery = "";
      searchInput.value = "";
      clearSearchBtn.addClass("is-hidden");
      this.renderAnnotationCardsList(cardsContainer, countLabel);
    };

    searchInput.oninput = () => {
      this.searchQuery = searchInput.value.trim();
      clearSearchBtn.toggleClass("is-hidden", !this.searchQuery);
      this.renderAnnotationCardsList(cardsContainer, countLabel);
    };

    // 2. Sorting & Grouping Controls Bar (Key feature!)
    const sortRow = viewWrap.createDiv({ cls: "pdf-annot-sort-row" });

    const sortLeft = sortRow.createDiv({ cls: "pdf-annot-sort-left" });
    const sortIcon = sortLeft.createSpan({ cls: "pdf-annot-sort-icon" });
    setIcon(sortIcon, "arrow-up-down");
    sortLeft.createSpan({ cls: "pdf-annot-sort-label", text: "Sort:" });

    const sortSelect = sortLeft.createEl("select", { cls: "pdf-annot-sort-select" });
    const sortOptions: { field: AnnotSortField; label: string }[] = [
      { field: "page", label: "Page (Reading Order)" },
      { field: "color", label: "Color (Grouped by Color)" },
      { field: "date", label: "Date (Modified / Added)" },
      { field: "type", label: "Type (Highlights, Notes...)" },
    ];
    for (const opt of sortOptions) {
      const el = sortSelect.createEl("option", { text: opt.label, value: opt.field });
      if (this.sortField === opt.field) el.selected = true;
    }
    sortSelect.onchange = () => {
      this.sortField = sortSelect.value as AnnotSortField;
      // When switching to date, default to descending (newest first)
      if (this.sortField === "date") {
        this.sortDirection = "desc";
        updateDirBtn();
      } else if (this.sortField === "page") {
        this.sortDirection = "asc";
        updateDirBtn();
      }
      this.renderAnnotationCardsList(cardsContainer, countLabel);
    };

    // Sort Direction Toggle Button
    const dirBtn = sortRow.createEl("button", {
      cls: "pdf-annot-dir-btn",
      attr: { title: `Sort ${this.sortDirection === "asc" ? "Ascending" : "Descending"}` },
    });
    const updateDirBtn = () => {
      dirBtn.empty();
      setIcon(dirBtn, this.sortDirection === "asc" ? "arrow-up-narrow-wide" : "arrow-down-narrow-wide");
      dirBtn.setAttribute("title", `Sort ${this.sortDirection === "asc" ? "Ascending" : "Descending"}`);
    };
    updateDirBtn();

    dirBtn.onclick = () => {
      this.sortDirection = this.sortDirection === "asc" ? "desc" : "asc";
      updateDirBtn();
      this.renderAnnotationCardsList(cardsContainer, countLabel);
    };

    // 3. Quick Color Filter Bar (Zotero Palette Swatches with Semantic Tags)
    const colorFilterRow = viewWrap.createDiv({ cls: "pdf-annot-color-filter-row" });
    this.renderColorFilterRow(colorFilterRow, () => {
      this.renderAnnotationCardsList(cardsContainer, countLabel);
    });

    // 4. Type Filter Chips Row
    const typeFilterRow = viewWrap.createDiv({ cls: "pdf-annot-type-filter-row" });
    const typeFilters: { type: "all" | PdfAnnotationType | "shape"; label: string }[] = [
      { type: "all", label: "All" },
      { type: "highlight", label: "Highlights" },
      { type: "underline", label: "Underlines" },
      { type: "text", label: "Notes" },
      { type: "freetext", label: "Text" },
      { type: "ink", label: "Drawings" },
      { type: "shape", label: "Shapes" },
    ];

    for (const tf of typeFilters) {
      const chip = typeFilterRow.createSpan({
        cls: `pdf-annot-type-chip ${this.filterType === tf.type ? "is-active" : ""}`,
        text: tf.label,
      });
      chip.onclick = () => {
        this.filterType = tf.type;
        typeFilterRow.querySelectorAll(".pdf-annot-type-chip").forEach((c) => c.removeClass("is-active"));
        chip.addClass("is-active");
        this.renderAnnotationCardsList(cardsContainer, countLabel);
      };
    }

    // 5. Actions & Count Header Bar
    const actionsRow = viewWrap.createDiv({ cls: "pdf-annot-summary-row" });
    const countLabel = actionsRow.createSpan({ cls: "pdf-annot-count-label" });

    const exportBtn = actionsRow.createEl("button", {
      cls: "pdf-annot-export-btn",
      attr: { title: "Export annotations to Obsidian Markdown Note" },
    });
    const exportIcon = exportBtn.createSpan({ cls: "pdf-annot-export-icon" });
    setIcon(exportIcon, "file-output");
    exportBtn.createSpan({ text: "Export to Note" });
    exportBtn.onclick = () => {
      this.delegate?.onExportMarkdown?.();
    };

    // 6. Cards Container
    const cardsContainer = viewWrap.createDiv({ cls: "pdf-annot-cards-container" });
    this.renderAnnotationCardsList(cardsContainer, countLabel);
  }

  /**
   * Renders the interactive color palette swatches row in Zotero style.
   */
  private renderColorFilterRow(container: HTMLElement, onFilterChange: () => void): void {
    container.empty();
    if (!this.pdfAdapter) return;

    const allAnnots = this.pdfAdapter.getAllAnnotations();

    // "All" Swatch Pill
    const allPill = container.createSpan({
      cls: `pdf-annot-color-pill ${this.filterColorHex === "all" ? "is-active" : ""}`,
      attr: { title: "Show annotations of all colors" },
    });
    allPill.createSpan({ cls: "pdf-annot-color-pill-dot mod-all" });
    allPill.createSpan({ cls: "pdf-annot-color-pill-text", text: "All" });
    allPill.createSpan({ cls: "pdf-annot-color-pill-count", text: `${allAnnots.length}` });

    allPill.onclick = () => {
      this.filterColorHex = "all";
      container.querySelectorAll(".pdf-annot-color-pill").forEach((p) => p.removeClass("is-active"));
      allPill.addClass("is-active");
      onFilterChange();
    };

    // Distinct colors present in scheme or annotations
    const paletteColors = this.colorScheme?.colors || PDF_COLOR_PRESETS.map((p) => ({
      id: p.name.toLowerCase(),
      hex: p.hex,
      name: p.name,
      tag: "",
      rgb: p.rgb,
    }));

    for (const c of paletteColors) {
      const hex = normalizeHexColor(c.hex);
      // Count annotations matching this color
      const matchingCount = allAnnots.filter((a) => {
        const info = getAnnotationColorInfo(a, this.colorScheme);
        return normalizeHexColor(info.hex) === hex;
      }).length;

      // Always show if it has annotations or if it's in active scheme
      const isSelected = this.filterColorHex === hex;
      const pill = container.createSpan({
        cls: `pdf-annot-color-pill ${isSelected ? "is-active" : ""}`,
        attr: {
          title: c.tag
            ? `${c.name} (${c.tag}): ${matchingCount} annotations`
            : `${c.name}: ${matchingCount} annotations`,
        },
      });

      const dot = pill.createSpan({ cls: "pdf-annot-color-pill-dot" });
      dot.style.backgroundColor = c.hex;

      // Label: Prefer semantic tag if available, otherwise color name
      const labelText = c.tag || c.name;
      pill.createSpan({ cls: "pdf-annot-color-pill-text", text: labelText });
      pill.createSpan({ cls: "pdf-annot-color-pill-count", text: `${matchingCount}` });

      pill.onclick = () => {
        if (this.filterColorHex === hex) {
          // Toggle off -> all
          this.filterColorHex = "all";
          container.querySelectorAll(".pdf-annot-color-pill").forEach((p) => p.removeClass("is-active"));
          allPill.addClass("is-active");
        } else {
          this.filterColorHex = hex;
          container.querySelectorAll(".pdf-annot-color-pill").forEach((p) => p.removeClass("is-active"));
          pill.addClass("is-active");
        }
        onFilterChange();
      };
    }
  }

  /**
   * Main rendering logic for the annotation cards.
   * Handles Sort by Color, Sort by Page, Sort by Date, and Sort by Type.
   */
  private renderAnnotationCardsList(container: HTMLElement, countLabel: HTMLElement): void {
    if (!this.pdfAdapter) return;
    container.empty();

    const rawAnnots = this.pdfAdapter.getAllAnnotations();

    // 1. Filter annotations
    const filtered = filterAnnotations(
      rawAnnots,
      {
        searchQuery: this.searchQuery,
        type: this.filterType,
        colorHex: this.filterColorHex,
      },
      this.colorScheme
    );

    // Update count label
    if (filtered.length === rawAnnots.length) {
      countLabel.setText(`${rawAnnots.length} annotations`);
    } else {
      countLabel.setText(`Showing ${filtered.length} of ${rawAnnots.length} annotations`);
    }

    if (filtered.length === 0) {
      const emptyWrap = container.createDiv({ cls: "pdf-annot-empty-wrap" });
      const emptyIcon = emptyWrap.createSpan({ cls: "pdf-annot-empty-icon" });
      setIcon(emptyIcon, "search-x");
      emptyWrap.createDiv({
        cls: "pdf-annot-empty-title",
        text: rawAnnots.length === 0 ? "No annotations in this document" : "No matching annotations found",
      });
      emptyWrap.createDiv({
        cls: "pdf-annot-empty-desc",
        text:
          rawAnnots.length === 0
            ? "Highlight or underline text, or click tools in the top bar to add annotations."
            : "Try adjusting your search query, color filter, or annotation type filters.",
      });
      return;
    }

    // 2. Render based on Sort Mode
    if (this.sortField === "color") {
      this.renderGroupedByColor(container, filtered);
    } else if (this.sortField === "type") {
      this.renderGroupedByType(container, filtered);
    } else {
      // Linear sorted list (Page or Date)
      const sorted = sortAnnotations(
        filtered,
        { field: this.sortField, direction: this.sortDirection },
        this.colorScheme
      );
      for (const annot of sorted) {
        this.renderSingleAnnotationCard(container, annot);
      }
    }
  }

  /**
   * Renders annotations grouped by color in Zotero style.
   * Each group has a colored section header with tag name, count, and collapsible toggle.
   */
  private renderGroupedByColor(container: HTMLElement, annotations: PdfAnnotation[]): void {
    const groups = groupAnnotationsByColor(annotations, this.colorScheme, this.sortDirection);

    for (const group of groups) {
      const groupEl = container.createDiv({ cls: "pdf-annot-group" });
      const isCollapsed = this.collapsedColorGroups.has(group.key);

      // Section Header
      const headerEl = groupEl.createDiv({
        cls: `pdf-annot-group-header ${isCollapsed ? "is-collapsed" : ""}`,
      });

      const headerLeft = headerEl.createDiv({ cls: "pdf-annot-group-header-left" });
      const chevron = headerLeft.createSpan({ cls: "pdf-annot-group-chevron" });
      setIcon(chevron, isCollapsed ? "chevron-right" : "chevron-down");

      // Color Circle Swatch
      const colorDot = headerLeft.createSpan({ cls: "pdf-annot-group-color-dot" });
      colorDot.style.backgroundColor = group.hex;

      // Title: e.g. "Green · New Info"
      const titleSpan = headerLeft.createSpan({ cls: "pdf-annot-group-title" });
      if (group.tag) {
        titleSpan.createSpan({ cls: "pdf-annot-group-color-name", text: group.name });
        titleSpan.createSpan({ cls: "pdf-annot-group-sep", text: " · " });
        titleSpan.createSpan({ cls: "pdf-annot-group-tag-name", text: group.tag });
      } else {
        titleSpan.createSpan({ cls: "pdf-annot-group-color-name", text: group.name });
      }

      // Count badge
      const countBadge = headerEl.createSpan({
        cls: "pdf-annot-group-count",
        text: `${group.annotations.length}`,
      });

      headerEl.onclick = () => {
        if (this.collapsedColorGroups.has(group.key)) {
          this.collapsedColorGroups.delete(group.key);
        } else {
          this.collapsedColorGroups.add(group.key);
        }
        this.renderActiveTab();
      };

      // Group Items (if expanded)
      if (!isCollapsed) {
        const itemsWrap = groupEl.createDiv({ cls: "pdf-annot-group-items" });
        for (const annot of group.annotations) {
          this.renderSingleAnnotationCard(itemsWrap, annot);
        }
      }
    }
  }

  /**
   * Renders annotations grouped by annotation type.
   */
  private renderGroupedByType(container: HTMLElement, annotations: PdfAnnotation[]): void {
    const groups = groupAnnotationsByType(annotations, this.sortDirection);

    for (const group of groups) {
      const groupEl = container.createDiv({ cls: "pdf-annot-group" });
      const isCollapsed = this.collapsedTypeGroups.has(group.type);

      const headerEl = groupEl.createDiv({
        cls: `pdf-annot-group-header ${isCollapsed ? "is-collapsed" : ""}`,
      });

      const headerLeft = headerEl.createDiv({ cls: "pdf-annot-group-header-left" });
      const chevron = headerLeft.createSpan({ cls: "pdf-annot-group-chevron" });
      setIcon(chevron, isCollapsed ? "chevron-right" : "chevron-down");

      const typeIcon = headerLeft.createSpan({ cls: "pdf-annot-group-type-icon" });
      setIcon(typeIcon, group.icon);

      headerLeft.createSpan({ cls: "pdf-annot-group-title", text: group.label });

      headerEl.createSpan({
        cls: "pdf-annot-group-count",
        text: `${group.annotations.length}`,
      });

      headerEl.onclick = () => {
        if (this.collapsedTypeGroups.has(group.type)) {
          this.collapsedTypeGroups.delete(group.type);
        } else {
          this.collapsedTypeGroups.add(group.type);
        }
        this.renderActiveTab();
      };

      if (!isCollapsed) {
        const itemsWrap = groupEl.createDiv({ cls: "pdf-annot-group-items" });
        for (const annot of group.annotations) {
          this.renderSingleAnnotationCard(itemsWrap, annot);
        }
      }
    }
  }

  /**
   * Renders an individual 1:1 Zotero-style Annotation Card:
   * - Left vertical accent color bar (4px)
   * - Header with Type icon, Semantic Tag badge, Page number link, Date, and Actions
   * - Quoted Text block for Highlights (with subtle tint) and Underlines (with text-decoration)
   * - Attached user comment bubble
   * - Linked Obsidian Markdown note badge
   * - Click to select, scroll to view, and pulse
   */
  private renderSingleAnnotationCard(container: HTMLElement, annot: PdfAnnotation): void {
    const colorInfo = getAnnotationColorInfo(annot, this.colorScheme);
    const hex = colorInfo.hex;
    const isSelected = this.selectedAnnotId === annot.id;

    const card = container.createDiv({
      cls: `pdf-annot-card ${isSelected ? "is-selected" : ""}`,
    });
    card.dataset.annotId = annot.id;
    // Zotero signature left vertical colored accent bar
    card.style.borderLeft = `5px solid ${hex}`;

    // Click card -> jump to page & pulse
    card.onclick = () => {
      this.selectedAnnotId = annot.id;
      container.querySelectorAll(".pdf-annot-card").forEach((c) => c.removeClass("is-selected"));
      card.addClass("is-selected");
      this.delegate?.onSelectAnnotation(annot);
    };

    // 1. Header Row
    const cardHeader = card.createDiv({ cls: "pdf-annot-card-header" });

    const headerLeft = cardHeader.createDiv({ cls: "pdf-annot-card-left" });

    // Annotation Type Icon
    const typeIconEl = headerLeft.createSpan({ cls: "pdf-annot-card-type-icon" });
    const iconName = this.getTypeIconName(annot.type);
    setIcon(typeIconEl, iconName);
    typeIconEl.setAttribute("title", formatTypeLabel(annot.type));

    // Semantic Tag Badge (e.g. "New Info" in green, "Important Points" in red)
    if (colorInfo.tag) {
      const tagBadge = headerLeft.createSpan({
        cls: "pdf-annot-card-tag-badge",
        text: colorInfo.tag,
      });
      tagBadge.style.color = hex;
      tagBadge.style.borderColor = `${hex}55`;
      tagBadge.style.backgroundColor = `${hex}18`;
    }

    // Page badge (clickable to jump)
    const pageBadge = headerLeft.createSpan({
      cls: "pdf-annot-card-page-badge",
      text: `p. ${annot.pageIndex + 1}`,
      attr: { title: `Jump to page ${annot.pageIndex + 1}` },
    });
    pageBadge.onclick = (e) => {
      e.stopPropagation();
      this.selectedAnnotId = annot.id;
      container.querySelectorAll(".pdf-annot-card").forEach((c) => c.removeClass("is-selected"));
      card.addClass("is-selected");
      this.delegate?.onSelectAnnotation(annot);
    };

    // Header Right: Date & Actions
    const headerRight = cardHeader.createDiv({ cls: "pdf-annot-card-right" });

    const dateStr = formatAnnotationDate(annot.modifiedAt || annot.createdAt);
    if (dateStr) {
      headerRight.createSpan({ cls: "pdf-annot-card-date", text: dateStr });
    }

    const actionsWrap = headerRight.createDiv({ cls: "pdf-annot-card-actions" });

    // Quick Color Switcher button
    const colorBtn = actionsWrap.createEl("button", {
      cls: "pdf-annot-card-color-btn",
      attr: { title: "Change annotation color" },
    });
    const colorDot = colorBtn.createSpan({ cls: "pdf-annot-card-color-dot" });
    colorDot.style.backgroundColor = hex;
    colorBtn.onclick = (e) => {
      e.stopPropagation();
      this.showColorPickerMenu(colorBtn, annot);
    };

    // Edit Comment / Note button
    const editBtn = actionsWrap.createEl("button", {
      cls: "pdf-annot-icon-btn",
      attr: { title: "Edit comment or link Obsidian note" },
    });
    setIcon(editBtn, "pencil");
    editBtn.onclick = (e) => {
      e.stopPropagation();
      this.delegate?.onEditAnnotation(annot);
    };

    // Copy Deep Link
    const linkBtn = actionsWrap.createEl("button", {
      cls: "pdf-annot-icon-btn",
      attr: { title: "Copy deep citation link" },
    });
    setIcon(linkBtn, "link");
    linkBtn.onclick = async (e) => {
      e.stopPropagation();
      const link = CitationBuilder.buildDeepLink({
        pdfFileName: this.fileName,
        pageNumber: annot.pageIndex + 1,
        annotationId: annot.id,
      });
      await navigator.clipboard.writeText(link);
      new Notice("Copied annotation deep link!");
    };

    // Delete Annotation
    const delBtn = actionsWrap.createEl("button", {
      cls: "pdf-annot-icon-btn mod-danger",
      attr: { title: "Delete annotation" },
    });
    setIcon(delBtn, "trash");
    delBtn.onclick = (e) => {
      e.stopPropagation();
      this.delegate?.onDeleteAnnotation(annot.id);
    };

    // 2. Card Body: Quoted Text Block (for Highlights & Underlines)
    if (annot.type === "highlight" || annot.type === "underline") {
      const hlText = (annot as any).highlightedText || annot.contents;
      if (hlText && hlText.trim()) {
        const quoteBox = card.createDiv({
          cls: `pdf-annot-card-quote ${annot.type === "underline" ? "mod-underline" : "mod-highlight"}`,
        });
        if (annot.type === "highlight") {
          quoteBox.style.backgroundColor = `${hex}1f`; // 12% opacity tint
          quoteBox.style.borderLeftColor = hex;
        } else {
          quoteBox.style.borderLeftColor = hex;
        }

        quoteBox.createSpan({ cls: "pdf-annot-quote-mark", text: "“" });
        const quoteTextEl = quoteBox.createSpan({
          cls: `pdf-annot-quote-text ${annot.type === "underline" ? "mod-underlined" : ""}`,
          text: hlText.trim(),
        });
        if (annot.type === "underline") {
          quoteTextEl.style.textDecoration = "underline";
          quoteTextEl.style.textDecorationColor = hex;
          quoteTextEl.style.textUnderlineOffset = "3px";
        }
        quoteBox.createSpan({ cls: "pdf-annot-quote-mark", text: "”" });
      } else {
        card.createDiv({
          cls: "pdf-annot-card-quote mod-empty",
          text: annot.type === "underline" ? "Underline" : "Highlight",
        });
      }
    }

    // 3. Card Attached User Comment / Note
    const isQuoteType = annot.type === "highlight" || annot.type === "underline";
    const hlSnippet = isQuoteType ? (annot as any).highlightedText : undefined;
    const hasDistinctComment = annot.contents && (!isQuoteType || annot.contents !== hlSnippet);

    if (hasDistinctComment && annot.contents?.trim()) {
      const commentBox = card.createDiv({ cls: "pdf-annot-card-comment" });
      const commentIcon = commentBox.createSpan({ cls: "pdf-annot-card-comment-icon" });
      setIcon(commentIcon, "message-square");
      commentBox.createSpan({ cls: "pdf-annot-card-comment-text", text: annot.contents.trim() });
    } else if (annot.type === "text" && !annot.contents) {
      card.createDiv({
        cls: "pdf-annot-card-comment mod-empty",
        text: "Sticky Note (click to add comment)",
      });
    } else if (annot.type === "freetext" && annot.contents) {
      const textPreview = card.createDiv({ cls: "pdf-annot-card-freetext-preview" });
      textPreview.createSpan({ text: annot.contents });
    } else if (annot.type === "freetext" && !annot.contents) {
      card.createDiv({
        cls: "pdf-annot-card-comment mod-empty",
        text: "Text Box (empty)",
      });
    } else if (annot.type === "ink" && !annot.contents) {
      card.createDiv({
        cls: "pdf-annot-card-comment mod-empty",
        text: "Handwritten drawing / ink mark",
      });
    }

    // 4. Linked Obsidian Note Badge
    if (annot.linkedNote) {
      const linkedBadge = card.createDiv({
        cls: "pdf-annot-card-linked-note",
        attr: { title: `Open linked note: ${annot.linkedNote}` },
      });
      const noteIcon = linkedBadge.createSpan({ cls: "pdf-annot-link-icon" });
      setIcon(noteIcon, "file-text");
      linkedBadge.createSpan({ cls: "pdf-annot-link-name", text: annot.linkedNote });

      linkedBadge.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onOpenLinkedNote?.(annot.linkedNote!);
      };
    }
  }

  /**
   * Displays quick color picker popup to re-color an annotation.
   */
  private showColorPickerMenu(anchorEl: HTMLElement, annot: PdfAnnotation): void {
    const existing = document.querySelector(".pdf-annot-color-popup");
    if (existing) existing.remove();

    const popup = document.body.createDiv({ cls: "pdf-annot-color-popup" });
    const rect = anchorEl.getBoundingClientRect();
    popup.style.left = `${Math.min(window.innerWidth - 230, Math.max(10, rect.left - 100))}px`;
    popup.style.top = `${rect.bottom + 6}px`;

    popup.createDiv({ cls: "pdf-annot-color-popup-title", text: "Change Annotation Color" });

    const grid = popup.createDiv({ cls: "pdf-annot-color-popup-grid" });
    const colors = this.colorScheme?.colors || PDF_COLOR_PRESETS.map((p) => ({
      id: p.name.toLowerCase(),
      hex: p.hex,
      name: p.name,
      tag: "",
      rgb: p.rgb,
    }));

    for (const c of colors) {
      const item = grid.createDiv({ cls: "pdf-annot-color-popup-item" });
      const dot = item.createSpan({ cls: "pdf-annot-color-popup-dot" });
      dot.style.backgroundColor = c.hex;

      const labelWrap = item.createDiv({ cls: "pdf-annot-color-popup-label-wrap" });
      labelWrap.createSpan({ cls: "pdf-annot-color-popup-name", text: c.name });
      if (c.tag) {
        labelWrap.createSpan({ cls: "pdf-annot-color-popup-tag", text: `· ${c.tag}` });
      }

      item.onclick = (e) => {
        e.stopPropagation();
        popup.remove();
        const rgb = c.rgb || hexToRgb(c.hex);
        this.delegate?.onChangeAnnotationColor?.(annot.id, rgb);
      };
    }

    const closeHandler = (e: MouseEvent) => {
      if (!popup.contains(e.target as Node)) {
        popup.remove();
        window.removeEventListener("mousedown", closeHandler);
      }
    };
    window.setTimeout(() => window.addEventListener("mousedown", closeHandler), 50);
  }

  private getTypeIconName(type: PdfAnnotationType): string {
    switch (type) {
      case "highlight":
        return "highlighter";
      case "underline":
        return "underline";
      case "text":
        return "message-square";
      case "freetext":
        return "type";
      case "ink":
        return "pen-tool";
      case "square":
        return "square";
      case "circle":
        return "circle";
      case "line":
        return "minus";
      default:
        return "bookmark";
    }
  }

  // =========================================================================
  // ==================== 2. OUTLINE TAB (ZOTERO TOC) ========================
  // =========================================================================

  private async renderOutline(): Promise<void> {
    if (!this.contentEl) return;
    const wrap = this.contentEl.createDiv({ cls: "pdf-outline-wrap" });

    if (!this.pdfjsDoc) {
      wrap.createDiv({ cls: "pdf-annot-empty-state", text: "Loading document outline..." });
      return;
    }

    try {
      const outline = await this.pdfjsDoc.getOutline();
      if (!outline || !Array.isArray(outline) || outline.length === 0) {
        const emptyWrap = wrap.createDiv({ cls: "pdf-annot-empty-wrap" });
        const iconSpan = emptyWrap.createSpan({ cls: "pdf-annot-empty-icon" });
        setIcon(iconSpan, "list-tree");
        emptyWrap.createDiv({ cls: "pdf-annot-empty-title", text: "No Table of Contents" });
        emptyWrap.createDiv({
          cls: "pdf-annot-empty-desc",
          text: "This document does not contain an embedded outline or bookmarks.",
        });
        return;
      }

      const listContainer = wrap.createDiv({ cls: "pdf-outline-list" });
      this.renderOutlineItems(listContainer, outline, 0);
    } catch (err) {
      wrap.createDiv({ cls: "pdf-annot-empty-state", text: "Could not load outline for this PDF." });
    }
  }

  private renderOutlineItems(container: HTMLElement, items: any[], depth: number): void {
    for (const item of items) {
      const itemRow = container.createDiv({ cls: "pdf-outline-item" });
      itemRow.style.paddingLeft = `${depth * 14 + 6}px`;

      const hasChildren = item.items && item.items.length > 0;
      let isCollapsed = false;

      // Expand/Collapse Caret
      const caret = itemRow.createSpan({ cls: "pdf-outline-caret" });
      if (hasChildren) {
        setIcon(caret, "chevron-down");
      } else {
        caret.addClass("is-leaf");
      }

      // Title Link
      const titleSpan = itemRow.createSpan({ cls: "pdf-outline-title", text: item.title || "Untitled" });
      if (item.bold) titleSpan.style.fontWeight = "bold";
      if (item.italic) titleSpan.style.fontStyle = "italic";

      let childrenWrap: HTMLElement | null = null;

      if (hasChildren) {
        childrenWrap = container.createDiv({ cls: "pdf-outline-children" });
        this.renderOutlineItems(childrenWrap, item.items, depth + 1);

        caret.onclick = (e) => {
          e.stopPropagation();
          isCollapsed = !isCollapsed;
          caret.empty();
          setIcon(caret, isCollapsed ? "chevron-right" : "chevron-down");
          childrenWrap?.toggleClass("is-hidden", isCollapsed);
        };
      }

      // Clicking item navigates to destination
      itemRow.onclick = async () => {
        container.querySelectorAll(".pdf-outline-item").forEach((el) => el.removeClass("is-active"));
        itemRow.addClass("is-active");
        if (item.dest) {
          await this.navigateToDestination(item.dest);
        }
      };
    }
  }

  private async navigateToDestination(dest: any): Promise<void> {
    if (!dest || !this.pdfjsDoc) return;
    try {
      let explicitDest = dest;
      if (typeof dest === "string") {
        explicitDest = await this.pdfjsDoc.getDestination(dest);
      }
      if (!explicitDest || !Array.isArray(explicitDest) || explicitDest.length === 0) {
        return;
      }
      const target = explicitDest[0];
      let pageIndex: number | null = null;
      if (typeof target === "number") {
        pageIndex = target;
      } else if (target && typeof target === "object") {
        pageIndex = await this.pdfjsDoc.getPageIndex(target);
      }
      if (pageIndex !== null && pageIndex >= 0) {
        this.delegate?.onSelectPage(pageIndex);
      }
    } catch (err) {
      console.warn("Failed to navigate to outline destination:", err);
    }
  }

  // =========================================================================
  // ==================== 3. THUMBNAILS TAB ==================================
  // =========================================================================

  private renderThumbnails(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    const count = this.pdfAdapter.getPageCount();
    const grid = this.contentEl.createDiv({ cls: "pdf-thumbnail-grid" });

    for (let i = 0; i < count; i++) {
      const pageNum = i + 1;
      const isCurrent = i === this.currentPageIndex;
      const thumbCard = grid.createDiv({
        cls: `pdf-thumbnail-card ${isCurrent ? "is-active" : ""}`,
      });
      thumbCard.dataset.pageIndex = `${i}`;
      thumbCard.onclick = () => {
        this.currentPageIndex = i;
        this.updateThumbnailActiveState();
        this.delegate?.onSelectPage(i);
      };

      const canvasWrap = thumbCard.createDiv({ cls: "pdf-thumbnail-canvas-wrap" });
      const canvas = canvasWrap.createEl("canvas", { cls: "pdf-thumbnail-canvas" });

      const footer = thumbCard.createDiv({ cls: "pdf-thumbnail-footer" });
      footer.createSpan({ cls: "pdf-thumbnail-page-num", text: `${pageNum}` });

      // Actions menu: Rotate, Duplicate, Delete
      const actionsBtn = footer.createEl("button", {
        cls: "pdf-thumbnail-menu-btn",
        attr: { title: "Page actions" },
      });
      setIcon(actionsBtn, "more-vertical");
      actionsBtn.onclick = (e) => {
        e.stopPropagation();
        this.showPageActionsMenu(actionsBtn, i);
      };

      if (this.pdfjsDoc) {
        this.renderMiniThumbnail(canvas, pageNum);
      }
    }
  }

  private updateThumbnailActiveState(): void {
    if (!this.contentEl) return;
    const cards = this.contentEl.querySelectorAll(".pdf-thumbnail-card");
    cards.forEach((c) => {
      const idx = parseInt((c as HTMLElement).dataset.pageIndex || "-1", 10);
      c.toggleClass("is-active", idx === this.currentPageIndex);
    });
  }

  private async renderMiniThumbnail(canvas: HTMLCanvasElement, pageNum: number): Promise<void> {
    try {
      const page = await this.pdfjsDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: 0.25 });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        await page.render({ canvasContext: ctx, viewport }).promise;
      }
    } catch {
      // Thumbnail rendering fallback
    }
  }

  private showPageActionsMenu(anchorEl: HTMLElement, pageIndex: number): void {
    const existing = document.querySelector(".pdf-thumb-action-popup");
    if (existing) existing.remove();

    const popup = document.body.createDiv({ cls: "pdf-thumb-action-popup" });
    const rect = anchorEl.getBoundingClientRect();
    popup.style.left = `${rect.right + 4}px`;
    popup.style.top = `${rect.top}px`;

    const rotateItem = popup.createDiv({ cls: "pdf-thumb-action-item", text: "Rotate 90°" });
    rotateItem.onclick = () => {
      popup.remove();
      this.delegate?.onRotatePage(pageIndex);
    };

    const dupItem = popup.createDiv({ cls: "pdf-thumb-action-item", text: "Duplicate Page" });
    dupItem.onclick = () => {
      popup.remove();
      this.delegate?.onDuplicatePage(pageIndex);
    };

    const delItem = popup.createDiv({ cls: "pdf-thumb-action-item mod-danger", text: "Delete Page" });
    delItem.onclick = () => {
      popup.remove();
      this.delegate?.onDeletePage(pageIndex);
    };

    const closeHandler = (e: MouseEvent) => {
      if (!popup.contains(e.target as Node)) {
        popup.remove();
        window.removeEventListener("mousedown", closeHandler);
      }
    };
    window.setTimeout(() => window.addEventListener("mousedown", closeHandler), 50);
  }

  dispose(): void {
    this.containerEl?.remove();
    this.containerEl = null;
    this.contentEl = null;
    this.delegate = null;
    this.pdfAdapter = null;
    this.pdfjsDoc = null;
  }
}
