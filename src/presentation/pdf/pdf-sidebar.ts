/**
 * Collapsible PDF Sidebar containing Page Thumbnails and the Annotations List.
 * Similar to Adobe Acrobat's Comments & Pages Navigation pane.
 */

import { setIcon, Notice } from "obsidian";
import { PdfAnnotation, PdfAnnotationType, rgbToHex, CitationBuilder } from "../../domain/pdf";
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
}

export type SidebarTab = "thumbnails" | "annotations";

export class PdfSidebar {
  private containerEl: HTMLElement | null = null;
  private contentEl: HTMLElement | null = null;
  private activeTab: SidebarTab = "annotations";
  private filterType: "all" | PdfAnnotationType = "all";
  private searchQuery = "";
  private delegate: PdfSidebarDelegate | null = null;
  private pdfAdapter: PdfDocumentAdapter | null = null;
  private pdfjsDoc: any = null;
  private fileName = "";
  private isVisible = true;
  private annotTabBtn: HTMLElement | null = null;
  public onVisibilityChange?: (visible: boolean) => void;

  build(
    parentEl: HTMLElement,
    pdfAdapter: PdfDocumentAdapter,
    pdfjsDoc: any,
    fileName: string,
    delegate: PdfSidebarDelegate
  ): void {
    this.pdfAdapter = pdfAdapter;
    this.pdfjsDoc = pdfjsDoc;
    this.fileName = fileName;
    this.delegate = delegate;

    this.containerEl = parentEl.createDiv({ cls: "pdf-sidebar-container" });

    // Resizer drag handle
    const resizer = this.containerEl.createDiv({ cls: "pdf-sidebar-resizer" });
    let startX = 0;
    let startWidth = 0;

    resizer.addEventListener("pointerdown", (e: PointerEvent) => {
      e.preventDefault();
      startX = e.clientX;
      startWidth = this.containerEl!.getBoundingClientRect().width;

      const onPointerMove = (me: PointerEvent) => {
        const newW = Math.max(240, Math.min(550, startWidth + (me.clientX - startX)));
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

    // Sidebar Header with Tabs & Close button
    const headerEl = this.containerEl.createDiv({ cls: "pdf-sidebar-header" });
    const tabsEl = headerEl.createDiv({ cls: "pdf-sidebar-tabs" });

    const annotTabBtn = tabsEl.createEl("button", {
      cls: `pdf-sidebar-tab ${this.activeTab === "annotations" ? "is-active" : ""}`,
      text: `Annotations (${pdfAdapter.getAllAnnotations().length})`,
    });
    annotTabBtn.onclick = () => this.switchTab("annotations");
    this.annotTabBtn = annotTabBtn;

    const thumbTabBtn = tabsEl.createEl("button", {
      cls: `pdf-sidebar-tab ${this.activeTab === "thumbnails" ? "is-active" : ""}`,
      text: "Pages",
    });
    thumbTabBtn.onclick = () => this.switchTab("thumbnails");

    const closeBtn = headerEl.createEl("button", { cls: "pdf-sidebar-close-btn", title: "Retract side panel" });
    setIcon(closeBtn, "chevrons-left");
    closeBtn.onclick = () => this.toggle();

    // Content container
    this.contentEl = this.containerEl.createDiv({ cls: "pdf-sidebar-content" });

    this.renderActiveTab();
  }

  toggle(): boolean {
    if (!this.containerEl) return false;
    this.isVisible = !this.isVisible;
    if (this.isVisible) {
      this.containerEl.removeClass("is-hidden");
    } else {
      this.containerEl.addClass("is-hidden");
    }
    this.onVisibilityChange?.(this.isVisible);
    return this.isVisible;
  }

  show(): void {
    this.isVisible = true;
    this.containerEl?.removeClass("is-hidden");
    this.onVisibilityChange?.(true);
  }

  hide(): void {
    this.isVisible = false;
    this.containerEl?.addClass("is-hidden");
    this.onVisibilityChange?.(false);
  }

  switchTab(tab: SidebarTab): void {
    this.activeTab = tab;
    if (!this.containerEl) return;
    const tabs = this.containerEl.querySelectorAll(".pdf-sidebar-tab");
    tabs.forEach((t) => {
      const isAnnot = t === this.annotTabBtn;
      if ((isAnnot && tab === "annotations") || (!isAnnot && tab === "thumbnails")) {
        t.addClass("is-active");
      } else {
        t.removeClass("is-active");
      }
    });
    this.renderActiveTab();
  }

  updateData(pdfAdapter: PdfDocumentAdapter, pdfjsDoc: any): void {
    this.pdfAdapter = pdfAdapter;
    this.pdfjsDoc = pdfjsDoc;
    if (this.annotTabBtn) {
      this.annotTabBtn.setText(`Annotations (${pdfAdapter.getAllAnnotations().length})`);
    }
    this.renderActiveTab();
  }

  private renderActiveTab(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    this.contentEl.empty();

    if (this.activeTab === "thumbnails") {
      this.renderThumbnails();
    } else {
      this.renderAnnotationsList();
    }
  }

  // ==================== Thumbnails Tab ====================
  private renderThumbnails(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    const count = this.pdfAdapter.getPageCount();
    const grid = this.contentEl.createDiv({ cls: "pdf-thumbnail-grid" });

    for (let i = 0; i < count; i++) {
      const pageNum = i + 1;
      const thumbCard = grid.createDiv({ cls: "pdf-thumbnail-card" });
      thumbCard.onclick = () => this.delegate?.onSelectPage(i);

      const canvasWrap = thumbCard.createDiv({ cls: "pdf-thumbnail-canvas-wrap" });
      const canvas = canvasWrap.createEl("canvas", { cls: "pdf-thumbnail-canvas" });

      const footer = thumbCard.createDiv({ cls: "pdf-thumbnail-footer" });
      footer.createSpan({ cls: "pdf-thumbnail-page-num", text: `${pageNum}` });

      // Actions menu button: Rotate, Duplicate, Delete
      const actionsBtn = footer.createEl("button", { cls: "pdf-thumbnail-menu-btn", title: "Page actions" });
      setIcon(actionsBtn, "more-vertical");
      actionsBtn.onclick = (e) => {
        e.stopPropagation();
        this.showPageActionsMenu(actionsBtn, i);
      };

      // Asynchronously render mini thumbnail
      if (this.pdfjsDoc) {
        this.renderMiniThumbnail(canvas, pageNum);
      }
    }
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
    } catch (err) {
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
    setTimeout(() => window.addEventListener("mousedown", closeHandler), 50);
  }

  // ==================== Annotations List Tab ====================
  private renderAnnotationsList(): void {
    if (!this.contentEl || !this.pdfAdapter) return;
    const listWrap = this.contentEl.createDiv({ cls: "pdf-annot-list-wrap" });

    // 1-Click Export to Obsidian Markdown Note
    const allAnnots = this.pdfAdapter.getAllAnnotations();
    const exportBtn = listWrap.createEl("button", {
      cls: "pdf-export-all-btn mod-cta",
      attr: { title: "Export all highlights, notes, text boxes, and shapes into an Obsidian Markdown note" },
    });
    exportBtn.createSpan({ text: "📥" });
    exportBtn.createSpan({
      text: allAnnots.length > 0 ? `Export All to Note (${allAnnots.length})` : "Export All to Note",
    });
    exportBtn.onclick = () => {
      this.delegate?.onExportMarkdown?.();
    };

    // Search bar
    const searchWrap = listWrap.createDiv({ cls: "pdf-annot-search-wrap" });
    const searchInput = searchWrap.createEl("input", {
      cls: "pdf-annot-search-input",
      type: "search",
      value: this.searchQuery,
      attr: { placeholder: "Search annotations..." },
    });
    searchInput.oninput = () => {
      this.searchQuery = searchInput.value.toLowerCase().trim();
      this.renderAnnotationCards(cardsContainer);
    };

    // Filter pills
    const filterRow = listWrap.createDiv({ cls: "pdf-annot-filter-row" });
    const filters: { type: "all" | PdfAnnotationType; label: string }[] = [
      { type: "all", label: "All" },
      { type: "highlight", label: "Highlights" },
      { type: "text", label: "Notes" },
      { type: "ink", label: "Drawings" },
      { type: "freetext", label: "Text" },
      { type: "square", label: "Shapes" },
    ];

    for (const f of filters) {
      const pill = filterRow.createSpan({
        cls: `pdf-annot-filter-pill ${this.filterType === f.type ? "is-active" : ""}`,
        text: f.label,
      });
      pill.onclick = () => {
        this.filterType = f.type;
        filterRow.querySelectorAll(".pdf-annot-filter-pill").forEach((p) => p.removeClass("is-active"));
        pill.addClass("is-active");
        this.renderAnnotationCards(cardsContainer);
      };
    }

    const cardsContainer = listWrap.createDiv({ cls: "pdf-annot-cards-container" });
    this.renderAnnotationCards(cardsContainer);
  }

  private renderAnnotationCards(container: HTMLElement): void {
    if (!this.pdfAdapter) return;
    container.empty();

    let annots = this.pdfAdapter.getAllAnnotations();

    // Type filter
    if (this.filterType !== "all") {
      if (this.filterType === "square") {
        annots = annots.filter((a) => a.type === "square" || a.type === "circle" || a.type === "line");
      } else {
        annots = annots.filter((a) => a.type === this.filterType);
      }
    }

    // Search query filter
    if (this.searchQuery) {
      annots = annots.filter((a) => {
        const text = ((a as any).highlightedText || a.contents || "").toLowerCase();
        const author = (a.author || "").toLowerCase();
        return text.includes(this.searchQuery) || author.includes(this.searchQuery);
      });
    }

    if (annots.length === 0) {
      container.createDiv({
        cls: "pdf-annot-empty-state",
        text: "No matching annotations found.",
      });
      return;
    }

    // Sort by pageIndex, then top position
    annots.sort((a, b) => (a.pageIndex - b.pageIndex) || ((b.rect?.[1] || 0) - (a.rect?.[1] || 0)));

    for (const annot of annots) {
      const card = container.createDiv({ cls: "pdf-annot-card" });
      card.dataset.annotId = annot.id;

      // Click card -> jump to page & highlight
      card.onclick = () => {
        this.delegate?.onSelectAnnotation(annot);
      };

      // Header: Dot, Type Name, Page badge, Actions
      const cardHeader = card.createDiv({ cls: "pdf-annot-card-header" });

      const leftInfo = cardHeader.createDiv({ cls: "pdf-annot-card-left" });
      const dot = leftInfo.createSpan({ cls: "pdf-annot-card-dot" });
      dot.style.backgroundColor = rgbToHex(annot.color);

      const typeLabel = this.formatTypeLabel(annot.type);
      leftInfo.createSpan({ cls: "pdf-annot-card-type", text: typeLabel });

      const pageBadge = leftInfo.createSpan({
        cls: "pdf-annot-card-page-badge",
        text: `p. ${annot.pageIndex + 1}`,
      });

      const cardActions = cardHeader.createDiv({ cls: "pdf-annot-card-actions" });

      // Edit Note button
      const editBtn = cardActions.createEl("button", {
        cls: "pdf-annot-icon-btn",
        title: "Edit comment or link note",
      });
      setIcon(editBtn, "pencil");
      editBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onEditAnnotation(annot);
      };

      // Copy Deep Link
      const linkBtn = cardActions.createEl("button", {
        cls: "pdf-annot-icon-btn",
        title: "Copy deep link",
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

      // Delete
      const deleteBtn = cardActions.createEl("button", {
        cls: "pdf-annot-icon-btn mod-danger",
        title: "Delete annotation",
      });
      setIcon(deleteBtn, "trash");
      deleteBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onDeleteAnnotation(annot.id);
      };

      // Quoted text for highlights
      if (annot.type === "highlight") {
        const hlText = (annot as any).highlightedText || annot.contents;
        if (hlText && hlText.trim()) {
          card.createDiv({
            cls: "pdf-annot-card-quote",
            text: `“${hlText.trim()}”`,
          });
        } else {
          card.createDiv({
            cls: "pdf-annot-card-quote mod-empty",
            text: "Highlight",
          });
        }
      }

      // Comment contents if present
      if (annot.contents && (annot.type !== "highlight" || annot.contents !== (annot as any).highlightedText)) {
        card.createDiv({
          cls: "pdf-annot-card-comment",
          text: annot.contents,
        });
      } else if (annot.type === "text" && !annot.contents) {
        card.createDiv({
          cls: "pdf-annot-card-comment mod-empty",
          text: "Note (click to edit comment)",
        });
      } else if (annot.type === "ink" && !annot.contents) {
        card.createDiv({
          cls: "pdf-annot-card-comment mod-empty",
          text: "Drawing",
        });
      } else if ((annot.type === "square" || annot.type === "circle" || annot.type === "line") && !annot.contents) {
        card.createDiv({
          cls: "pdf-annot-card-comment mod-empty",
          text: this.formatTypeLabel(annot.type),
        });
      }

      // Linked note tag
      if (annot.linkedNote) {
        const linkedBadge = card.createDiv({ cls: "pdf-annot-card-linked-note" });
        setIcon(linkedBadge.createSpan({ cls: "pdf-annot-link-icon" }), "file-text");
        linkedBadge.createSpan({ text: annot.linkedNote });
      }
    }
  }

  private formatTypeLabel(type: PdfAnnotationType): string {
    switch (type) {
      case "highlight":
        return "Highlight";
      case "text":
        return "Sticky Note";
      case "ink":
        return "Drawing";
      case "freetext":
        return "Text Box";
      case "square":
        return "Rectangle";
      case "circle":
        return "Circle";
      case "line":
        return "Line";
      default:
        return "Annotation";
    }
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
