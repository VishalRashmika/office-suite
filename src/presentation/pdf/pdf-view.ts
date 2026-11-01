/**
 * PDF View component for Obsidian.
 * Integrates native PDF editing, cross-compatible ISO 32000-1 annotations,
 * citations, deep links, and annotations list panel.
 */

import { FileView, TFile, WorkspaceLeaf, Notice, loadPdfJs } from "obsidian";
import {
  PdfAnnotation,
  PdfHighlightAnnotation,
  PdfUnderlineAnnotation,
  PdfNoteAnnotation,
  PdfFreeTextAnnotation,
  PdfInkAnnotation,
  PdfShapeAnnotation,
  PdfRect,
  PdfRgbColor,
  PDF_COLOR_PRESETS,
  rgbToHex,
  hexToRgb,
  CitationService,
  CitationMetadata,
  PdfColorSchemeProfile,
  ZOTERO_DEFAULT_PROFILE,
  cloneProfile,
  ensureValidProfile,
  findTagForColor,
  normalizeHexColor,
} from "../../domain/pdf";
import { PdfDocumentAdapter } from "../../infrastructure/pdf";
import { PdfToolbar, PdfToolbarDelegate, PdfActiveTool } from "./pdf-toolbar";
import { PdfSidebar, PdfSidebarDelegate } from "./pdf-sidebar";
import { PdfPageView, PdfPageViewDelegate } from "./pdf-page-view";
import { PdfNotePopover } from "./pdf-note-popover";
import { PdfCiteModal } from "./pdf-cite-modal";
import { PdfColorSchemeModal } from "./pdf-color-scheme-modal";
import { TextPromptModal } from "../common/text-prompt-modal";
import type OfficeSuitePlugin from "../../main";

export const VIEW_TYPE_PDF = "pdf-editor-view";

export class PdfView
  extends FileView
  implements PdfToolbarDelegate, PdfSidebarDelegate, PdfPageViewDelegate
{
  allowNoFile = false;
  pdfDoc: PdfDocumentAdapter | null = null;
  readerTheme: "light" | "sepia" | "dark" = "light";
  paperTheme: "light" | "dark" = "light";

  private plugin?: OfficeSuitePlugin;
  private activeColorScheme: PdfColorSchemeProfile = ZOTERO_DEFAULT_PROFILE;
  private unregisterColorSchemeListener?: () => void;

  private pdfjsLib: any = null;
  private pdfjsDoc: any = null;
  private toolbar = new PdfToolbar();
  private sidebar = new PdfSidebar();
  private notePopover: PdfNotePopover;

  private rootEl: HTMLElement | null = null;
  private toolbarEl: HTMLElement | null = null;
  private mainContainerEl: HTMLElement | null = null;
  private pagesScrollEl: HTMLElement | null = null;
  private pagesContainerEl: HTMLElement | null = null;
  private selectionBarEl: HTMLElement | null = null;
  private selectionColorsWrap: HTMLElement | null = null;

  private pageViews: Map<number, PdfPageView> = new Map();
  private activeTool: PdfActiveTool = "hand";
  private activeColor: PdfRgbColor = [0.99, 0.88, 0.28];
  private strokeWidth = 2;
  private zoomScale = 1.0;
  private currentPage = 1;
  private saveTimer: number | null = null;
  private undoStack: Array<string> = [];
  private redoStack: Array<string> = [];
  private retractTabEl: HTMLElement | null = null;
  private isSaving = false;
  private hasPendingSave = false;
  private isInternalSaving = false;
  private lastInternalSaveTime = 0;

  // Zoom & Pinch Gestures State
  private isGesturePinching = false;
  private currentRenderId = 0;

  // Trackpad & Wheel zoom state
  private isWheelZooming = false;
  private wheelZoomDebounceTimer: number | null = null;
  private startWheelZoomScale = 1.0;
  private cumulativeWheelScale = 1.0;
  private wheelFocalX = 0;
  private wheelFocalY = 0;

  // Touchscreen 2-finger pinch state
  private isTouchPinching = false;
  private touchStartDist = 0;
  private touchStartScale = 1.0;
  private touchStartMidX = 0;
  private touchStartMidY = 0;
  private touchCurrentMidX = 0;
  private touchCurrentMidY = 0;
  private touchCurrentRatio = 1.0;
  private touchCooldownTimer: number | null = null;

  constructor(leaf: WorkspaceLeaf, plugin?: OfficeSuitePlugin) {
    super(leaf);
    this.plugin = plugin;
    this.notePopover = new PdfNotePopover(this.app);
  }

  getPlugin(): OfficeSuitePlugin | undefined {
    if (this.plugin) return this.plugin;
    return (
      (this.app as any).plugins?.getPlugin?.("office-suite") ||
      (this.app as any).plugins?.plugins?.["office-suite"]
    );
  }

  getViewType(): string {
    return VIEW_TYPE_PDF;
  }

  getIcon(): string {
    return "file-text";
  }

  getDisplayText(): string {
    return this.file?.basename ?? "PDF";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "pdf";
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("pdf-editor-root");

    const statusEl = this.contentEl.createDiv({
      cls: "pdf-editor-status",
      text: "Loading PDF document...",
    });

    let bytes: ArrayBuffer;
    try {
      bytes = await this.app.vault.readBinary(file);
      this.pdfDoc = await PdfDocumentAdapter.load(bytes);
    } catch (err) {
      statusEl.setText(`Could not open PDF file: ${(err as Error).message}`);
      console.error(err);
      return;
    }

    try {
      const pdfjsLib = await this.getPdfJsLibrary();
      this.pdfjsLib = pdfjsLib;
      this.pdfjsDoc = await pdfjsLib.getDocument({
        data: new Uint8Array(bytes.slice(0)),
        cMapUrl: "https://unpkg.com/pdfjs-dist@3.11.174/cmaps/",
        cMapPacked: true,
      }).promise;
    } catch (err) {
      statusEl.setText(`Could not render PDF preview: ${(err as Error).message}`);
      console.error(err);
      return;
    }

    statusEl.remove();

    // Enrich annotations from PDF.js (discovers PDF.js parsed annotations & extracts underlying text)
    await this.enrichAnnotationsFromPdfJs();

    // Initialize active color scheme for this PDF
    this.initColorSchemeForFile(file.path);

    // 1. Build Toolbar
    this.toolbarEl = this.contentEl.createDiv({ cls: "pdf-editor-toolbar-wrap" });
    this.toolbar.build(this.toolbarEl, this);
    this.toolbar.setColorScheme(this.activeColorScheme, this.getPlugin()?.settings.pdfProfiles);
    this.toolbar.setPageData(1, this.pdfDoc.getPageCount());

    // 2. Main split view (Sidebar + Scroll Area)
    this.mainContainerEl = this.contentEl.createDiv({ cls: "pdf-editor-main-container" });

    // Retractable Edge Tab for expanding/retracting sidebar when retracted
    this.retractTabEl = this.mainContainerEl.createDiv({
      cls: "pdf-sidebar-retract-tab is-hidden",
      attr: { title: "Expand Annotations & Pages Panel" },
    });
    this.retractTabEl.createSpan({ text: "📑" });
    this.retractTabEl.createSpan({ cls: "pdf-sidebar-retract-label", text: "Annotations" });
    this.retractTabEl.onclick = () => this.onToggleSidebar();

    // Sidebar
    this.sidebar.build(
      this.mainContainerEl,
      this.pdfDoc,
      this.pdfjsDoc,
      file.name,
      this,
      this.activeColorScheme
    );
    this.sidebar.onVisibilityChange = (visible) => {
      this.toolbar.setSidebarActive(visible);
      if (visible) {
        this.retractTabEl?.addClass("is-hidden");
      } else {
        this.retractTabEl?.removeClass("is-hidden");
      }
    };

    // Note Popover
    this.notePopover.setColorScheme(this.activeColorScheme);

    // Register color scheme listener from plugin
    this.unregisterColorSchemeListener?.();
    this.unregisterColorSchemeListener = this.getPlugin()?.onColorSchemeChanged(() => {
      const plugin = this.getPlugin();
      if (this.file && plugin) {
        const fileScheme = plugin.settings.pdfFileSchemes[this.file.path];
        if (fileScheme) {
          this.applyColorScheme(fileScheme, false);
        } else {
          this.toolbar.setColorScheme(this.activeColorScheme, plugin.settings.pdfProfiles);
        }
      }
    });

    // Scrollable Pages Area
    this.pagesScrollEl = this.mainContainerEl.createDiv({ cls: "pdf-pages-scroll-area" });
    this.pagesScrollEl.addEventListener("scroll", this.onScroll);
    this.notePopover.setScrollContainer(this.pagesScrollEl);

    // Inner Pages Container for hardware-accelerated transforms and flex alignment
    this.pagesContainerEl = this.pagesScrollEl.createDiv({ cls: "pdf-pages-container" });

    // 2-Finger Zoom Gestures (Trackpad Pinch, Ctrl/Cmd + Wheel, Touchscreen Pinch, WebKit Gestures)
    this.pagesScrollEl.addEventListener("wheel", this.onWheel, { passive: false });
    this.pagesScrollEl.addEventListener("touchstart", this.onTouchStart, { passive: false });
    window.addEventListener("touchmove", this.onTouchMove, { passive: false });
    window.addEventListener("touchend", this.onTouchEnd);
    window.addEventListener("touchcancel", this.onTouchCancel);

    this.pagesScrollEl.addEventListener("gesturestart", this.onGestureStart as any, { passive: false });
    this.pagesScrollEl.addEventListener("gesturechange", this.onGestureChange as any, { passive: false });
    this.pagesScrollEl.addEventListener("gestureend", this.onGestureEnd as any, { passive: false });

    // Contextual Text Selection Action Bar
    this.buildTextSelectionBar();

    // Render all pages
    await this.renderAllPages();

    // Handle deep-link target if leaf URL has #page=N
    this.handleDeepLinkNavigation();

    // Listen for external sync file modifications (e.g. Obsidian Sync from another computer)
    this.registerEvent(
      this.app.vault.on("modify", async (modifiedFile) => {
        if (
          modifiedFile.path === this.file?.path &&
          !this.isInternalSaving &&
          Date.now() - this.lastInternalSaveTime > 2500 &&
          this.pdfDoc
        ) {
          await this.reloadFromExternalSync();
        }
      })
    );

    window.addEventListener("blur", this.onWindowBlur);
    this.contentEl.addEventListener("keydown", this.onKeyDown);
  }

  private onKeyDown = (e: KeyboardEvent) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      e.stopPropagation();
      void this.requestSave();
    }
  };

  private onWindowBlur = () => {
    if (this.pdfDoc?.dirty) void this.saveNow();
  };

  async onUnloadFile(file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
    this.contentEl.removeEventListener("keydown", this.onKeyDown);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.wheelZoomDebounceTimer) window.clearTimeout(this.wheelZoomDebounceTimer);
    if (this.touchCooldownTimer) window.clearTimeout(this.touchCooldownTimer);
    if (this.pdfDoc?.dirty) await this.saveNow();

    this.pagesScrollEl?.removeEventListener("wheel", this.onWheel);
    this.pagesScrollEl?.removeEventListener("touchstart", this.onTouchStart);
    window.removeEventListener("touchmove", this.onTouchMove);
    window.removeEventListener("touchend", this.onTouchEnd);
    window.removeEventListener("touchcancel", this.onTouchCancel);

    this.pagesScrollEl?.removeEventListener("gesturestart", this.onGestureStart as any);
    this.pagesScrollEl?.removeEventListener("gesturechange", this.onGestureChange as any);
    this.pagesScrollEl?.removeEventListener("gestureend", this.onGestureEnd as any);

    this.notePopover.hide();
    this.toolbar.dispose();
    this.sidebar.dispose();
    for (const pv of this.pageViews.values()) {
      pv.dispose();
    }
    this.pageViews.clear();

    if (this.pdfjsDoc) {
      this.pdfjsDoc.destroy();
      this.pdfjsDoc = null;
    }
    this.pdfDoc = null;
    this.pagesContainerEl = null;
  }

  // ==================== Page Rendering & Layout ====================

  private async renderAllPages(preserveViews = false): Promise<void> {
    if (!this.pagesScrollEl || !this.pagesContainerEl || !this.pdfDoc || !this.pdfjsDoc) return;
    const renderId = ++this.currentRenderId;
    const pageCount = this.pdfDoc.getPageCount();

    if (preserveViews && this.pageViews.size === pageCount) {
      for (let i = 0; i < pageCount; i++) {
        if (this.currentRenderId !== renderId) return;
        const pageView = this.pageViews.get(i);
        if (pageView) {
          try {
            const pdfjsPage = await this.pdfjsDoc.getPage(i + 1);
            if (this.currentRenderId !== renderId) return;
            await pageView.render(pdfjsPage, this.zoomScale, this.pdfjsLib);
          } catch (err) {
            console.error(`Error re-rendering page ${i + 1}:`, err);
          }
        }
      }
      this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
      return;
    }

    this.pagesContainerEl.empty();
    for (const pv of this.pageViews.values()) {
      pv.dispose();
    }
    this.pageViews.clear();

    for (let i = 0; i < pageCount; i++) {
      if (this.currentRenderId !== renderId) return;
      const pageView = new PdfPageView(i, this.pagesContainerEl);
      pageView.setDelegate(this);
      pageView.setColorScheme(this.activeColorScheme);
      pageView.setTool(this.activeTool, this.activeColor, this.strokeWidth);
      pageView.setAnnotations(this.pdfDoc.getPageAnnotations(i));
      this.pageViews.set(i, pageView);

      try {
        const pdfjsPage = await this.pdfjsDoc.getPage(i + 1);
        if (this.currentRenderId !== renderId) return;
        await pageView.render(pdfjsPage, this.zoomScale, this.pdfjsLib);
      } catch (err) {
        console.error(`Error rendering page ${i + 1}:`, err);
      }
    }

    this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
  }

  private onScroll = (): void => {
    if (!this.pagesScrollEl) return;
    const scrollTop = this.pagesScrollEl.scrollTop;
    let closestPage = 1;
    let minDistance = Infinity;

    for (const [idx, pv] of this.pageViews.entries()) {
      const top = pv.element.offsetTop;
      const dist = Math.abs(top - scrollTop);
      if (dist < minDistance) {
        minDistance = dist;
        closestPage = idx + 1;
      }
    }

    if (closestPage !== this.currentPage) {
      this.currentPage = closestPage;
      this.toolbar.setPageData(this.currentPage, this.pdfDoc?.getPageCount() || 1);
      this.sidebar.setCurrentPage(closestPage - 1);
    }
  };

  private handleDeepLinkNavigation(): void {
    // Check if leaf state has subpath or URL params
    const state = this.leaf.getViewState();
    const subpath = (state.state as any)?.subpath || "";
    if (!subpath) return;

    const pageMatch = subpath.match(/page=(\d+)/);
    if (pageMatch) {
      const pageNum = parseInt(pageMatch[1], 10);
      this.scrollToPage(pageNum - 1);
    }

    const annotMatch = subpath.match(/annot=([a-zA-Z0-9_\-]+)/);
    if (annotMatch) {
      const annotId = annotMatch[1];
      window.setTimeout(() => this.pulseAnnotation(annotId), 300);
    }
  }

  scrollToPage(pageIndex: number): void {
    const pv = this.pageViews.get(pageIndex);
    if (pv && this.pagesScrollEl) {
      pv.element.scrollIntoView({ behavior: "smooth", block: "start" });
      this.currentPage = pageIndex + 1;
      this.toolbar.setPageData(this.currentPage, this.pdfDoc?.getPageCount() || 1);
    }
  }

  pulseAnnotation(annotId: string): void {
    for (const pv of this.pageViews.values()) {
      pv.pulseAnnotation(annotId);
    }
  }

  // ==================== Text Selection Floating Toolbar ====================

  private buildTextSelectionBar(): void {
    const bar = document.createElement("div");
    bar.className = "pdf-selection-action-bar is-hidden";

    // Prevent button clicks and touches from taking focus or canceling selection
    bar.addEventListener("mousedown", (e) => {
      e.preventDefault();
    });

    // Quick Color Swatches from active Scheme
    this.selectionColorsWrap = bar.createDiv({ cls: "pdf-selection-colors" });
    this.renderSelectionBarColors();

    // Action 1: Highlight Button
    const hlBtn = bar.createEl("button", { cls: "pdf-selection-action-btn", text: "Highlight" });
    hlBtn.onclick = () => {
      if (this.selectedTextInfo) {
        const pv = this.pageViews.get(this.selectedTextInfo.pageIndex);
        pv?.createHighlightFromSelection(
          this.selectedTextInfo.text,
          this.activeColor,
          this.selectedTextInfo.lines
        );
      }
      this.hideSelectionBar();
    };

    // Action 2: Underline Button
    const ulBtn = bar.createEl("button", { cls: "pdf-selection-action-btn", text: "Underline" });
    ulBtn.onclick = () => {
      if (this.selectedTextInfo) {
        const pv = this.pageViews.get(this.selectedTextInfo.pageIndex);
        pv?.createUnderlineFromSelection(
          this.selectedTextInfo.text,
          this.activeColor,
          this.selectedTextInfo.lines
        );
      }
      this.hideSelectionBar();
    };

    // Action 3: Add Sticky Note attached to selection
    const noteBtn = bar.createEl("button", { cls: "pdf-selection-action-btn", text: "Add Note" });
    noteBtn.onclick = () => {
      if (this.selectedTextInfo) {
        const pv = this.pageViews.get(this.selectedTextInfo.pageIndex);
        pv?.addAnnotatedNoteFromSelection(
          this.selectedTextInfo.text,
          this.activeColor,
          this.selectedTextInfo.lines
        );
      }
      this.hideSelectionBar();
    };

    // Action 4: Cite Button
    const citeBtn = bar.createEl("button", { cls: "pdf-selection-action-btn", text: "Cite" });
    citeBtn.onclick = () => {
      if (this.selectedTextInfo) {
        this.openCitationModal(this.selectedTextInfo.text, this.selectedTextInfo.pageIndex + 1);
      }
      this.hideSelectionBar();
    };

    // Action 5: Copy Deep Link
    const linkBtn = bar.createEl("button", { cls: "pdf-selection-action-btn", text: "Copy Link" });
    linkBtn.onclick = async () => {
      if (this.selectedTextInfo) {
        const link = CitationService.generateReferenceLink(
          this.file?.name || "document.pdf",
          this.selectedTextInfo.pageIndex + 1
        );
        await navigator.clipboard.writeText(link);
        new Notice("Copied page link to clipboard!");
      }
      this.hideSelectionBar();
    };

    document.body.appendChild(bar);
    this.selectionBarEl = bar;

    // Dismiss bar when clicking outside
    window.addEventListener("pointerdown", (e) => {
      if (this.selectionBarEl && !this.selectionBarEl.classList.contains("is-hidden")) {
        if (!this.selectionBarEl.contains(e.target as Node)) {
          this.hideSelectionBar();
        }
      }
    });
  }

  private selectedTextInfo: {
    text: string;
    pageIndex: number;
    lines?: { left: number; top: number; right: number; bottom: number }[];
  } | null = null;

  onTextSelected(
    selectedText: string,
    screenRect: DOMRect,
    pageIndex: number,
    lines?: { left: number; top: number; right: number; bottom: number }[]
  ): void {
    if (!this.selectionBarEl) return;
    this.selectedTextInfo = { text: selectedText, pageIndex, lines };

    const barW = 380;
    let left = screenRect.left + screenRect.width / 2 - barW / 2;
    let top = screenRect.top - 48;

    if (left < 16) left = 16;
    if (left + barW > window.innerWidth - 16) left = window.innerWidth - barW - 16;
    if (top < 16) top = screenRect.bottom + 8;

    this.selectionBarEl.style.left = `${Math.round(left)}px`;
    this.selectionBarEl.style.top = `${Math.round(top)}px`;
    this.selectionBarEl.removeClass("is-hidden");
  }

  private hideSelectionBar(): void {
    this.selectionBarEl?.addClass("is-hidden");
    this.selectedTextInfo = null;
  }

  // ==================== PdfToolbarDelegate Implementation ====================

  onToolChanged(tool: PdfActiveTool): void {
    this.activeTool = tool;
    for (const pv of this.pageViews.values()) {
      pv.setTool(tool, this.activeColor, this.strokeWidth);
    }
  }

  onColorChanged(color: PdfRgbColor): void {
    this.activeColor = color;
    for (const pv of this.pageViews.values()) {
      pv.setTool(this.activeTool, color, this.strokeWidth);
    }
  }

  onStrokeWidthChanged(width: number): void {
    this.strokeWidth = width;
    for (const pv of this.pageViews.values()) {
      pv.setTool(this.activeTool, this.activeColor, width);
    }
  }

  isPinching(): boolean {
    return this.isGesturePinching;
  }

  cancelActiveInteractions(): void {
    for (const pv of this.pageViews.values()) {
      pv.cancelActiveInteractions();
    }
  }

  public async applyZoomWithFocalPoint(
    newScale: number,
    focalScreenX: number,
    focalScreenY: number
  ): Promise<void> {
    if (!this.pagesScrollEl || !this.pdfDoc) return;

    const clampedScale = Math.max(0.25, Math.min(5.0, Math.round(newScale * 100) / 100));
    const oldScale = this.zoomScale;

    const scrollEl = this.pagesScrollEl;
    const scrollRect = scrollEl.getBoundingClientRect();
    const focalX = focalScreenX - scrollRect.left;
    const focalY = focalScreenY - scrollRect.top;

    let anchorPage: PdfPageView | null = null;
    let normPageX = 0;
    let normPageY = 0;

    for (const pv of this.pageViews.values()) {
      const rect = pv.element.getBoundingClientRect();
      if (
        focalScreenX >= rect.left &&
        focalScreenX <= rect.right &&
        focalScreenY >= rect.top &&
        focalScreenY <= rect.bottom
      ) {
        anchorPage = pv;
        normPageX = (focalScreenX - rect.left) / oldScale;
        normPageY = (focalScreenY - rect.top) / oldScale;
        break;
      }
    }

    if (!anchorPage) {
      const targetIndex = Math.max(0, Math.min(this.currentPage - 1, this.pageViews.size - 1));
      anchorPage = this.pageViews.get(targetIndex) || this.pageViews.get(0) || null;
      if (anchorPage) {
        const rect = anchorPage.element.getBoundingClientRect();
        const clampedX = Math.max(rect.left, Math.min(rect.right, focalScreenX));
        const clampedY = Math.max(rect.top, Math.min(rect.bottom, focalScreenY));
        normPageX = (clampedX - rect.left) / oldScale;
        normPageY = (clampedY - rect.top) / oldScale;
      }
    }

    this.zoomScale = clampedScale;
    this.toolbar.setZoom(this.zoomScale);

    await this.renderAllPages(true);

    if (anchorPage && anchorPage.element.isConnected) {
      let pageLeft = anchorPage.element.offsetLeft;
      let pageTop = anchorPage.element.offsetTop;

      if (this.pagesContainerEl && anchorPage.element.offsetParent === this.pagesContainerEl) {
        pageLeft += this.pagesContainerEl.offsetLeft;
        pageTop += this.pagesContainerEl.offsetTop;
      }

      const targetScrollX = pageLeft + normPageX * clampedScale - focalX;
      const targetScrollY = pageTop + normPageY * clampedScale - focalY;

      scrollEl.scrollLeft = Math.max(0, targetScrollX);
      scrollEl.scrollTop = Math.max(0, targetScrollY);
    }
  }

  async onZoomChanged(scale: number | "fit-width" | "fit-page"): Promise<void> {
    let targetScale = this.zoomScale;
    if (typeof scale === "number") {
      targetScale = scale;
    } else if (scale === "fit-width" && this.pagesScrollEl) {
      const containerWidth = this.pagesScrollEl.clientWidth - 60;
      const firstPageWidth = this.pdfDoc?.getPageInfo(0).width || 612;
      targetScale = Math.max(0.25, Math.min(5.0, containerWidth / firstPageWidth));
    } else if (scale === "fit-page" && this.pagesScrollEl) {
      const containerHeight = this.pagesScrollEl.clientHeight - 60;
      const firstPageHeight = this.pdfDoc?.getPageInfo(0).height || 792;
      targetScale = Math.max(0.25, Math.min(5.0, containerHeight / firstPageHeight));
    }

    if (this.pagesScrollEl) {
      const rect = this.pagesScrollEl.getBoundingClientRect();
      const focalX = rect.left + this.pagesScrollEl.clientWidth / 2;
      const focalY = rect.top + this.pagesScrollEl.clientHeight / 2;
      await this.applyZoomWithFocalPoint(targetScale, focalX, focalY);
    } else {
      this.zoomScale = targetScale;
      this.toolbar.setZoom(this.zoomScale);
      await this.renderAllPages(true);
    }
  }

  // ==================== 2-Finger Zoom & Pinch Handlers ====================

  private onWheel = (e: WheelEvent): void => {
    if (!e.ctrlKey && !e.metaKey) {
      return;
    }

    e.preventDefault();
    e.stopPropagation();

    this.hideSelectionBar();
    this.notePopover.hide();

    let delta = e.deltaY;
    if (e.deltaMode === 1) {
      delta *= 16;
    } else if (e.deltaMode === 2) {
      delta *= 100;
    }

    const clampedDelta = Math.max(-100, Math.min(100, delta));
    const zoomFactor = Math.exp(-clampedDelta * 0.003);

    this.handleWheelZoom(zoomFactor, e.clientX, e.clientY);
  };

  private handleWheelZoom(zoomFactor: number, focalX: number, focalY: number): void {
    if (!this.pagesContainerEl || !this.pagesScrollEl) return;

    if (!this.isWheelZooming) {
      this.isWheelZooming = true;
      this.isGesturePinching = true;
      this.cancelActiveInteractions();
      this.hideSelectionBar();
      this.notePopover.hide();

      this.startWheelZoomScale = this.zoomScale;
      this.cumulativeWheelScale = 1.0;
      this.wheelFocalX = focalX;
      this.wheelFocalY = focalY;

      const rect = this.pagesContainerEl.getBoundingClientRect();
      const originX = focalX - rect.left;
      const originY = focalY - rect.top;
      this.pagesContainerEl.style.transformOrigin = `${originX}px ${originY}px`;
    }

    this.cumulativeWheelScale *= zoomFactor;
    const tentativeScale = Math.max(
      0.25,
      Math.min(5.0, this.startWheelZoomScale * this.cumulativeWheelScale)
    );
    const effectiveRatio = tentativeScale / this.startWheelZoomScale;

    this.pagesContainerEl.style.transform = `scale(${effectiveRatio})`;

    if (this.wheelZoomDebounceTimer) {
      window.clearTimeout(this.wheelZoomDebounceTimer);
    }

    this.wheelZoomDebounceTimer = window.setTimeout(async () => {
      this.isWheelZooming = false;
      this.isGesturePinching = false;

      if (this.pagesContainerEl) {
        this.pagesContainerEl.style.transform = "none";
        this.pagesContainerEl.style.transformOrigin = "";
      }

      const finalScale = Math.round(this.startWheelZoomScale * this.cumulativeWheelScale * 100) / 100;
      await this.applyZoomWithFocalPoint(finalScale, this.wheelFocalX, this.wheelFocalY);
    }, 130);
  }

  private onTouchStart = (e: TouchEvent): void => {
    if (e.touches.length === 2) {
      this.isTouchPinching = true;
      this.isGesturePinching = true;
      if (this.touchCooldownTimer) {
        window.clearTimeout(this.touchCooldownTimer);
        this.touchCooldownTimer = null;
      }

      this.cancelActiveInteractions();
      this.hideSelectionBar();
      this.notePopover.hide();

      const t1 = e.touches[0];
      const t2 = e.touches[1];
      this.touchStartDist = Math.hypot(t1.clientX - t2.clientX, t1.clientY - t2.clientY);
      this.touchStartMidX = (t1.clientX + t2.clientX) / 2;
      this.touchStartMidY = (t1.clientY + t2.clientY) / 2;
      this.touchCurrentMidX = this.touchStartMidX;
      this.touchCurrentMidY = this.touchStartMidY;
      this.touchStartScale = this.zoomScale;
      this.touchCurrentRatio = 1.0;

      if (this.pagesContainerEl) {
        const rect = this.pagesContainerEl.getBoundingClientRect();
        const originX = this.touchStartMidX - rect.left;
        const originY = this.touchStartMidY - rect.top;
        this.pagesContainerEl.style.transformOrigin = `${originX}px ${originY}px`;
      }
    } else if (e.touches.length > 2) {
      this.isGesturePinching = true;
      this.cancelActiveInteractions();
    }
  };

  private onTouchMove = (e: TouchEvent): void => {
    if (!this.isTouchPinching || e.touches.length !== 2) return;

    e.preventDefault();

    const t1 = e.touches[0];
    const t2 = e.touches[1];
    const currentDist = Math.hypot(t1.clientX - t2.clientX, t1.clientY - t2.clientY);
    if (this.touchStartDist <= 0) return;

    this.touchCurrentMidX = (t1.clientX + t2.clientX) / 2;
    this.touchCurrentMidY = (t1.clientY + t2.clientY) / 2;

    const rawRatio = currentDist / this.touchStartDist;
    const tentativeScale = Math.max(0.25, Math.min(5.0, this.touchStartScale * rawRatio));
    this.touchCurrentRatio = tentativeScale / this.touchStartScale;

    const panX = this.touchCurrentMidX - this.touchStartMidX;
    const panY = this.touchCurrentMidY - this.touchStartMidY;

    if (this.pagesContainerEl) {
      this.pagesContainerEl.style.transform = `translate3d(${panX}px, ${panY}px, 0) scale(${this.touchCurrentRatio})`;
    }
  };

  private onTouchEnd = async (e: TouchEvent): Promise<void> => {
    if (!this.isTouchPinching) return;

    if (e.touches.length < 2) {
      this.isTouchPinching = false;

      if (this.pagesContainerEl) {
        this.pagesContainerEl.style.transform = "none";
        this.pagesContainerEl.style.transformOrigin = "";
      }

      const finalScale = Math.round(this.touchStartScale * this.touchCurrentRatio * 100) / 100;
      await this.applyZoomWithFocalPoint(finalScale, this.touchCurrentMidX, this.touchCurrentMidY);

      if (this.touchCooldownTimer) window.clearTimeout(this.touchCooldownTimer);
      this.touchCooldownTimer = window.setTimeout(() => {
        this.isGesturePinching = false;
        this.touchCooldownTimer = null;
      }, 150);
    }
  };

  private onTouchCancel = (e: TouchEvent): void => {
    void this.onTouchEnd(e);
  };

  private onGestureStart = (e: any): void => {
    e.preventDefault();
    this.isGesturePinching = true;
    this.cancelActiveInteractions();
    this.hideSelectionBar();
    this.notePopover.hide();

    this.touchStartScale = this.zoomScale;
    this.touchCurrentRatio = 1.0;
    this.touchStartMidX = e.clientX ?? ((this.pagesScrollEl?.clientWidth ?? 0) / 2);
    this.touchStartMidY = e.clientY ?? ((this.pagesScrollEl?.clientHeight ?? 0) / 2);

    if (this.pagesContainerEl) {
      const rect = this.pagesContainerEl.getBoundingClientRect();
      const originX = this.touchStartMidX - rect.left;
      const originY = this.touchStartMidY - rect.top;
      this.pagesContainerEl.style.transformOrigin = `${originX}px ${originY}px`;
    }
  };

  private onGestureChange = (e: any): void => {
    e.preventDefault();
    const scale = e.scale ?? 1.0;
    const tentativeScale = Math.max(0.25, Math.min(5.0, this.touchStartScale * scale));
    this.touchCurrentRatio = tentativeScale / this.touchStartScale;

    if (this.pagesContainerEl) {
      this.pagesContainerEl.style.transform = `scale(${this.touchCurrentRatio})`;
    }
  };

  private onGestureEnd = async (e: any): Promise<void> => {
    e.preventDefault();
    if (this.pagesContainerEl) {
      this.pagesContainerEl.style.transform = "none";
      this.pagesContainerEl.style.transformOrigin = "";
    }
    const finalScale = Math.round(this.touchStartScale * this.touchCurrentRatio * 100) / 100;
    await this.applyZoomWithFocalPoint(finalScale, this.touchStartMidX, this.touchStartMidY);

    if (this.touchCooldownTimer) window.clearTimeout(this.touchCooldownTimer);
    this.touchCooldownTimer = window.setTimeout(() => {
      this.isGesturePinching = false;
      this.touchCooldownTimer = null;
    }, 150);
  };

  onPageChanged(pageNumber: number): void {
    this.scrollToPage(pageNumber - 1);
  }

  async onRotatePage(): Promise<void> {
    if (!this.pdfDoc) return;
    const pageIndex = this.currentPage - 1;
    this.pdfDoc.rotatePage(pageIndex, 90);
    this.scheduleSave();
    await this.reloadDocumentRendering();
    new Notice(`Rotated page ${this.currentPage} by 90°`);
  }

  async onDuplicatePage(): Promise<void> {
    if (!this.pdfDoc) return;
    const pageIndex = this.currentPage - 1;
    await this.pdfDoc.duplicatePage(pageIndex);
    this.scheduleSave();
    await this.reloadDocumentRendering();
    new Notice(`Duplicated page ${this.currentPage}`);
  }

  async onDeletePage(): Promise<void> {
    if (!this.pdfDoc) return;
    if (this.pdfDoc.getPageCount() <= 1) {
      new Notice("Cannot delete the only page of a PDF document.");
      return;
    }
    const pageIndex = this.currentPage - 1;
    this.pdfDoc.deletePage(pageIndex);
    this.currentPage = Math.min(this.currentPage, this.pdfDoc.getPageCount());
    this.scheduleSave();
    await this.reloadDocumentRendering();
    new Notice(`Deleted page ${pageIndex + 1}`);
  }

  async onInsertBlankPage(): Promise<void> {
    if (!this.pdfDoc) return;
    const pageIndex = this.currentPage;
    this.pdfDoc.insertBlankPage(pageIndex);
    this.scheduleSave();
    await this.reloadDocumentRendering();
    new Notice(`Inserted blank page at position ${pageIndex + 1}`);
  }

  onToggleSidebar(): void {
    const isVisible = this.sidebar.toggle();
    this.toolbar.setSidebarActive(isVisible);
    if (isVisible) {
      this.retractTabEl?.addClass("is-hidden");
    } else {
      this.retractTabEl?.removeClass("is-hidden");
    }
  }

  onCite(): void {
    this.openCitationModal(undefined, this.currentPage);
  }

  async onCopyReference(): Promise<void> {
    const link = CitationService.generateReferenceLink(this.file?.name || "document.pdf", this.currentPage);
    await navigator.clipboard.writeText(link);
    new Notice(`Copied reference link: ${link}`);
  }

  async onExportMarkdown(): Promise<void> {
    if (!this.pdfDoc || !this.file) return;
    const annots = this.pdfDoc.getAllAnnotations();
    if (annots.length === 0) {
      new Notice("No annotations found in this PDF to export.");
      return;
    }

    const md = CitationService.exportAnnotationsToMarkdown(
      this.file.name,
      this.pdfDoc.getTitle(),
      annots,
      this.pdfDoc.getPageCount(),
      this.activeColorScheme
    );

    const parentPath = this.file.parent?.path && this.file.parent.path !== "/" ? this.file.parent.path : "";
    const exportPath = parentPath
      ? `${parentPath}/Annotations - ${this.file.basename}.md`
      : `Annotations - ${this.file.basename}.md`;

    try {
      let exportFile = this.app.vault.getAbstractFileByPath(exportPath);
      if (exportFile instanceof TFile) {
        await this.app.vault.modify(exportFile, md);
      } else {
        exportFile = await this.app.vault.create(exportPath, md);
      }
      new Notice(`Exported ${annots.length} annotations to "${exportFile.name}"`);
      await this.app.workspace.openLinkText(exportPath, "", "split");
    } catch (err) {
      console.error("Failed to export annotations:", err);
      new Notice(`Failed to export annotations: ${(err as Error).message}`);
    }
  }

  onUndo(): void {
    // History stack undo
    new Notice("Undo action");
  }

  onRedo(): void {
    // History stack redo
    new Notice("Redo action");
  }

  onThemeChanged(theme: "light" | "sepia" | "dark"): void {
    this.readerTheme = theme;
    this.paperTheme = theme === "dark" ? "dark" : "light";
    this.contentEl.removeClass("pdf-theme-light");
    this.contentEl.removeClass("pdf-theme-sepia");
    this.contentEl.removeClass("pdf-theme-dark");
    this.contentEl.addClass(`pdf-theme-${theme}`);
  }

  onToggleTheme(): void {
    const next = this.readerTheme === "light" ? "sepia" : this.readerTheme === "sepia" ? "dark" : "light";
    this.onThemeChanged(next);
  }

  async onSave(): Promise<void> {
    await this.saveNow();
  }

  // ==================== Color Scheme & Tag Delegate Methods ====================

  initColorSchemeForFile(filePath: string): void {
    // 1. Embedded scheme in PDF catalog
    const embeddedScheme = this.pdfDoc?.getColorScheme();
    if (embeddedScheme) {
      this.activeColorScheme = embeddedScheme;
      return;
    }

    // 2. File-specific scheme in plugin settings
    const plugin = this.getPlugin();
    if (plugin && plugin.settings.pdfFileSchemes[filePath]) {
      this.activeColorScheme = plugin.settings.pdfFileSchemes[filePath];
      return;
    }

    // 3. Default template profile from plugin settings
    if (plugin) {
      const defaultProfile =
        plugin.settings.pdfProfiles.find((p) => p.id === plugin.settings.pdfDefaultProfileId) ||
        plugin.settings.pdfProfiles[0];
      if (defaultProfile) {
        this.activeColorScheme = cloneProfile(defaultProfile);
        return;
      }
    }

    // 4. Fallback to built-in Zotero Default
    this.activeColorScheme = cloneProfile(ZOTERO_DEFAULT_PROFILE);
  }

  private renderSelectionBarColors(): void {
    if (!this.selectionColorsWrap) return;
    this.selectionColorsWrap.empty();

    const colors =
      this.activeColorScheme.colors && this.activeColorScheme.colors.length > 0
        ? this.activeColorScheme.colors
        : PDF_COLOR_PRESETS.map((p) => ({
            id: p.name.toLowerCase(),
            hex: p.hex,
            name: p.name,
            tag: "",
            rgb: p.rgb,
          }));

    for (const item of colors) {
      const dot = this.selectionColorsWrap.createSpan({ cls: "pdf-selection-color-dot" });
      dot.style.backgroundColor = item.hex;
      const tagText = item.tag ? ` (${item.tag})` : "";
      dot.title = `${item.name}${tagText}`;

      if (rgbToHex(this.activeColor).toLowerCase() === normalizeHexColor(item.hex)) {
        dot.addClass("is-active");
      }

      dot.onclick = (e) => {
        e.stopPropagation();
        this.activeColor = item.rgb || hexToRgb(item.hex);
        this.selectionColorsWrap
          ?.querySelectorAll(".pdf-selection-color-dot")
          .forEach((d) => d.removeClass("is-active"));
        dot.addClass("is-active");
        this.toolbar.setActiveTool(this.activeTool);
        if (this.selectedTextInfo) {
          const pv = this.pageViews.get(this.selectedTextInfo.pageIndex);
          pv?.createHighlightFromSelection(
            this.selectedTextInfo.text,
            this.activeColor,
            this.selectedTextInfo.lines
          );
          this.hideSelectionBar();
        }
      };
    }
  }

  onSelectProfile(profileId: string): void {
    const plugin = this.getPlugin();
    const profile = plugin?.settings.pdfProfiles.find((p) => p.id === profileId);
    if (!profile) return;

    this.applyColorScheme(cloneProfile(profile, undefined, profile.name), true);
    new Notice(`Applied color scheme template "${profile.name}"`);
  }

  onOpenColorSchemeModal(): void {
    const plugin = this.getPlugin();
    const available = plugin ? plugin.settings.pdfProfiles : [this.activeColorScheme];

    new PdfColorSchemeModal(this.app, {
      fileName: this.file?.name || "document.pdf",
      currentScheme: this.activeColorScheme,
      availableProfiles: available,
      onApply: (newScheme) => {
        this.applyColorScheme(newScheme, true);
      },
      onSaveAsTemplate: async (newProfile) => {
        if (plugin) {
          plugin.settings.pdfProfiles.push(newProfile);
          await plugin.saveSettings();
          plugin.notifyColorSchemeChanged();
        }
        this.applyColorScheme(newProfile, true);
      },
    }).open();
  }

  async onSaveSchemeAsProfile(): Promise<void> {
    new TextPromptModal(
      this.app,
      "Save as template profile",
      this.activeColorScheme.name,
      async (name) => {
        if (!name || !name.trim()) return;

        const plugin = this.getPlugin();
        if (!plugin) return;

        const newProfile = cloneProfile(this.activeColorScheme, undefined, name.trim());
        plugin.settings.pdfProfiles.push(newProfile);
        await plugin.saveSettings();
        plugin.notifyColorSchemeChanged();
        new Notice(`Saved template profile "${newProfile.name}"! Available in all PDFs.`);
      }
    ).open();
  }

  async onResetSchemeToDefault(): Promise<void> {
    const plugin = this.getPlugin();
    const defaultProfile =
      plugin?.settings.pdfProfiles.find((p) => p.id === plugin?.settings.pdfDefaultProfileId) ||
      ZOTERO_DEFAULT_PROFILE;

    this.applyColorScheme(cloneProfile(defaultProfile), true);
    new Notice(`Reset color scheme to "${defaultProfile.name}" default`);
  }

  applyColorScheme(scheme: PdfColorSchemeProfile, persist = true): void {
    this.activeColorScheme = scheme;

    // Ensure active tool color matches one in the scheme, else default to first
    const activeHex = rgbToHex(this.activeColor).toLowerCase();
    const hasColor = scheme.colors.some((c) => normalizeHexColor(c.hex) === activeHex);
    if (!hasColor && scheme.colors[0]) {
      this.activeColor = scheme.colors[0].rgb || hexToRgb(scheme.colors[0].hex);
    }

    // 1. Update Toolbar
    this.toolbar.setColorScheme(this.activeColorScheme, this.getPlugin()?.settings.pdfProfiles);

    // 2. Update Sidebar
    this.sidebar.setColorScheme(this.activeColorScheme);

    // 3. Update Note Popover
    this.notePopover.setColorScheme(this.activeColorScheme);

    // 4. Update Selection Bar
    this.renderSelectionBarColors();

    // 5. Update all Page Views
    for (const pv of this.pageViews.values()) {
      pv.setColorScheme(this.activeColorScheme);
      pv.setTool(this.activeTool, this.activeColor, this.strokeWidth);
    }

    // 6. Persist to PDF catalog and plugin settings
    if (persist) {
      if (this.pdfDoc) {
        this.pdfDoc.setColorScheme(this.activeColorScheme);
        this.scheduleSave();
      }
      if (this.file && this.getPlugin()) {
        void this.getPlugin()?.setColorSchemeForFile(this.file.path, this.activeColorScheme);
      }
    }
  }

  async onClose(): Promise<void> {
    this.unregisterColorSchemeListener?.();
    if (this.saveTimer) {
      window.clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
    await this.requestSave();
    this.toolbar.dispose();
    this.sidebar.dispose();
    this.notePopover.hide();
    this.selectionBarEl?.remove();
    this.pageViews.clear();
  }

  // ==================== PdfSidebarDelegate Implementation ====================

  onSelectPage(pageIndex: number): void {
    this.scrollToPage(pageIndex);
  }

  onSelectAnnotation(annotation: PdfAnnotation): void {
    this.scrollToAnnotation(annotation);
    this.pulseAnnotation(annotation.id);
    for (const [idx, pv] of this.pageViews.entries()) {
      if (idx === annotation.pageIndex) {
        pv.selectAnnotation(annotation.id);
      } else {
        pv.selectAnnotation(null);
      }
    }
  }

  getPageView(pageIndex: number): PdfPageView | undefined {
    return this.pageViews.get(pageIndex);
  }

  scrollToAnnotation(annotation: PdfAnnotation): void {
    const pv = this.pageViews.get(annotation.pageIndex);
    if (pv && this.pagesScrollEl) {
      const el = pv.element.querySelector(`[data-annot-id="${annotation.id}"]`) as HTMLElement;
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "center" });
      } else {
        pv.element.scrollIntoView({ behavior: "smooth", block: "start" });
      }
      this.currentPage = annotation.pageIndex + 1;
      this.toolbar.setPageData(this.currentPage, this.pdfDoc?.getPageCount() || 1);
      this.sidebar.setCurrentPage(annotation.pageIndex);
    }
  }

  onChangeAnnotationColor(id: string, color: PdfRgbColor): void {
    if (!this.pdfDoc) return;
    const annot = this.pdfDoc.getAllAnnotations().find((a) => a.id === id);
    if (!annot) return;
    annot.color = color;
    if (annot.type === "square" || annot.type === "circle" || annot.type === "line") {
      (annot as any).strokeColor = color;
    }
    this.pdfDoc.upsertAnnotation(annot);
    this.syncAnnotationsToViews();
    this.scheduleSave();
    new Notice("Annotation color updated");
  }

  onOpenLinkedNote(vaultPath: string): void {
    if (this.app?.workspace) {
      this.app.workspace.openLinkText(vaultPath, "", false);
    }
  }

  onEditAnnotation(annotation: PdfAnnotation): void {
    this.scrollToAnnotation(annotation);
    const pv = this.pageViews.get(annotation.pageIndex);
    if (pv) {
      const el = pv.element.querySelector(`[data-annot-id="${annotation.id}"]`) as HTMLElement;
      this.notePopover.show(el || pv.element, annotation, this, this.activeColorScheme);
    }
  }

  onDeleteAnnotation(id: string): void {
    if (!this.pdfDoc) return;
    this.pdfDoc.deleteAnnotation(id);
    this.syncAnnotationsToViews();
    this.scheduleSave();
    new Notice("Deleted annotation");
  }

  // ==================== PdfPageViewDelegate Implementation ====================

  onAddAnnotation(annotation: PdfAnnotation): void {
    if (!this.pdfDoc) return;
    this.pdfDoc.upsertAnnotation(annotation);
    this.syncAnnotationsToViews();
    this.scheduleSave();
  }

  onUpdateAnnotation(annotation: PdfAnnotation): void {
    if (!this.pdfDoc) return;
    this.pdfDoc.upsertAnnotation(annotation);
    this.syncAnnotationsToViews();
    this.scheduleSave();
  }

  onClickAnnotation(annotation: PdfAnnotation, element: HTMLElement | SVGElement): void {
    this.notePopover.show(element, annotation, this, this.activeColorScheme);
  }

  private syncAnnotationsToViews(): void {
    if (!this.pdfDoc) return;
    for (const [idx, pv] of this.pageViews.entries()) {
      pv.setAnnotations(this.pdfDoc.getPageAnnotations(idx));
    }
    this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
  }

  // ==================== Document Reloading after Page Operations ====================

  private async getPdfJsLibrary(): Promise<any> {
    if (this.pdfjsLib) {
      return this.pdfjsLib;
    }
    if ((window as any).pdfjsLib) {
      this.pdfjsLib = (window as any).pdfjsLib;
      return this.pdfjsLib;
    }
    if (typeof loadPdfJs === "function") {
      try {
        const lib = await loadPdfJs();
        if (lib) {
          this.pdfjsLib = lib;
          (window as any).pdfjsLib = lib;
          return lib;
        }
      } catch (e) {
        console.warn("loadPdfJs failed, falling back to window.pdfjsLib", e);
      }
    }
    if ((window as any).pdfjsLib) {
      this.pdfjsLib = (window as any).pdfjsLib;
      return this.pdfjsLib;
    }
    throw new Error("PDF.js engine is not available in this Obsidian environment.");
  }

  private async reloadDocumentRendering(): Promise<void> {
    if (!this.pdfDoc) return;
    const bytes = await this.pdfDoc.getBytes();
    const pdfjsLib = await this.getPdfJsLibrary();
    if (this.pdfjsDoc) {
      this.pdfjsDoc.destroy();
    }
    this.pdfjsDoc = await pdfjsLib.getDocument({
      data: new Uint8Array(bytes.slice(0)),
    }).promise;

    this.toolbar.setPageData(this.currentPage, this.pdfDoc.getPageCount());
    this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
    await this.renderAllPages();
  }

  private openCitationModal(quotedText?: string, pageNumber = 1): void {
    if (!this.pdfDoc || !this.file) return;
    const meta: CitationMetadata = {
      title: this.pdfDoc.getTitle() || this.file.basename,
      author: this.pdfDoc.getAuthor(),
      year: new Date().getFullYear(),
      pdfFileName: this.file.name,
      pageNumber,
      quotedText,
    };
    new PdfCiteModal(this.app, meta).open();
  }

  // ==================== Persistence & Saving ====================

  private scheduleSave(): void {
    this.toolbar.updateSaveStatus(true);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => {
      void this.saveNow();
    }, 400);
  }

  private async saveNow(): Promise<void> {
    if (!this.pdfDoc || !this.file) return;
    if (this.isSaving) {
      this.hasPendingSave = true;
      return;
    }
    this.isSaving = true;
    this.toolbar.updateSaveStatus(true, true);
    this.isInternalSaving = true;
    this.lastInternalSaveTime = Date.now();
    try {
      const bytes = await this.pdfDoc.save();
      await this.app.vault.modifyBinary(
        this.file,
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      );
      this.lastInternalSaveTime = Date.now();
      this.toolbar.updateSaveStatus(false);
    } catch (err) {
      console.error("Failed to save PDF:", err);
      new Notice(`Failed to save PDF ${this.file.name}: ${(err as Error).message}`);
    } finally {
      this.isSaving = false;
      window.setTimeout(() => {
        this.isInternalSaving = false;
      }, 2000);
      if (this.hasPendingSave) {
        this.hasPendingSave = false;
        void this.saveNow();
      }
    }
  }

  async requestSave(): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.pdfDoc?.dirty) await this.saveNow();
  }

  /**
   * Reloads document binary from disk when modified externally (e.g. via Obsidian Sync on another computer).
   */
  private async reloadFromExternalSync(): Promise<void> {
    if (!this.file) return;
    try {
      const bytes = await this.app.vault.readBinary(this.file);
      this.pdfDoc = await PdfDocumentAdapter.load(bytes);
      const pdfjsLib = await this.getPdfJsLibrary();
      if (this.pdfjsDoc) {
        this.pdfjsDoc.destroy();
      }
      this.pdfjsDoc = await pdfjsLib.getDocument({
        data: new Uint8Array(bytes.slice(0)),
        cMapUrl: "https://unpkg.com/pdfjs-dist@3.11.174/cmaps/",
        cMapPacked: true,
      }).promise;

      await this.enrichAnnotationsFromPdfJs();

      this.toolbar.setPageData(this.currentPage, this.pdfDoc.getPageCount());
      this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
      await this.renderAllPages();
      new Notice(`PDF annotations synchronized from ${this.file.name}`);
    } catch (err) {
      console.warn("Failed to reload PDF from external sync:", err);
    }
  }

  /**
   * Scans PDF pages via PDF.js to:
   * 1. Extract underlying page text for Highlights/Underlines whose text is empty
   * 2. Discover any annotations parsed by PDF.js that may not have been in pdf-lib's initial scan
   */
  private async enrichAnnotationsFromPdfJs(): Promise<void> {
    if (!this.pdfDoc || !this.pdfjsDoc) return;

    let hasChanges = false;
    const pageCount = this.pdfDoc.getPageCount();

    for (let i = 0; i < pageCount; i++) {
      try {
        const page = await this.pdfjsDoc.getPage(i + 1);
        const [textContent, rawAnnots] = await Promise.all([
          page.getTextContent().catch(() => null),
          page.getAnnotations({ intent: "display" }).catch(() => []),
        ]);

        // 1. Discover any annotations from PDF.js not in adapter
        for (const pAnnot of rawAnnots || []) {
          let subtype = (pAnnot.subtype || "").toLowerCase();
          if (!subtype && typeof pAnnot.annotationType === "number") {
            switch (pAnnot.annotationType) {
              case 1: subtype = "text"; break;
              case 3: subtype = "freetext"; break;
              case 4: subtype = "line"; break;
              case 5: subtype = "square"; break;
              case 6: subtype = "circle"; break;
              case 9: subtype = "highlight"; break;
              case 10: subtype = "underline"; break;
              case 11: subtype = "squiggly"; break;
              case 12: subtype = "strikeout"; break;
              case 13: subtype = "stamp"; break;
              case 14: subtype = "caret"; break;
              case 15: subtype = "ink"; break;
            }
          }
          if (subtype === "popup" || subtype === "link" || subtype === "widget" || !subtype) continue;

          const existingAnnots = this.pdfDoc.getPageAnnotations(i);
          const pRect = pAnnot.rect as number[] | undefined;
          const alreadyExists = existingAnnots.some((a) => {
            if (a.id === pAnnot.id) return true;
            if (pRect && Array.isArray(pRect) && pRect.length >= 4) {
              const dx = Math.abs(a.rect[0] - pRect[0]);
              const dy = Math.abs(a.rect[1] - pRect[1]);
              return dx < 3 && dy < 3;
            }
            return false;
          });

          if (!alreadyExists && pRect && Array.isArray(pRect) && pRect.length >= 4) {
            const parsed = this.convertPdfJsAnnotation(pAnnot, i, pRect);
            if (parsed) {
              this.pdfDoc.registerExistingAnnotation(parsed);
              hasChanges = true;
            }
          }
        }

        // 2. Extract underlying text from textContent for Highlights/Underlines that have empty text
        if (textContent && textContent.items && textContent.items.length > 0) {
          const pageAnnots = this.pdfDoc.getPageAnnotations(i);
          for (const annot of pageAnnots) {
            if (annot.type === "highlight" || annot.type === "underline") {
              const hl = annot as (PdfHighlightAnnotation | PdfUnderlineAnnotation);
              if (!hl.highlightedText || hl.highlightedText.trim() === "") {
                const extracted = this.extractTextUnderAnnotation(textContent, hl.rect, hl.quadPoints);
                if (extracted && extracted.trim()) {
                  hl.highlightedText = extracted;
                  if (!hl.contents || hl.contents.trim() === "") {
                    hl.contents = extracted;
                  }
                  hasChanges = true;
                }
              }
            }
          }
        }
      } catch (err) {
        console.warn(`Error enriching annotations on page ${i + 1}:`, err);
      }
    }

    if (hasChanges) {
      this.sidebar.updateData(this.pdfDoc, this.pdfjsDoc);
      for (const [idx, pv] of this.pageViews.entries()) {
        pv.setAnnotations(this.pdfDoc.getPageAnnotations(idx));
      }
    }
  }

  private extractTextUnderAnnotation(textContent: any, rect: PdfRect, quadPoints?: number[]): string {
    const [minX, minY, maxX, maxY] = rect;
    const matchingItems: { str: string; x: number; y: number }[] = [];
    const tolerance = 2;

    for (const item of textContent.items) {
      if (!("str" in item) || !item.str.trim()) continue;

      const tx = item.transform[4];
      const ty = item.transform[5];
      const width = item.width || 0;
      const height = item.height || Math.abs(item.transform[3]) || 12;

      const itemLeft = tx;
      const itemRight = tx + width;
      const itemBottom = ty;
      const itemTop = ty + height;

      // Check overlap with highlight rect
      const overlaps = !(
        itemRight < minX - tolerance ||
        itemLeft > maxX + tolerance ||
        itemTop < minY - tolerance ||
        itemBottom > maxY + tolerance
      );

      if (overlaps) {
        matchingItems.push({ str: item.str, x: tx, y: ty });
      }
    }

    // Sort top-to-bottom (higher Y first in PDF user space), then left-to-right (lower X first)
    matchingItems.sort((a, b) => {
      const yDiff = Math.abs(a.y - b.y);
      if (yDiff > 6) {
        return b.y - a.y;
      }
      return a.x - b.x;
    });

    return matchingItems.map((i) => i.str).join(" ").replace(/\s+/g, " ").trim();
  }

  private convertPdfJsAnnotation(pAnnot: any, pageIndex: number, rect: number[]): PdfAnnotation | null {
    const id = pAnnot.id || `pdfjs_${pageIndex}_${Math.random().toString(36).slice(2, 9)}`;
    let subtype = (pAnnot.subtype || "").toLowerCase();
    if (!subtype && typeof pAnnot.annotationType === "number") {
      switch (pAnnot.annotationType) {
        case 1: subtype = "text"; break;
        case 3: subtype = "freetext"; break;
        case 4: subtype = "line"; break;
        case 5: subtype = "square"; break;
        case 6: subtype = "circle"; break;
        case 9: subtype = "highlight"; break;
        case 10: subtype = "underline"; break;
        case 11: subtype = "squiggly"; break;
        case 12: subtype = "strikeout"; break;
        case 13: subtype = "stamp"; break;
        case 14: subtype = "caret"; break;
        case 15: subtype = "ink"; break;
      }
    }
    const pdfRect: PdfRect = [rect[0], rect[1], rect[2], rect[3]];

    let color: PdfRgbColor = [0.99, 0.88, 0.28];
    if (pAnnot.color && Array.isArray(pAnnot.color) && pAnnot.color.length >= 3) {
      const c = pAnnot.color;
      color = [
        c[0] > 1 ? c[0] / 255 : c[0],
        c[1] > 1 ? c[1] / 255 : c[1],
        c[2] > 1 ? c[2] / 255 : c[2],
      ];
    }

    const contents = pAnnot.contentsObj?.str || pAnnot.contents || "";
    const author = pAnnot.titleObj?.str || pAnnot.title || "Author";

    if (subtype === "highlight" || subtype === "underline" || subtype === "strikeout" || subtype === "squiggly") {
      const quads =
        pAnnot.quadPoints && Array.isArray(pAnnot.quadPoints) && pAnnot.quadPoints.length > 0
          ? pAnnot.quadPoints
          : [pdfRect[0], pdfRect[3], pdfRect[2], pdfRect[3], pdfRect[0], pdfRect[1], pdfRect[2], pdfRect[1]];
      if (subtype === "underline") {
        return {
          id,
          type: "underline",
          pageIndex,
          rect: pdfRect,
          color,
          contents,
          highlightedText: contents,
          quadPoints: quads,
          author,
        } as PdfUnderlineAnnotation;
      }
      return {
        id,
        type: "highlight",
        pageIndex,
        rect: pdfRect,
        color,
        contents,
        highlightedText: contents,
        quadPoints: quads,
        author,
      } as PdfHighlightAnnotation;
    }

    if (subtype === "text" || subtype === "caret" || subtype === "stamp") {
      return {
        id,
        type: "text",
        pageIndex,
        rect: pdfRect,
        color,
        contents,
        author,
        iconName: pAnnot.name || "Comment",
      } as PdfNoteAnnotation;
    }

    if (subtype === "freetext") {
      return {
        id,
        type: "freetext",
        pageIndex,
        rect: pdfRect,
        color,
        contents,
        fontSize: pAnnot.fontSize || 12,
        textColor: color,
        author,
      } as PdfFreeTextAnnotation;
    }

    if (subtype === "ink") {
      return {
        id,
        type: "ink",
        pageIndex,
        rect: pdfRect,
        color,
        contents,
        paths: pAnnot.inkLists || [],
        strokeWidth: 2,
        author,
      } as PdfInkAnnotation;
    }

    if (subtype === "square" || subtype === "circle" || subtype === "line") {
      return {
        id,
        type: subtype as any,
        pageIndex,
        rect: pdfRect,
        color,
        strokeColor: color,
        strokeWidth: 2,
        contents,
        author,
      } as PdfShapeAnnotation;
    }

    return null;
  }
}
