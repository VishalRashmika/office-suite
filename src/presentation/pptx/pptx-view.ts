import { FileView, TFile, WorkspaceLeaf, Notice } from "obsidian";
import { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { TransitionType, PptxPresentation } from "../../domain/pptx";
import { PptxToolbar, PptxToolbarDelegate, PptxViewMode } from "./pptx-toolbar";
import { PptxSidebar, PptxSidebarDelegate } from "./pptx-sidebar";
import { PptxStage, PptxStageDelegate } from "./pptx-stage";
import { PptxNotes, PptxNotesDelegate } from "./pptx-notes";
import { PptxSlideshow, PptxSlideshowDelegate } from "./pptx-slideshow";
import type OfficeSuitePlugin from "../../main";

export const VIEW_TYPE_PPTX = "pptx-view";

export class PptxView
  extends FileView
  implements
    PptxToolbarDelegate,
    PptxSidebarDelegate,
    PptxStageDelegate,
    PptxNotesDelegate,
    PptxSlideshowDelegate
{
  allowNoFile = false;
  pptxDoc: PptxDocument | null = null;

  private plugin?: OfficeSuitePlugin;
  private toolbar = new PptxToolbar();
  private sidebar = new PptxSidebar();
  private stage = new PptxStage();
  private notes = new PptxNotes();
  private slideshow = new PptxSlideshow();

  private activeSlideIndex = 0;
  private isSidebarOpen = true;
  private isNotesOpen = false;
  private viewMode: PptxViewMode = "single";
  private transitionOverride: TransitionType | "auto" = "auto";
  private zoomLevel: number | "fit" = "fit";

  private dirty = false;
  private saveTimer: number | null = null;

  private rootEl: HTMLElement | null = null;
  private toolbarEl: HTMLElement | null = null;
  private mainBodyEl: HTMLElement | null = null;
  private sidebarEl: HTMLElement | null = null;
  private centerAreaEl: HTMLElement | null = null;
  private stageEl: HTMLElement | null = null;
  private notesEl: HTMLElement | null = null;

  constructor(leaf: WorkspaceLeaf, plugin?: OfficeSuitePlugin) {
    super(leaf);
    this.plugin = plugin;
  }

  getViewType(): string {
    return VIEW_TYPE_PPTX;
  }

  getDisplayText(): string {
    return this.file ? this.file.basename : "PowerPoint Presentation";
  }

  getIcon(): string {
    return "presentation";
  }

  async onOpen(): Promise<void> {
    const container = this.contentEl;
    container.empty();
    container.addClass("office-pptx-view-root");

    this.rootEl = container;

    // 1. Top Toolbar
    this.toolbarEl = container.createDiv({ cls: "office-pptx-toolbar-container" });
    this.toolbar.render(this.toolbarEl, this);

    // 2. Main Workspace Body (Sidebar + Center Stage)
    this.mainBodyEl = container.createDiv({ cls: "office-pptx-main-body" });

    // Left Sidebar (Thumbnails)
    this.sidebarEl = this.mainBodyEl.createDiv({ cls: "office-pptx-sidebar-wrap" });

    // Center Stage + Notes Drawer Wrap
    this.centerAreaEl = this.mainBodyEl.createDiv({ cls: "office-pptx-center-area" });
    this.stageEl = this.centerAreaEl.createDiv({ cls: "office-pptx-stage-host" });
    this.notesEl = this.centerAreaEl.createDiv({ cls: "office-pptx-notes-host is-hidden" });
    this.notes.render(this.notesEl, this);

    // Keyboard Shortcuts inside view
    this.registerDomEvent(container, "keydown", (e: KeyboardEvent) => {
      // Don't intercept if typing in an input, textarea, or contenteditable
      const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      const isEditable =
        targetTag === "input" ||
        targetTag === "textarea" ||
        targetTag === "select" ||
        (e.target as HTMLElement)?.isContentEditable;
      if (isEditable) {
        return;
      }

      if (e.key === "Delete" || e.key === "Backspace") {
        if (this.stage.getSelectedElementId() && !this.stage.isEditing()) {
          e.preventDefault();
          this.onDeleteElement();
          return;
        }
      }

      if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") {
        e.preventDefault();
        this.onNextSlide();
      } else if (e.key === "ArrowLeft" || e.key === "PageUp" || e.key === "Backspace") {
        e.preventDefault();
        this.onPrevSlide();
      } else if (e.key === "Home") {
        e.preventDefault();
        this.onGoToSlide(1);
      } else if (e.key === "End") {
        if (this.pptxDoc) {
          this.onGoToSlide(this.pptxDoc.getSlideCount());
        }
      } else if (e.key === "F5") {
        e.preventDefault();
        this.onStartSlideshow();
      } else if (e.key === "s" || e.key === "S") {
        e.preventDefault();
        this.onToggleNotes();
      }
    });
  }

  async onLoadFile(file: TFile): Promise<void> {
    try {
      const data = await this.app.vault.readBinary(file);
      this.pptxDoc = await PptxDocument.load(data);
      const pres = this.pptxDoc.getPresentation();

      this.activeSlideIndex = 0;
      this.dirty = false;

      // Check URL deep link / subpath (e.g. #slide=3)
      const state = this.leaf.getViewState()?.state as Record<string, unknown> | undefined;
      const subpath = typeof state?.subpath === "string" ? state.subpath : undefined;
      if (subpath && subpath.includes("slide=")) {
        const match = subpath.match(/slide=(\d+)/);
        if (match) {
          const targetSlide = parseInt(match[1], 10) - 1;
          if (targetSlide >= 0 && targetSlide < pres.slides.length) {
            this.activeSlideIndex = targetSlide;
          }
        }
      }

      // Render components
      if (this.sidebarEl) {
        this.sidebar.render(this.sidebarEl, pres, this.activeSlideIndex, this);
      }

      if (this.stageEl) {
        this.stage.render(
          this.stageEl,
          pres,
          this.activeSlideIndex,
          this.viewMode,
          this
        );
      }

      this.toolbar.updateSlideState(this.activeSlideIndex + 1, pres.slides.length);
      this.toolbar.updateViewMode(this.viewMode);
      this.toolbar.updateSidebarState(this.isSidebarOpen);
      this.toolbar.updateNotesState(this.isNotesOpen);
      this.toolbar.setSaveStatus("saved");

      this.updateNotesPanel();
    } catch (err: any) {
      console.error("Failed to load PowerPoint file:", err);
      new Notice(`Failed to load PPTX presentation: ${err?.message || err}`);
      if (this.stageEl) {
        this.stageEl.empty();
        const errBox = this.stageEl.createDiv({ cls: "office-pptx-error-box" });
        errBox.createEl("h3", { text: "Could not open presentation" });
        errBox.createEl("p", { text: err?.message || String(err) });
      }
    }
  }

  async onUnloadFile(file: TFile): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    return super.onUnloadFile(file);
  }

  async onClose(): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    this.stage.destroy();
    this.slideshow.stop();
  }

  onResize(): void {
    if (this.stageEl && this.viewMode === "single") {
      this.stage.setZoom(this.zoomLevel);
    }
  }

  // -------------------------------------------------------------
  // AUTOSAVE & PERSISTENCE
  // -------------------------------------------------------------

  private requestSave(): void {
    this.dirty = true;
    this.toolbar.setSaveStatus("unsaved");
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => {
      void this.saveNow();
    }, 800);
  }

  private async saveNow(): Promise<void> {
    if (!this.dirty || !this.pptxDoc || !this.file) return;
    try {
      this.toolbar.setSaveStatus("saving");
      const bytes = await this.pptxDoc.save();
      const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
      await this.app.vault.modifyBinary(this.file, arrayBuffer);
      this.dirty = false;
      this.toolbar.setSaveStatus("saved");
    } catch (err: any) {
      console.error("Failed to save PowerPoint presentation:", err);
      new Notice(`Failed to save presentation: ${err?.message || err}`);
      this.toolbar.setSaveStatus("error");
    }
  }

  onSave(): void {
    void this.saveNow();
  }

  // -------------------------------------------------------------
  // TOOLBAR & NAVIGATION DELEGATE
  // -------------------------------------------------------------

  onPrevSlide(): void {
    if (!this.pptxDoc) return;
    if (this.activeSlideIndex > 0) {
      this.goToSlide(this.activeSlideIndex - 1);
    }
  }

  onNextSlide(): void {
    if (!this.pptxDoc) return;
    if (this.activeSlideIndex < this.pptxDoc.getSlideCount() - 1) {
      this.goToSlide(this.activeSlideIndex + 1);
    }
  }

  onGoToSlide(slideNum: number): void {
    if (!this.pptxDoc) return;
    const target = Math.max(0, Math.min(this.pptxDoc.getSlideCount() - 1, slideNum - 1));
    this.goToSlide(target);
  }

  private goToSlide(index: number): void {
    this.activeSlideIndex = index;
    this.stage.goToSlide(index);
    this.sidebar.setActiveSlide(index);
    this.toolbar.updateSlideState(index + 1, this.pptxDoc?.getSlideCount() || 1);
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
    this.updateNotesPanel();
  }

  onToggleSidebar(): void {
    this.isSidebarOpen = !this.isSidebarOpen;
    if (this.sidebarEl) {
      this.sidebarEl.style.display = this.isSidebarOpen ? "" : "none";
    }
    this.toolbar.updateSidebarState(this.isSidebarOpen);
    this.onResize();
  }

  onToggleNotes(): void {
    this.isNotesOpen = !this.isNotesOpen;
    if (this.notesEl) {
      this.notesEl.toggleClass("is-hidden", !this.isNotesOpen);
    }
    this.toolbar.updateNotesState(this.isNotesOpen);
    this.updateNotesPanel();
  }

  private updateNotesPanel(): void {
    if (!this.pptxDoc) return;
    const slide = this.pptxDoc.getSlide(this.activeSlideIndex);
    this.notes.updateSlide(slide);
  }

  onSetViewMode(mode: PptxViewMode): void {
    this.viewMode = mode;
    this.stage.setViewMode(mode);
    this.toolbar.updateViewMode(mode);
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
  }

  onStartSlideshow(): void {
    if (!this.pptxDoc) return;
    this.slideshow.start(
      this.pptxDoc.getPresentation(),
      this.activeSlideIndex,
      this.transitionOverride,
      this
    );
  }

  onSetTransition(transition: TransitionType | "auto"): void {
    this.transitionOverride = transition;
    this.stage.setTransitionOverride(transition);
  }

  onReplayTransition(): void {
    this.stage.replayTransition();
  }

  onZoomChange(zoom: number | "fit"): void {
    this.zoomLevel = zoom;
    this.stage.setZoom(zoom);
  }

  onSearchChange(query: string): void {
    this.sidebar.filterSlides(query);
  }

  async onExportMarkdown(): Promise<void> {
    if (!this.pptxDoc || !this.file) return;

    try {
      const markdown = this.pptxDoc.exportMarkdown(this.file.name);
      const notePath = this.file.path.replace(/\.pptx$/i, " Notes.md");

      const existingFile = this.app.vault.getAbstractFileByPath(notePath);
      if (existingFile instanceof TFile) {
        await this.app.vault.modify(existingFile, markdown);
        new Notice(`Updated presentation notes: ${notePath}`);
      } else {
        await this.app.vault.create(notePath, markdown);
        new Notice(`Created presentation notes: ${notePath}`);
      }
    } catch (err: any) {
      new Notice(`Failed to export markdown: ${err?.message || err}`);
    }
  }

  async onCopySlideLink(): Promise<void> {
    if (!this.file) return;
    const link = `[[${this.file.name}#slide=${this.activeSlideIndex + 1}]]`;
    try {
      await navigator.clipboard.writeText(link);
      new Notice(`Copied slide link to clipboard: ${link}`);
    } catch (e) {
      new Notice(`Slide link: ${link}`);
    }
  }

  onPrint(): void {
    window.print();
  }

  // -------------------------------------------------------------
  // CRUD OPERATIONS: SLIDES & CONTENT
  // -------------------------------------------------------------

  onAddSlide(): void {
    if (!this.pptxDoc) return;
    const targetIdx = this.activeSlideIndex + 1;
    const slide = this.pptxDoc.addSlide(targetIdx);
    if (slide) {
      this.activeSlideIndex = targetIdx;
      this.refreshAll();
      this.requestSave();
      new Notice(`Added Slide ${targetIdx + 1}`);
    }
  }

  onDuplicateSlide(slideIndex?: number): void {
    if (!this.pptxDoc) return;
    const targetIdx = slideIndex ?? this.activeSlideIndex;
    const dup = this.pptxDoc.duplicateSlide(targetIdx);
    if (dup) {
      this.activeSlideIndex = targetIdx + 1;
      this.refreshAll();
      this.requestSave();
      new Notice(`Duplicated slide to Slide ${this.activeSlideIndex + 1}`);
    }
  }

  onDeleteSlide(slideIndex?: number): void {
    if (!this.pptxDoc) return;
    const targetIdx = slideIndex ?? this.activeSlideIndex;
    if (this.pptxDoc.getSlideCount() <= 1) {
      new Notice("A presentation must contain at least one slide.");
      return;
    }
    const success = this.pptxDoc.deleteSlide(targetIdx);
    if (success) {
      this.activeSlideIndex = Math.max(
        0,
        Math.min(this.activeSlideIndex, this.pptxDoc.getSlideCount() - 1)
      );
      this.refreshAll();
      this.requestSave();
      new Notice(`Deleted slide`);
    }
  }

  onMoveSlideUp(slideIndex?: number): void {
    if (!this.pptxDoc) return;
    const targetIdx = slideIndex ?? this.activeSlideIndex;
    if (targetIdx > 0) {
      this.pptxDoc.reorderSlide(targetIdx, targetIdx - 1);
      this.activeSlideIndex = targetIdx - 1;
      this.refreshAll();
      this.requestSave();
    }
  }

  onMoveSlideDown(slideIndex?: number): void {
    if (!this.pptxDoc) return;
    const targetIdx = slideIndex ?? this.activeSlideIndex;
    if (targetIdx < this.pptxDoc.getSlideCount() - 1) {
      this.pptxDoc.reorderSlide(targetIdx, targetIdx + 1);
      this.activeSlideIndex = targetIdx + 1;
      this.refreshAll();
      this.requestSave();
    }
  }

  onAddTextBox(): void {
    if (!this.pptxDoc) return;
    const el = this.pptxDoc.addElement(this.activeSlideIndex, "textbox");
    if (el) {
      this.refreshStage();
      this.refreshSidebarThumbnails();
      this.stage.selectElement(el.id);
      this.requestSave();
      new Notice("Text box added (double-click to edit text)");
    }
  }

  onAddShape(type: "roundRect" | "rect" | "ellipse"): void {
    if (!this.pptxDoc) return;
    const el = this.pptxDoc.addElement(this.activeSlideIndex, type);
    if (el) {
      this.refreshStage();
      this.refreshSidebarThumbnails();
      this.stage.selectElement(el.id);
      this.requestSave();
      new Notice(`Shape added (${type})`);
    }
  }

  onDeleteElement(slideIndex?: number, elementId?: string): void {
    if (!this.pptxDoc) return;
    const sIdx = slideIndex ?? this.activeSlideIndex;
    const elId = elementId ?? this.stage.getSelectedElementId();
    if (elId) {
      this.pptxDoc.deleteElement(sIdx, elId);
      this.stage.selectElement(null);
      this.refreshStage();
      this.refreshSidebarThumbnails();
      this.requestSave();
      new Notice("Element deleted");
    }
  }

  onElementSelected(slideIndex: number, elementId: string | null): void {
    this.toolbar.setElementSelected(Boolean(elementId));
  }

  onElementTextChange(
    slideIndex: number,
    elementId: string,
    newText: string
  ): void {
    if (!this.pptxDoc) return;
    this.pptxDoc.updateElementText(slideIndex, elementId, newText);
    this.refreshSidebarThumbnails();
    this.requestSave();
  }

  onNotesChange(slideIndex: number, newNotes: string): void {
    if (!this.pptxDoc) return;
    this.pptxDoc.updateSpeakerNotes(slideIndex, newNotes);
    this.requestSave();
  }

  // -------------------------------------------------------------
  // REFRESH HELPERS
  // -------------------------------------------------------------

  private refreshAll(): void {
    if (!this.pptxDoc) return;
    const pres = this.pptxDoc.getPresentation();
    if (this.sidebarEl) {
      this.sidebar.render(this.sidebarEl, pres, this.activeSlideIndex, this);
    }
    if (this.stageEl) {
      this.stage.render(
        this.stageEl,
        pres,
        this.activeSlideIndex,
        this.viewMode,
        this
      );
    }
    this.toolbar.updateSlideState(this.activeSlideIndex + 1, pres.slides.length);
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
    this.updateNotesPanel();
  }

  private refreshStage(): void {
    if (!this.pptxDoc) return;
    this.stage.goToSlide(this.activeSlideIndex, true);
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
  }

  private refreshSidebarThumbnails(): void {
    if (!this.pptxDoc || !this.sidebarEl) return;
    this.sidebar.render(
      this.sidebarEl,
      this.pptxDoc.getPresentation(),
      this.activeSlideIndex,
      this
    );
  }

  // -------------------------------------------------------------
  // SIDEBAR & STAGE DELEGATE EVENTS
  // -------------------------------------------------------------

  onSlideSelected(slideIndex: number): void {
    this.goToSlide(slideIndex);
  }

  onSlideChanged(newIndex: number): void {
    this.activeSlideIndex = newIndex;
    this.sidebar.setActiveSlide(newIndex);
    this.toolbar.updateSlideState(newIndex + 1, this.pptxDoc?.getSlideCount() || 1);
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
    this.updateNotesPanel();
  }

  onSelectSlide(index: number): void {
    this.activeSlideIndex = index;
    this.onSetViewMode("single");
    this.goToSlide(index);
  }

  onExit(): void {
    this.stage.goToSlide(this.activeSlideIndex, true);
    this.sidebar.setActiveSlide(this.activeSlideIndex);
    this.toolbar.updateSlideState(
      this.activeSlideIndex + 1,
      this.pptxDoc?.getSlideCount() || 1
    );
    this.toolbar.setElementSelected(Boolean(this.stage.getSelectedElementId()));
    this.updateNotesPanel();
  }
}
