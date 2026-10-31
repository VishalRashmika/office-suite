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
    this.notesEl = this.centerAreaEl.createDiv({ cls: "office-pptx-notes-host" });
    this.notesEl.style.display = "none";
    this.notes.render(this.notesEl, this);

    // Keyboard Shortcuts inside view
    this.registerDomEvent(container, "keydown", (e: KeyboardEvent) => {
      // Don't intercept if typing in an input
      const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      if (targetTag === "input" || targetTag === "textarea" || targetTag === "select") {
        return;
      }

      if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " " || e.key === "Enter") {
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

      // Check URL deep link / subpath (e.g. #slide=3)
      const subpath = this.leaf.getViewState()?.state?.subpath;
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

      this.updateNotesPanel();
    } catch (err: any) {
      console.error("Failed to load PowerPoint file:", err);
      new Notice(`Failed to load PPTX presentation: ${err?.message || err}`);
      if (this.stageEl) {
        this.stageEl.empty();
        const errBox = this.stageEl.createDiv({ cls: "office-pptx-error-box" });
        errBox.createEl("h3", { text: "Could not open PowerPoint presentation" });
        errBox.createEl("p", { text: err?.message || String(err) });
      }
    }
  }

  async onClose(): Promise<void> {
    this.stage.destroy();
    this.slideshow.stop();
  }

  onResize(): void {
    if (this.stageEl && this.viewMode === "single") {
      this.stage.setZoom(this.zoomLevel);
    }
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
      this.notesEl.style.display = this.isNotesOpen ? "" : "none";
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
  // SIDEBAR & STAGE DELEGATE EVENTS
  // -------------------------------------------------------------

  onSlideSelected(slideIndex: number): void {
    this.goToSlide(slideIndex);
  }

  onSlideChanged(newIndex: number): void {
    this.activeSlideIndex = newIndex;
    this.sidebar.setActiveSlide(newIndex);
    this.toolbar.updateSlideState(newIndex + 1, this.pptxDoc?.getSlideCount() || 1);
    this.updateNotesPanel();
  }

  onSelectSlide(index: number): void {
    // When clicking in grid mode, switch to single slide mode
    this.activeSlideIndex = index;
    this.onSetViewMode("single");
    this.goToSlide(index);
  }

  onExit(): void {
    // Slideshow exited, sync current active slide
    this.stage.goToSlide(this.activeSlideIndex, true);
    this.sidebar.setActiveSlide(this.activeSlideIndex);
    this.toolbar.updateSlideState(this.activeSlideIndex + 1, this.pptxDoc?.getSlideCount() || 1);
    this.updateNotesPanel();
  }
}
