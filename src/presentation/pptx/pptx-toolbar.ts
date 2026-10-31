import { setIcon } from "obsidian";
import { TransitionType, SUPPORTED_TRANSITION_OPTIONS } from "../../domain/pptx";

export type PptxViewMode = "single" | "scroll" | "grid";

export interface PptxToolbarDelegate {
  onPrevSlide(): void;
  onNextSlide(): void;
  onGoToSlide(slideNum: number): void;
  onToggleSidebar(): void;
  onToggleNotes(): void;
  onSetViewMode(mode: PptxViewMode): void;
  onStartSlideshow(): void;
  onSetTransition(transition: TransitionType | "auto"): void;
  onReplayTransition(): void;
  onZoomChange(zoom: number | "fit"): void;
  onSearchChange(query: string): void;
  onExportMarkdown(): void;
  onCopySlideLink(): void;
  onPrint(): void;
}

export class PptxToolbar {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxToolbarDelegate | null = null;

  private slideInputEl: HTMLInputElement | null = null;
  private totalSlidesEl: HTMLElement | null = null;
  private prevBtn: HTMLButtonElement | null = null;
  private nextBtn: HTMLButtonElement | null = null;
  private singleModeBtn: HTMLButtonElement | null = null;
  private scrollModeBtn: HTMLButtonElement | null = null;
  private gridModeBtn: HTMLButtonElement | null = null;
  private zoomSelectEl: HTMLSelectElement | null = null;
  private transitionSelectEl: HTMLSelectElement | null = null;
  private notesBtn: HTMLButtonElement | null = null;
  private sidebarBtn: HTMLButtonElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;

  render(container: HTMLElement, delegate: PptxToolbarDelegate): void {
    this.containerEl = container;
    this.delegate = delegate;
    this.containerEl.empty();
    this.containerEl.addClass("office-pptx-toolbar");

    // 1. Left Group: Sidebar toggle + Slide Navigation
    const leftGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Sidebar Toggle
    this.sidebarBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn is-active",
      attr: { title: "Toggle Slide Thumbnails (Sidebar)" },
    });
    setIcon(this.sidebarBtn, "sidebar-left");
    this.sidebarBtn.onclick = () => this.delegate?.onToggleSidebar();

    leftGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Prev Button
    this.prevBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Previous Slide (Left Arrow, PageUp, Backspace)" },
    });
    setIcon(this.prevBtn, "chevron-left");
    this.prevBtn.onclick = () => this.delegate?.onPrevSlide();

    // Slide Counter Input: [ 1 ] / 5
    const counterWrap = leftGroup.createDiv({ cls: "office-pptx-counter-wrap" });
    this.slideInputEl = counterWrap.createEl("input", {
      type: "number",
      cls: "office-pptx-slide-input",
      value: "1",
      attr: { min: "1", max: "999", title: "Current Slide (Type number and press Enter)" },
    });
    this.slideInputEl.onchange = () => {
      const val = parseInt(this.slideInputEl?.value || "1", 10);
      if (!isNaN(val)) {
        this.delegate?.onGoToSlide(val);
      }
    };
    this.slideInputEl.onkeydown = (e) => {
      if (e.key === "Enter") {
        this.slideInputEl?.blur();
      }
    };

    this.totalSlidesEl = counterWrap.createSpan({
      cls: "office-pptx-total-slides",
      text: "/ 1",
    });

    // Next Button
    this.nextBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Next Slide (Right Arrow, PageDown, Space)" },
    });
    setIcon(this.nextBtn, "chevron-right");
    this.nextBtn.onclick = () => this.delegate?.onNextSlide();

    // 2. Center Group: View Modes + Slideshow
    const centerGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Segmented View Modes
    const modeSegment = centerGroup.createDiv({ cls: "office-pptx-mode-segment" });
    this.singleModeBtn = modeSegment.createEl("button", {
      text: "Single",
      cls: "office-pptx-mode-btn is-active",
      attr: { title: "Single Slide View" },
    });
    this.singleModeBtn.onclick = () => this.delegate?.onSetViewMode("single");

    this.scrollModeBtn = modeSegment.createEl("button", {
      text: "Scroll",
      cls: "office-pptx-mode-btn",
      attr: { title: "Continuous Scroll View" },
    });
    this.scrollModeBtn.onclick = () => this.delegate?.onSetViewMode("scroll");

    this.gridModeBtn = modeSegment.createEl("button", {
      text: "Grid",
      cls: "office-pptx-mode-btn",
      attr: { title: "Slide Sorter Grid View" },
    });
    this.gridModeBtn.onclick = () => this.delegate?.onSetViewMode("grid");

    centerGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Slideshow / Present Button
    const presentBtn = centerGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn mod-cta office-pptx-present-btn",
      attr: { title: "Start Slideshow / Presentation (F5)" },
    });
    const playIcon = presentBtn.createSpan({ cls: "office-pptx-btn-icon" });
    setIcon(playIcon, "play");
    presentBtn.createSpan({ text: "Present", cls: "office-pptx-btn-text" });
    presentBtn.onclick = () => this.delegate?.onStartSlideshow();

    centerGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Transition Selector
    const transWrap = centerGroup.createDiv({ cls: "office-pptx-trans-wrap" });
    transWrap.createSpan({ text: "Transition:", cls: "office-pptx-trans-label" });

    this.transitionSelectEl = transWrap.createEl("select", {
      cls: "office-pptx-toolbar-select",
      attr: { title: "Slide Transition Effect" },
    });
    const autoOpt = this.transitionSelectEl.createEl("option", {
      value: "auto",
      text: "Auto (From PPTX)",
    });
    autoOpt.selected = true;

    for (const opt of SUPPORTED_TRANSITION_OPTIONS) {
      this.transitionSelectEl.createEl("option", {
        value: opt.id,
        text: opt.label,
      });
    }

    this.transitionSelectEl.onchange = () => {
      const val = this.transitionSelectEl?.value as TransitionType | "auto";
      this.delegate?.onSetTransition(val);
    };

    // Replay Transition Button
    const replayBtn = transWrap.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Replay Slide Transition" },
    });
    setIcon(replayBtn, "rotate-cw");
    replayBtn.onclick = () => this.delegate?.onReplayTransition();

    // 3. Right Group: Zoom, Speaker Notes, Search, More Actions
    const rightGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Zoom Controls
    const zoomOutBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Zoom Out (-)" },
    });
    setIcon(zoomOutBtn, "zoom-out");
    zoomOutBtn.onclick = () => this.stepZoom(-25);

    this.zoomSelectEl = rightGroup.createEl("select", {
      cls: "office-pptx-toolbar-select office-pptx-zoom-select",
      attr: { title: "Zoom Level" },
    });
    const fitOpt = this.zoomSelectEl.createEl("option", { value: "fit", text: "Fit" });
    fitOpt.selected = true;
    this.zoomSelectEl.createEl("option", { value: "50", text: "50%" });
    this.zoomSelectEl.createEl("option", { value: "75", text: "75%" });
    this.zoomSelectEl.createEl("option", { value: "100", text: "100%" });
    this.zoomSelectEl.createEl("option", { value: "125", text: "125%" });
    this.zoomSelectEl.createEl("option", { value: "150", text: "150%" });
    this.zoomSelectEl.createEl("option", { value: "200", text: "200%" });

    this.zoomSelectEl.onchange = () => {
      const val = this.zoomSelectEl?.value;
      if (val === "fit") {
        this.delegate?.onZoomChange("fit");
      } else {
        this.delegate?.onZoomChange(parseInt(val || "100", 10));
      }
    };

    const zoomInBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Zoom In (+)" },
    });
    setIcon(zoomInBtn, "zoom-in");
    zoomInBtn.onclick = () => this.stepZoom(25);

    rightGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Speaker Notes Button
    this.notesBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Toggle Speaker Notes Drawer (S)" },
    });
    setIcon(this.notesBtn, "file-text");
    this.notesBtn.onclick = () => this.delegate?.onToggleNotes();

    // Search input
    const searchWrap = rightGroup.createDiv({ cls: "office-pptx-search-wrap" });
    this.searchInputEl = searchWrap.createEl("input", {
      type: "text",
      placeholder: "Search slides...",
      cls: "office-pptx-search-input",
    });
    this.searchInputEl.oninput = () => {
      this.delegate?.onSearchChange(this.searchInputEl?.value || "");
    };

    // Actions dropdown / menu
    const moreBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "More Options (Export, Copy Link, Print)" },
    });
    setIcon(moreBtn, "more-vertical");
    moreBtn.onclick = (e) => {
      e.stopPropagation();
      this.openActionsMenu(moreBtn);
    };
  }

  updateSlideState(currentSlide: number, totalSlides: number): void {
    if (this.slideInputEl) {
      this.slideInputEl.value = String(currentSlide);
      this.slideInputEl.max = String(totalSlides);
    }
    if (this.totalSlidesEl) {
      this.totalSlidesEl.textContent = `/ ${totalSlides}`;
    }
    if (this.prevBtn) {
      this.prevBtn.disabled = currentSlide <= 1;
    }
    if (this.nextBtn) {
      this.nextBtn.disabled = currentSlide >= totalSlides;
    }
  }

  updateViewMode(mode: PptxViewMode): void {
    this.singleModeBtn?.classList.toggle("is-active", mode === "single");
    this.scrollModeBtn?.classList.toggle("is-active", mode === "scroll");
    this.gridModeBtn?.classList.toggle("is-active", mode === "grid");
  }

  updateSidebarState(isOpen: boolean): void {
    this.sidebarBtn?.classList.toggle("is-active", isOpen);
  }

  updateNotesState(isOpen: boolean): void {
    this.notesBtn?.classList.toggle("is-active", isOpen);
  }

  private stepZoom(delta: number): void {
    if (!this.zoomSelectEl) return;
    const currentVal = this.zoomSelectEl.value;
    let numVal = currentVal === "fit" ? 100 : parseInt(currentVal, 10);
    numVal = Math.max(50, Math.min(200, numVal + delta));

    // Find nearest option or select fit
    this.zoomSelectEl.value = String(numVal);
    this.delegate?.onZoomChange(numVal);
  }

  private openActionsMenu(anchorEl: HTMLElement): void {
    const existingMenu = document.querySelector(".office-pptx-dropdown-menu");
    if (existingMenu) {
      existingMenu.remove();
      return;
    }

    const rect = anchorEl.getBoundingClientRect();
    const menu = document.body.createDiv({ cls: "office-pptx-dropdown-menu" });
    menu.style.position = "fixed";
    menu.style.top = `${rect.bottom + 4}px`;
    menu.style.right = `${window.innerWidth - rect.right}px`;
    menu.style.zIndex = "9999";

    const exportItem = menu.createDiv({ cls: "office-pptx-dropdown-item" });
    const exportIcon = exportItem.createSpan({ cls: "office-pptx-dropdown-icon" });
    setIcon(exportIcon, "file-output");
    exportItem.createSpan({ text: "Export Slides to Markdown Note" });
    exportItem.onclick = () => {
      menu.remove();
      this.delegate?.onExportMarkdown();
    };

    const linkItem = menu.createDiv({ cls: "office-pptx-dropdown-item" });
    const linkIcon = linkItem.createSpan({ cls: "office-pptx-dropdown-icon" });
    setIcon(linkIcon, "link");
    linkItem.createSpan({ text: "Copy Deep Link to Current Slide" });
    linkItem.onclick = () => {
      menu.remove();
      this.delegate?.onCopySlideLink();
    };

    const printItem = menu.createDiv({ cls: "office-pptx-dropdown-item" });
    const printIcon = printItem.createSpan({ cls: "office-pptx-dropdown-icon" });
    setIcon(printIcon, "printer");
    printItem.createSpan({ text: "Print Slides / Save as PDF" });
    printItem.onclick = () => {
      menu.remove();
      this.delegate?.onPrint();
    };

    const closeHandler = (e: MouseEvent) => {
      if (!menu.contains(e.target as Node) && e.target !== anchorEl) {
        menu.remove();
        document.removeEventListener("click", closeHandler);
      }
    };
    setTimeout(() => document.addEventListener("click", closeHandler), 10);
  }
}
