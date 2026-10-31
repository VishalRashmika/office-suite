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
  // CRUD Actions
  onAddSlide(): void;
  onDuplicateSlide(): void;
  onDeleteSlide(): void;
  onMoveSlideUp(): void;
  onMoveSlideDown(): void;
  onAddTextBox(): void;
  onAddShape(type: "roundRect" | "rect" | "ellipse"): void;
  onDeleteElement(): void;
  onSave(): void;
}

export class PptxToolbar {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxToolbarDelegate | null = null;

  private slideInputEl: HTMLInputElement | null = null;
  private totalSlidesEl: HTMLElement | null = null;
  private prevBtn: HTMLButtonElement | null = null;
  private nextBtn: HTMLButtonElement | null = null;
  private addSlideBtn: HTMLButtonElement | null = null;
  private dupSlideBtn: HTMLButtonElement | null = null;
  private delSlideBtn: HTMLButtonElement | null = null;
  private moveUpBtn: HTMLButtonElement | null = null;
  private moveDownBtn: HTMLButtonElement | null = null;
  private addTextBtn: HTMLButtonElement | null = null;
  private addShapeBtn: HTMLButtonElement | null = null;
  private delElementBtn: HTMLButtonElement | null = null;
  private saveStatusEl: HTMLElement | null = null;
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

    // 1. Left Group: Sidebar toggle + Slide Navigation + Slide CRUD
    const leftGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Sidebar Toggle
    this.sidebarBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn is-active",
      attr: { title: "Toggle slide thumbnails" },
    });
    setIcon(this.sidebarBtn, "sidebar-left");
    this.sidebarBtn.onclick = () => this.delegate?.onToggleSidebar();

    leftGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Prev Button
    this.prevBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Previous slide" },
    });
    setIcon(this.prevBtn, "chevron-left");
    this.prevBtn.onclick = () => this.delegate?.onPrevSlide();

    // Slide Counter Input: [ 1 ] / 5
    const counterWrap = leftGroup.createDiv({ cls: "office-pptx-counter-wrap" });
    this.slideInputEl = counterWrap.createEl("input", {
      type: "number",
      cls: "office-pptx-slide-input",
      value: "1",
      attr: { min: "1", max: "999", title: "Current slide" },
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
      attr: { title: "Next slide" },
    });
    setIcon(this.nextBtn, "chevron-right");
    this.nextBtn.onclick = () => this.delegate?.onNextSlide();

    leftGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Slide CRUD Buttons
    this.addSlideBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "New slide" },
    });
    setIcon(this.addSlideBtn, "plus");
    this.addSlideBtn.onclick = () => this.delegate?.onAddSlide();

    this.dupSlideBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Duplicate current slide" },
    });
    setIcon(this.dupSlideBtn, "copy");
    this.dupSlideBtn.onclick = () => this.delegate?.onDuplicateSlide();

    this.delSlideBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn mod-danger",
      attr: { title: "Delete current slide" },
    });
    setIcon(this.delSlideBtn, "trash-2");
    this.delSlideBtn.onclick = () => this.delegate?.onDeleteSlide();

    this.moveUpBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Move slide up" },
    });
    setIcon(this.moveUpBtn, "arrow-up");
    this.moveUpBtn.onclick = () => this.delegate?.onMoveSlideUp();

    this.moveDownBtn = leftGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Move slide down" },
    });
    setIcon(this.moveDownBtn, "arrow-down");
    this.moveDownBtn.onclick = () => this.delegate?.onMoveSlideDown();

    // 2. Center Group: Content Editing + View Modes + Present
    const centerGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Insert Text Box
    this.addTextBtn = centerGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Insert text box" },
    });
    setIcon(this.addTextBtn, "type");
    this.addTextBtn.onclick = () => this.delegate?.onAddTextBox();

    // Insert Shape
    this.addShapeBtn = centerGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Insert vector shape" },
    });
    setIcon(this.addShapeBtn, "square");
    this.addShapeBtn.onclick = () => this.delegate?.onAddShape("roundRect");

    // Delete Element
    this.delElementBtn = centerGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Delete selected element" },
    });
    setIcon(this.delElementBtn, "trash");
    this.delElementBtn.disabled = true;
    this.delElementBtn.onclick = () => this.delegate?.onDeleteElement();

    centerGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Segmented View Modes
    const modeSegment = centerGroup.createDiv({ cls: "office-pptx-mode-segment" });
    this.singleModeBtn = modeSegment.createEl("button", {
      text: "Single",
      cls: "office-pptx-mode-btn is-active",
      attr: { title: "Single slide view" },
    });
    this.singleModeBtn.onclick = () => this.delegate?.onSetViewMode("single");

    this.scrollModeBtn = modeSegment.createEl("button", {
      text: "Scroll",
      cls: "office-pptx-mode-btn",
      attr: { title: "Continuous scroll view" },
    });
    this.scrollModeBtn.onclick = () => this.delegate?.onSetViewMode("scroll");

    this.gridModeBtn = modeSegment.createEl("button", {
      text: "Grid",
      cls: "office-pptx-mode-btn",
      attr: { title: "Slide sorter grid view" },
    });
    this.gridModeBtn.onclick = () => this.delegate?.onSetViewMode("grid");

    centerGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Slideshow / Present Button
    const presentBtn = centerGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn mod-cta office-pptx-present-btn",
      attr: { title: "Start slideshow" },
    });
    const playIcon = presentBtn.createSpan({ cls: "office-pptx-btn-icon" });
    setIcon(playIcon, "play");
    presentBtn.createSpan({ text: "Present", cls: "office-pptx-btn-text" });
    presentBtn.onclick = () => this.delegate?.onStartSlideshow();

    // 3. Right Group: Transitions, Zoom, Speaker Notes, Search, Save Status, Actions
    const rightGroup = this.containerEl.createDiv({ cls: "office-pptx-toolbar-group" });

    // Transition Selector
    const transWrap = rightGroup.createDiv({ cls: "office-pptx-trans-wrap" });
    transWrap.createSpan({ text: "Trans:", cls: "office-pptx-trans-label" });

    this.transitionSelectEl = transWrap.createEl("select", {
      cls: "office-pptx-toolbar-select",
      attr: { title: "Slide transition effect" },
    });
    const autoOpt = this.transitionSelectEl.createEl("option", {
      value: "auto",
      text: "Auto",
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
      attr: { title: "Replay slide transition" },
    });
    setIcon(replayBtn, "rotate-cw");
    replayBtn.onclick = () => this.delegate?.onReplayTransition();

    rightGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Zoom Controls
    const zoomOutBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Zoom out" },
    });
    setIcon(zoomOutBtn, "zoom-out");
    zoomOutBtn.onclick = () => this.stepZoom(-25);

    this.zoomSelectEl = rightGroup.createEl("select", {
      cls: "office-pptx-toolbar-select office-pptx-zoom-select",
      attr: { title: "Zoom level" },
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
      attr: { title: "Zoom in" },
    });
    setIcon(zoomInBtn, "zoom-in");
    zoomInBtn.onclick = () => this.stepZoom(25);

    rightGroup.createDiv({ cls: "office-pptx-toolbar-sep" });

    // Speaker Notes Button
    this.notesBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "Toggle speaker notes drawer" },
    });
    setIcon(this.notesBtn, "file-text");
    this.notesBtn.onclick = () => this.delegate?.onToggleNotes();

    // Search input
    const searchWrap = rightGroup.createDiv({ cls: "office-pptx-search-wrap" });
    this.searchInputEl = searchWrap.createEl("input", {
      type: "text",
      placeholder: "Search...",
      cls: "office-pptx-search-input",
    });
    this.searchInputEl.oninput = () => {
      this.delegate?.onSearchChange(this.searchInputEl?.value || "");
    };

    // Save Status Indicator
    this.saveStatusEl = rightGroup.createDiv({
      cls: "office-pptx-save-status is-saved",
      attr: { title: "Click to save presentation immediately" },
    });
    const saveIconSpan = this.saveStatusEl.createSpan({ cls: "office-pptx-save-icon" });
    setIcon(saveIconSpan, "check");
    const saveTextSpan = this.saveStatusEl.createSpan({ text: "Saved", cls: "office-pptx-save-label" });
    this.saveStatusEl.onclick = () => this.delegate?.onSave();

    // Actions dropdown / menu
    const moreBtn = rightGroup.createEl("button", {
      cls: "office-pptx-toolbar-btn",
      attr: { title: "More options" },
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
    if (this.moveUpBtn) {
      this.moveUpBtn.disabled = currentSlide <= 1;
    }
    if (this.moveDownBtn) {
      this.moveDownBtn.disabled = currentSlide >= totalSlides;
    }
    if (this.delSlideBtn) {
      this.delSlideBtn.disabled = totalSlides <= 1;
    }
  }

  setElementSelected(hasSelection: boolean): void {
    if (this.delElementBtn) {
      this.delElementBtn.disabled = !hasSelection;
    }
  }

  setSaveStatus(status: "saved" | "saving" | "unsaved" | "error"): void {
    if (!this.saveStatusEl) return;
    this.saveStatusEl.className = `office-pptx-save-status is-${status}`;
    const iconSpan = this.saveStatusEl.querySelector(".office-pptx-save-icon") as HTMLElement;
    const labelSpan = this.saveStatusEl.querySelector(".office-pptx-save-label") as HTMLElement;

    if (status === "saved") {
      if (iconSpan) setIcon(iconSpan, "check");
      if (labelSpan) labelSpan.textContent = "Saved";
    } else if (status === "saving") {
      if (iconSpan) setIcon(iconSpan, "loader");
      if (labelSpan) labelSpan.textContent = "Saving...";
    } else if (status === "unsaved") {
      if (iconSpan) setIcon(iconSpan, "disc");
      if (labelSpan) labelSpan.textContent = "Unsaved";
    } else if (status === "error") {
      if (iconSpan) setIcon(iconSpan, "alert-circle");
      if (labelSpan) labelSpan.textContent = "Save error";
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
    menu.style.top = `${rect.bottom + 4}px`;
    menu.style.right = `${window.innerWidth - rect.right}px`;

    const exportItem = menu.createDiv({ cls: "office-pptx-dropdown-item" });
    const exportIcon = exportItem.createSpan({ cls: "office-pptx-dropdown-icon" });
    setIcon(exportIcon, "file-output");
    exportItem.createSpan({ text: "Export slides to Markdown note" });
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
    window.setTimeout(() => document.addEventListener("click", closeHandler), 10);
  }
}
