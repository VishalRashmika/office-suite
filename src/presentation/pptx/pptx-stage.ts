import {
  PptxPresentation,
  PptxSlide,
  PptxRenderer,
  TransitionType,
  PptxTransition,
} from "../../domain/pptx";
import { PptxViewMode } from "./pptx-toolbar";

export interface PptxStageDelegate {
  onSlideChanged(newIndex: number): void;
  onSelectSlide(index: number): void;
}

export class PptxStage {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxStageDelegate | null = null;
  private presentation: PptxPresentation | null = null;
  private activeSlideIndex = 0;
  private viewMode: PptxViewMode = "single";
  private zoomLevel: number | "fit" = "fit";
  private transitionOverride: TransitionType | "auto" = "auto";
  private isTransitioning = false;

  private stageInnerEl: HTMLElement | null = null;
  private currentSlideEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;

  render(
    container: HTMLElement,
    presentation: PptxPresentation,
    activeSlideIndex: number,
    viewMode: PptxViewMode,
    delegate: PptxStageDelegate
  ): void {
    this.containerEl = container;
    this.presentation = presentation;
    this.activeSlideIndex = activeSlideIndex;
    this.viewMode = viewMode;
    this.delegate = delegate;

    this.containerEl.empty();
    this.containerEl.addClass("office-pptx-stage-container");

    this.stageInnerEl = this.containerEl.createDiv({ cls: "office-pptx-stage-inner" });

    // Setup ResizeObserver for responsive scaling
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
    }
    this.resizeObserver = new ResizeObserver(() => {
      if (this.viewMode === "single") {
        this.updateScale();
      }
    });
    this.resizeObserver.observe(this.containerEl);

    // Initial render based on mode
    this.refreshView();
  }

  destroy(): void {
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }
  }

  setViewMode(mode: PptxViewMode): void {
    this.viewMode = mode;
    this.refreshView();
  }

  setZoom(zoom: number | "fit"): void {
    this.zoomLevel = zoom;
    if (this.viewMode === "single") {
      this.updateScale();
    }
  }

  setTransitionOverride(t: TransitionType | "auto"): void {
    this.transitionOverride = t;
  }

  goToSlide(newIndex: number, forceInstant = false): void {
    if (!this.presentation) return;
    if (newIndex < 0 || newIndex >= this.presentation.slides.length) return;
    if (newIndex === this.activeSlideIndex && !forceInstant) return;

    const oldIndex = this.activeSlideIndex;
    this.activeSlideIndex = newIndex;

    if (this.viewMode === "single") {
      if (forceInstant || this.isTransitioning) {
        this.renderSingleSlide(newIndex);
      } else {
        this.executeTransition(oldIndex, newIndex);
      }
    } else if (this.viewMode === "scroll") {
      this.scrollToSlide(newIndex);
    } else if (this.viewMode === "grid") {
      this.highlightGridCard(newIndex);
    }

    this.delegate?.onSlideChanged(newIndex);
  }

  replayTransition(): void {
    if (this.viewMode === "single" && this.presentation) {
      this.executeTransition(this.activeSlideIndex, this.activeSlideIndex);
    }
  }

  private refreshView(): void {
    if (!this.stageInnerEl || !this.presentation) return;
    this.stageInnerEl.empty();
    this.stageInnerEl.className = `office-pptx-stage-inner mode-${this.viewMode}`;

    if (this.viewMode === "single") {
      this.renderSingleSlide(this.activeSlideIndex);
    } else if (this.viewMode === "scroll") {
      this.renderContinuousScroll();
    } else if (this.viewMode === "grid") {
      this.renderSlideSorterGrid();
    }
  }

  // -------------------------------------------------------------
  // SINGLE SLIDE MODE WITH TRANSITIONS & AUTO-FIT
  // -------------------------------------------------------------

  private renderSingleSlide(index: number): void {
    if (!this.stageInnerEl || !this.presentation) return;
    this.stageInnerEl.empty();

    const slide = this.presentation.slides[index];
    if (!slide) return;

    // Viewport wrapper for scaling
    const viewportWrapper = this.stageInnerEl.createDiv({ cls: "office-pptx-viewport-wrapper" });

    // Transition host
    const transHost = viewportWrapper.createDiv({ cls: "office-pptx-transition-host" });

    // Render slide canvas
    const slideNode = PptxRenderer.renderSlide(slide, this.presentation, {
      showSlideNumberBadge: true,
      totalSlides: this.presentation.slides.length,
    });
    slideNode.addClass("office-pptx-active-slide");
    transHost.appendChild(slideNode);
    this.currentSlideEl = slideNode;

    // Hover navigation overlay arrows
    const leftArrow = viewportWrapper.createDiv({
      cls: "office-pptx-nav-overlay nav-left",
      attr: { title: "Previous Slide" },
    });
    leftArrow.createSpan({ cls: "office-pptx-nav-icon", text: "❮" });
    leftArrow.onclick = (e) => {
      e.stopPropagation();
      this.goToSlide(this.activeSlideIndex - 1);
    };

    const rightArrow = viewportWrapper.createDiv({
      cls: "office-pptx-nav-overlay nav-right",
      attr: { title: "Next Slide" },
    });
    rightArrow.createSpan({ cls: "office-pptx-nav-icon", text: "❯" });
    rightArrow.onclick = (e) => {
      e.stopPropagation();
      this.goToSlide(this.activeSlideIndex + 1);
    };

    this.updateScale();
  }

  private executeTransition(fromIndex: number, toIndex: number): void {
    if (!this.stageInnerEl || !this.presentation) return;
    const transHost = this.stageInnerEl.querySelector(".office-pptx-transition-host") as HTMLElement;
    if (!transHost) {
      this.renderSingleSlide(toIndex);
      return;
    }

    const toSlide = this.presentation.slides[toIndex];
    if (!toSlide) return;

    this.isTransitioning = true;

    // Determine transition type and direction
    let transType: TransitionType = "fade";
    let transDir = "l";
    let duration = 400;

    if (this.transitionOverride !== "auto") {
      transType = this.transitionOverride;
    } else if (toSlide.transition) {
      transType = toSlide.transition.type;
      transDir = toSlide.transition.direction || "l";
      duration = toSlide.transition.durationMs || 400;
    }

    if (transType === "none" || transType === "cut") {
      this.renderSingleSlide(toIndex);
      this.isTransitioning = false;
      return;
    }

    // Auto direction based on slide progression if not specified
    if (transType === "push" && !toSlide.transition?.direction) {
      transDir = toIndex >= fromIndex ? "l" : "r";
    }

    // Existing slide element
    const oldSlideEl = transHost.querySelector(".office-pptx-active-slide") as HTMLElement;

    // New slide element
    const newSlideEl = PptxRenderer.renderSlide(toSlide, this.presentation, {
      showSlideNumberBadge: true,
      totalSlides: this.presentation.slides.length,
    });
    newSlideEl.addClass("office-pptx-active-slide");

    // Apply animation classes
    const animOutClass = `pptx-trans-${transType}-out${transType === "push" ? `-${transDir}` : ""}`;
    const animInClass = `pptx-trans-${transType}-in${transType === "push" ? `-${transDir}` : ""}`;

    if (oldSlideEl) {
      oldSlideEl.style.animationDuration = `${duration}ms`;
      oldSlideEl.classList.add(animOutClass);
    }

    newSlideEl.style.animationDuration = `${duration}ms`;
    newSlideEl.classList.add(animInClass);
    transHost.appendChild(newSlideEl);

    const finishTransition = () => {
      if (oldSlideEl && oldSlideEl.parentElement === transHost) {
        oldSlideEl.remove();
      }
      newSlideEl.classList.remove(animInClass);
      newSlideEl.style.animationDuration = "";
      this.currentSlideEl = newSlideEl;
      this.isTransitioning = false;
    };

    setTimeout(finishTransition, duration + 50);
  }

  private updateScale(): void {
    if (!this.containerEl || !this.presentation) return;
    const transHost = this.containerEl.querySelector(".office-pptx-transition-host") as HTMLElement;
    if (!transHost) return;

    const presW = this.presentation.width || 960;
    const presH = this.presentation.height || 540;

    const containerW = this.containerEl.clientWidth - 48; // padding
    const containerH = this.containerEl.clientHeight - 48;

    let scale = 1;

    if (this.zoomLevel === "fit") {
      scale = Math.min(containerW / presW, containerH / presH);
      // Reasonable scale boundaries
      scale = Math.max(0.2, Math.min(2.5, scale));
    } else {
      scale = this.zoomLevel / 100;
    }

    transHost.style.transform = `scale(${scale})`;
    transHost.style.transformOrigin = "center center";
  }

  // -------------------------------------------------------------
  // CONTINUOUS SCROLL MODE
  // -------------------------------------------------------------

  private renderContinuousScroll(): void {
    if (!this.stageInnerEl || !this.presentation) return;

    const scrollList = this.stageInnerEl.createDiv({ cls: "office-pptx-scroll-list" });

    for (let i = 0; i < this.presentation.slides.length; i++) {
      const slide = this.presentation.slides[i];
      const card = scrollList.createDiv({ cls: "office-pptx-scroll-card" });
      card.dataset.slideIndex = String(i);

      // Card Header
      const header = card.createDiv({ cls: "office-pptx-scroll-header" });
      header.createSpan({
        cls: "office-pptx-scroll-num-badge",
        text: `Slide ${slide.slideNumber}`,
      });
      if (slide.title) {
        header.createSpan({ cls: "office-pptx-scroll-title", text: slide.title });
      }

      // Slide canvas
      const slideNode = PptxRenderer.renderSlide(slide, this.presentation, {
        showSlideNumberBadge: true,
        totalSlides: this.presentation.slides.length,
      });
      card.appendChild(slideNode);
    }

    // Scroll to active slide
    setTimeout(() => this.scrollToSlide(this.activeSlideIndex), 50);
  }

  private scrollToSlide(index: number): void {
    if (!this.stageInnerEl) return;
    const card = this.stageInnerEl.querySelector(
      `.office-pptx-scroll-card[data-slide-index="${index}"]`
    );
    if (card) {
      card.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }

  // -------------------------------------------------------------
  // SLIDE SORTER GRID MODE
  // -------------------------------------------------------------

  private renderSlideSorterGrid(): void {
    if (!this.stageInnerEl || !this.presentation) return;

    const grid = this.stageInnerEl.createDiv({ cls: "office-pptx-sorter-grid" });

    for (let i = 0; i < this.presentation.slides.length; i++) {
      const slide = this.presentation.slides[i];
      const card = grid.createDiv({
        cls: `office-pptx-grid-card ${i === this.activeSlideIndex ? "is-active" : ""}`,
      });
      card.dataset.slideIndex = String(i);

      // Slide number badge
      card.createDiv({
        cls: "office-pptx-grid-badge",
        text: String(slide.slideNumber),
      });

      // Preview box
      const previewBox = card.createDiv({ cls: "office-pptx-grid-preview" });
      const slidePreview = PptxRenderer.renderSlide(slide, this.presentation, {
        showSlideNumberBadge: false,
      });

      const presW = this.presentation.width || 960;
      const presH = this.presentation.height || 540;
      const targetW = 280;
      const scale = targetW / presW;
      const targetH = Math.round(presH * scale);

      previewBox.style.width = `${targetW}px`;
      previewBox.style.height = `${targetH}px`;
      previewBox.style.position = "relative";
      previewBox.style.overflow = "hidden";

      slidePreview.style.transform = `scale(${scale})`;
      slidePreview.style.transformOrigin = "top left";
      slidePreview.style.pointerEvents = "none";
      previewBox.appendChild(slidePreview);

      // Title footer
      const titleFooter = card.createDiv({ cls: "office-pptx-grid-title" });
      titleFooter.textContent = slide.title || `Slide ${slide.slideNumber}`;

      card.onclick = () => {
        this.delegate?.onSelectSlide(i);
      };
    }
  }

  private highlightGridCard(index: number): void {
    if (!this.stageInnerEl) return;
    const cards = this.stageInnerEl.querySelectorAll(".office-pptx-grid-card");
    cards.forEach((c, idx) => {
      c.classList.toggle("is-active", idx === index);
    });
  }
}
