import { setIcon } from "obsidian";
import {
  PptxPresentation,
  PptxSlide,
  PptxRenderer,
  TransitionType,
} from "../../domain/pptx";

export interface PptxSlideshowDelegate {
  onExit(): void;
  onSlideChanged(index: number): void;
}

export class PptxSlideshow {
  private overlayEl: HTMLElement | null = null;
  private stageEl: HTMLElement | null = null;
  private hudEl: HTMLElement | null = null;
  private timerEl: HTMLElement | null = null;
  private laserDotEl: HTMLElement | null = null;
  private blankScreenEl: HTMLElement | null = null;

  private presentation: PptxPresentation | null = null;
  private activeSlideIndex = 0;
  private delegate: PptxSlideshowDelegate | null = null;
  private transitionOverride: TransitionType | "auto" = "auto";

  private timerInterval: any = null;
  private startTime = 0;
  private hudTimeout: any = null;
  private isLaserActive = false;
  private isBlackScreen = false;
  private isWhiteScreen = false;
  private isTransitioning = false;

  private keyListener: ((e: KeyboardEvent) => void) | null = null;
  private mouseMoveListener: ((e: MouseEvent) => void) | null = null;

  start(
    presentation: PptxPresentation,
    initialIndex: number,
    transitionOverride: TransitionType | "auto",
    delegate: PptxSlideshowDelegate
  ): void {
    this.presentation = presentation;
    this.activeSlideIndex = initialIndex;
    this.transitionOverride = transitionOverride;
    this.delegate = delegate;

    // Create fullscreen overlay
    this.overlayEl = document.body.createDiv({ cls: "office-pptx-slideshow-overlay" });

    // Stage container
    this.stageEl = this.overlayEl.createDiv({ cls: "office-pptx-slideshow-stage" });

    // Blank screen layer
    this.blankScreenEl = this.overlayEl.createDiv({ cls: "office-pptx-blank-layer" });
    this.blankScreenEl.style.display = "none";

    // Laser pointer dot
    this.laserDotEl = this.overlayEl.createDiv({ cls: "office-pptx-laser-dot" });
    this.laserDotEl.style.display = "none";

    // Presenter HUD
    this.buildHud();

    // Start presentation timer
    this.startTime = Date.now();
    this.timerInterval = setInterval(() => this.updateTimer(), 1000);

    // Initial slide render
    this.renderCurrentSlide();

    // Attach listeners
    this.setupListeners();

    // Try requesting fullscreen
    try {
      if (this.overlayEl.requestFullscreen) {
        void this.overlayEl.requestFullscreen();
      }
    } catch (e) {
      // Fullscreen API may be blocked or restricted, overlay still functions perfectly
    }
  }

  stop(): void {
    if (this.timerInterval) {
      clearInterval(this.timerInterval);
      this.timerInterval = null;
    }
    if (this.hudTimeout) {
      clearTimeout(this.hudTimeout);
      this.hudTimeout = null;
    }
    if (this.keyListener) {
      window.removeEventListener("keydown", this.keyListener);
      this.keyListener = null;
    }
    if (this.mouseMoveListener) {
      window.removeEventListener("mousemove", this.mouseMoveListener);
      this.mouseMoveListener = null;
    }

    if (document.fullscreenElement) {
      try {
        void document.exitFullscreen();
      } catch (e) {}
    }

    if (this.overlayEl) {
      this.overlayEl.remove();
      this.overlayEl = null;
    }

    this.delegate?.onExit();
  }

  private buildHud(): void {
    if (!this.overlayEl) return;

    this.hudEl = this.overlayEl.createDiv({ cls: "office-pptx-slideshow-hud" });

    // Prev Button
    const prevBtn = this.hudEl.createEl("button", {
      cls: "office-pptx-hud-btn",
      attr: { title: "Previous Slide" },
    });
    setIcon(prevBtn, "chevron-left");
    prevBtn.onclick = () => this.prevSlide();

    // Slide Counter
    const counter = this.hudEl.createSpan({ cls: "office-pptx-hud-counter" });
    counter.textContent = `${this.activeSlideIndex + 1} / ${this.presentation?.slides.length || 1}`;

    // Next Button
    const nextBtn = this.hudEl.createEl("button", {
      cls: "office-pptx-hud-btn",
      attr: { title: "Next Slide" },
    });
    setIcon(nextBtn, "chevron-right");
    nextBtn.onclick = () => this.nextSlide();

    this.hudEl.createDiv({ cls: "office-pptx-hud-sep" });

    // Timer
    this.timerEl = this.hudEl.createSpan({ cls: "office-pptx-hud-timer", text: "00:00" });

    this.hudEl.createDiv({ cls: "office-pptx-hud-sep" });

    // Laser pointer toggle button
    const laserBtn = this.hudEl.createEl("button", {
      cls: "office-pptx-hud-btn",
      attr: { title: "Toggle Laser Pointer (L)" },
    });
    setIcon(laserBtn, "target");
    laserBtn.onclick = () => this.toggleLaser();

    // Black screen toggle
    const blackBtn = this.hudEl.createEl("button", {
      cls: "office-pptx-hud-btn",
      attr: { title: "Blank Screen (B)" },
    });
    setIcon(blackBtn, "eye-off");
    blackBtn.onclick = () => this.toggleBlackScreen();

    // Exit Button
    const exitBtn = this.hudEl.createEl("button", {
      cls: "office-pptx-hud-btn mod-exit",
      attr: { title: "Exit Presentation (Esc)" },
    });
    setIcon(exitBtn, "x");
    exitBtn.onclick = () => this.stop();
  }

  private setupListeners(): void {
    // Keyboard navigation
    this.keyListener = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        this.stop();
      } else if (
        e.key === "ArrowRight" ||
        e.key === "ArrowDown" ||
        e.key === "PageDown" ||
        e.key === " " ||
        e.key === "Enter"
      ) {
        e.preventDefault();
        this.nextSlide();
      } else if (
        e.key === "ArrowLeft" ||
        e.key === "ArrowUp" ||
        e.key === "PageUp" ||
        e.key === "Backspace"
      ) {
        e.preventDefault();
        this.prevSlide();
      } else if (e.key === "Home") {
        e.preventDefault();
        this.goToSlide(0);
      } else if (e.key === "End") {
        e.preventDefault();
        this.goToSlide((this.presentation?.slides.length || 1) - 1);
      } else if (e.key === "b" || e.key === "B") {
        e.preventDefault();
        this.toggleBlackScreen();
      } else if (e.key === "w" || e.key === "W") {
        e.preventDefault();
        this.toggleWhiteScreen();
      } else if (e.key === "l" || e.key === "L") {
        e.preventDefault();
        this.toggleLaser();
      }
    };
    window.addEventListener("keydown", this.keyListener);

    // Mouse movement: Laser tracking & HUD auto-hide
    this.mouseMoveListener = (e: MouseEvent) => {
      if (this.isLaserActive && this.laserDotEl) {
        this.laserDotEl.style.left = `${e.clientX}px`;
        this.laserDotEl.style.top = `${e.clientY}px`;
      }
      this.showHudBriefly();
    };
    window.addEventListener("mousemove", this.mouseMoveListener);

    // Window resize
    window.addEventListener("resize", () => this.scaleSlideToWindow());
  }

  private showHudBriefly(): void {
    if (!this.hudEl) return;
    this.hudEl.classList.add("is-visible");

    if (this.hudTimeout) {
      clearTimeout(this.hudTimeout);
    }
    this.hudTimeout = setTimeout(() => {
      this.hudEl?.classList.remove("is-visible");
    }, 3000);
  }

  private updateTimer(): void {
    if (!this.timerEl) return;
    const elapsedSec = Math.floor((Date.now() - this.startTime) / 1000);
    const mins = Math.floor(elapsedSec / 60);
    const secs = elapsedSec % 60;
    this.timerEl.textContent = `${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
  }

  private nextSlide(): void {
    if (!this.presentation) return;
    if (this.activeSlideIndex < this.presentation.slides.length - 1) {
      this.goToSlide(this.activeSlideIndex + 1);
    }
  }

  private prevSlide(): void {
    if (this.activeSlideIndex > 0) {
      this.goToSlide(this.activeSlideIndex - 1);
    }
  }

  private goToSlide(newIndex: number): void {
    if (!this.presentation) return;
    if (newIndex < 0 || newIndex >= this.presentation.slides.length) return;
    if (this.isTransitioning) return;

    const oldIndex = this.activeSlideIndex;
    this.activeSlideIndex = newIndex;

    this.executeSlideTransition(oldIndex, newIndex);
    this.updateHudCounter();
    this.delegate?.onSlideChanged(newIndex);
  }

  private renderCurrentSlide(): void {
    if (!this.stageEl || !this.presentation) return;
    this.stageEl.empty();

    const slide = this.presentation.slides[this.activeSlideIndex];
    if (!slide) return;

    const slideNode = PptxRenderer.renderSlide(slide, this.presentation, {
      showSlideNumberBadge: true,
      totalSlides: this.presentation.slides.length,
    });
    slideNode.addClass("office-pptx-slideshow-active-slide");
    this.stageEl.appendChild(slideNode);

    this.scaleSlideToWindow();
  }

  private executeSlideTransition(fromIndex: number, toIndex: number): void {
    if (!this.stageEl || !this.presentation) return;
    const toSlide = this.presentation.slides[toIndex];
    if (!toSlide) return;

    this.isTransitioning = true;

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
      this.renderCurrentSlide();
      this.isTransitioning = false;
      return;
    }

    if (transType === "push" && !toSlide.transition?.direction) {
      transDir = toIndex >= fromIndex ? "l" : "r";
    }

    const oldSlide = this.stageEl.querySelector(".office-pptx-slideshow-active-slide") as HTMLElement;
    const newSlide = PptxRenderer.renderSlide(toSlide, this.presentation, {
      showSlideNumberBadge: true,
      totalSlides: this.presentation.slides.length,
    });
    newSlide.addClass("office-pptx-slideshow-active-slide");

    const animOutClass = `pptx-trans-${transType}-out${transType === "push" ? `-${transDir}` : ""}`;
    const animInClass = `pptx-trans-${transType}-in${transType === "push" ? `-${transDir}` : ""}`;

    if (oldSlide) {
      oldSlide.style.animationDuration = `${duration}ms`;
      oldSlide.classList.add(animOutClass);
    }

    newSlide.style.animationDuration = `${duration}ms`;
    newSlide.classList.add(animInClass);
    this.stageEl.appendChild(newSlide);

    this.scaleSlideToWindow();

    setTimeout(() => {
      if (oldSlide && oldSlide.parentElement === this.stageEl) {
        oldSlide.remove();
      }
      newSlide.classList.remove(animInClass);
      newSlide.style.animationDuration = "";
      this.isTransitioning = false;
    }, duration + 50);
  }

  private scaleSlideToWindow(): void {
    if (!this.stageEl || !this.presentation) return;

    const presW = this.presentation.width || 960;
    const presH = this.presentation.height || 540;

    const winW = window.innerWidth;
    const winH = window.innerHeight;

    const scale = Math.min(winW / presW, winH / presH);

    const slides = this.stageEl.querySelectorAll(".office-pptx-slide") as NodeListOf<HTMLElement>;
    slides.forEach((s) => {
      s.style.transform = `scale(${scale})`;
      s.style.transformOrigin = "center center";
    });
  }

  private updateHudCounter(): void {
    if (!this.hudEl || !this.presentation) return;
    const counter = this.hudEl.querySelector(".office-pptx-hud-counter");
    if (counter) {
      counter.textContent = `${this.activeSlideIndex + 1} / ${this.presentation.slides.length}`;
    }
  }

  private toggleLaser(): void {
    this.isLaserActive = !this.isLaserActive;
    if (this.laserDotEl) {
      this.laserDotEl.style.display = this.isLaserActive ? "block" : "none";
    }
    if (this.overlayEl) {
      this.overlayEl.style.cursor = this.isLaserActive ? "none" : "default";
    }
  }

  private toggleBlackScreen(): void {
    this.isBlackScreen = !this.isBlackScreen;
    this.isWhiteScreen = false;
    if (this.blankScreenEl) {
      this.blankScreenEl.style.backgroundColor = "#000000";
      this.blankScreenEl.style.display = this.isBlackScreen ? "block" : "none";
    }
  }

  private toggleWhiteScreen(): void {
    this.isWhiteScreen = !this.isWhiteScreen;
    this.isBlackScreen = false;
    if (this.blankScreenEl) {
      this.blankScreenEl.style.backgroundColor = "#ffffff";
      this.blankScreenEl.style.display = this.isWhiteScreen ? "block" : "none";
    }
  }
}
