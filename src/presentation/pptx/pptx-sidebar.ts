import { setIcon } from "obsidian";
import { PptxPresentation, PptxSlide, PptxRenderer } from "../../domain/pptx";

export interface PptxSidebarDelegate {
  onSlideSelected(slideIndex: number): void;
  onAddSlide?(): void;
  onDuplicateSlide?(slideIndex: number): void;
  onDeleteSlide?(slideIndex: number): void;
  onMoveSlideUp?(slideIndex: number): void;
  onMoveSlideDown?(slideIndex: number): void;
}

export class PptxSidebar {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxSidebarDelegate | null = null;
  private presentation: PptxPresentation | null = null;
  private activeSlideIndex = 0;
  private searchQuery = "";
  private slideCardEls: HTMLElement[] = [];

  render(
    container: HTMLElement,
    presentation: PptxPresentation,
    activeSlideIndex: number,
    delegate: PptxSidebarDelegate
  ): void {
    this.containerEl = container;
    this.presentation = presentation;
    this.activeSlideIndex = activeSlideIndex;
    this.delegate = delegate;
    this.slideCardEls = [];

    this.containerEl.empty();
    this.containerEl.addClass("office-pptx-sidebar");

    // Sidebar header
    const headerEl = this.containerEl.createDiv({ cls: "office-pptx-sidebar-header" });
    const headerLeft = headerEl.createDiv({ cls: "office-pptx-sidebar-header-left" });
    headerLeft.createSpan({ cls: "office-pptx-sidebar-title", text: "Slides" });
    headerLeft.createSpan({
      cls: "office-pptx-sidebar-count",
      text: `${presentation.slides.length} slides`,
    });

    const addBtn = headerEl.createEl("button", {
      cls: "office-pptx-sidebar-add-btn office-pptx-thumb-btn",
      attr: { title: "Add New Slide" },
    });
    setIcon(addBtn, "plus");
    addBtn.onclick = (e) => {
      e.stopPropagation();
      this.delegate?.onAddSlide?.();
    };

    // Thumbnail List Container
    const listEl = this.containerEl.createDiv({ cls: "office-pptx-thumbnail-list" });

    for (let i = 0; i < presentation.slides.length; i++) {
      const slide = presentation.slides[i];
      const cardEl = listEl.createDiv({
        cls: `office-pptx-thumbnail-card ${i === this.activeSlideIndex ? "is-active" : ""}`,
      });
      cardEl.dataset.slideIndex = String(i);

      // Slide number badge
      cardEl.createDiv({
        cls: "office-pptx-thumbnail-num",
        text: String(slide.slideNumber),
      });

      // Quick hover action buttons for slide
      const actionsEl = cardEl.createDiv({ cls: "office-pptx-thumbnail-actions" });

      const dupBtn = actionsEl.createEl("button", {
        cls: "office-pptx-thumb-btn",
        attr: { title: "Duplicate Slide" },
      });
      setIcon(dupBtn, "copy");
      dupBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onDuplicateSlide?.(i);
      };

      const moveUpBtn = actionsEl.createEl("button", {
        cls: "office-pptx-thumb-btn",
        attr: { title: "Move Slide Up" },
      });
      setIcon(moveUpBtn, "arrow-up");
      if (i === 0) moveUpBtn.disabled = true;
      moveUpBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onMoveSlideUp?.(i);
      };

      const moveDownBtn = actionsEl.createEl("button", {
        cls: "office-pptx-thumb-btn",
        attr: { title: "Move Slide Down" },
      });
      setIcon(moveDownBtn, "arrow-down");
      if (i === presentation.slides.length - 1) moveDownBtn.disabled = true;
      moveDownBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onMoveSlideDown?.(i);
      };

      const delBtn = actionsEl.createEl("button", {
        cls: "office-pptx-thumb-btn mod-danger",
        attr: { title: "Delete Slide" },
      });
      setIcon(delBtn, "trash-2");
      if (presentation.slides.length <= 1) delBtn.disabled = true;
      delBtn.onclick = (e) => {
        e.stopPropagation();
        this.delegate?.onDeleteSlide?.(i);
      };

      // Preview box
      const previewBox = cardEl.createDiv({ cls: "office-pptx-thumbnail-preview" });
      const slidePreview = PptxRenderer.renderSlide(slide, presentation, {
        showSlideNumberBadge: false,
      });

      // Scale thumbnail preview to fit preview box (160px width)
      const presW = presentation.width || 960;
      const presH = presentation.height || 540;
      const targetW = 160;
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

      // Title / summary
      const infoEl = cardEl.createDiv({ cls: "office-pptx-thumbnail-info" });
      const title = slide.title || `Slide ${slide.slideNumber}`;
      infoEl.createSpan({ cls: "office-pptx-thumbnail-title", text: title });

      // Click handler
      cardEl.onclick = () => {
        this.delegate?.onSlideSelected(i);
      };

      this.slideCardEls.push(cardEl);
    }
  }

  setActiveSlide(index: number): void {
    this.activeSlideIndex = index;
    for (let i = 0; i < this.slideCardEls.length; i++) {
      const card = this.slideCardEls[i];
      const isActive = i === index;
      card.classList.toggle("is-active", isActive);
      if (isActive) {
        card.scrollIntoView({ block: "nearest", behavior: "smooth" });
      }
    }
  }

  filterSlides(query: string): void {
    this.searchQuery = query.toLowerCase().trim();
    if (!this.presentation) return;

    for (let i = 0; i < this.presentation.slides.length; i++) {
      const slide = this.presentation.slides[i];
      const card = this.slideCardEls[i];
      if (!card) continue;

      if (!this.searchQuery) {
        card.style.display = "";
        card.classList.remove("is-match");
        continue;
      }

      // Check title, elements text, and notes
      let match = (slide.title || "").toLowerCase().includes(this.searchQuery);
      if (!match) {
        for (const el of slide.elements) {
          if (el.paragraphs) {
            for (const p of el.paragraphs) {
              for (const r of p.runs) {
                if (r.text.toLowerCase().includes(this.searchQuery)) {
                  match = true;
                  break;
                }
              }
              if (match) break;
            }
          }
          if (match) break;
        }
      }
      if (!match && slide.notes) {
        match = slide.notes.toLowerCase().includes(this.searchQuery);
      }

      card.style.display = match ? "" : "none";
      card.classList.toggle("is-match", match);
    }
  }
}
