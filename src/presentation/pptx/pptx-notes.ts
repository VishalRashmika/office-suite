import { setIcon } from "obsidian";
import { PptxSlide } from "../../domain/pptx";

export interface PptxNotesDelegate {
  onClose(): void;
  onNotesChange?(slideIndex: number, newNotes: string): void;
}

export class PptxNotes {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxNotesDelegate | null = null;
  private contentEl: HTMLElement | null = null;
  private titleEl: HTMLElement | null = null;
  private textareaEl: HTMLTextAreaElement | null = null;
  private currentSlideIndex = 0;

  render(container: HTMLElement, delegate: PptxNotesDelegate): void {
    this.containerEl = container;
    this.delegate = delegate;

    this.containerEl.empty();
    this.containerEl.addClass("office-pptx-notes-drawer");

    // Drawer header
    const headerEl = this.containerEl.createDiv({ cls: "office-pptx-notes-header" });
    const titleWrap = headerEl.createDiv({ cls: "office-pptx-notes-title-wrap" });
    const iconSpan = titleWrap.createSpan({ cls: "office-pptx-notes-icon" });
    setIcon(iconSpan, "file-text");
    this.titleEl = titleWrap.createSpan({
      cls: "office-pptx-notes-title",
      text: "Speaker Notes",
    });

    const closeBtn = headerEl.createEl("button", {
      cls: "office-pptx-toolbar-btn office-pptx-notes-close",
      attr: { title: "Close Notes (S)" },
    });
    setIcon(closeBtn, "x");
    closeBtn.onclick = () => this.delegate?.onClose();

    // Notes content body with editable textarea
    this.contentEl = this.containerEl.createDiv({ cls: "office-pptx-notes-content" });
    this.textareaEl = this.contentEl.createEl("textarea", {
      cls: "office-pptx-notes-editor",
      attr: { placeholder: "Add speaker notes for this slide here..." },
    });

    this.textareaEl.oninput = () => {
      this.delegate?.onNotesChange?.(
        this.currentSlideIndex,
        this.textareaEl?.value || ""
      );
    };
  }

  updateSlide(slide: PptxSlide | undefined): void {
    if (!this.titleEl || !this.textareaEl) return;

    if (!slide) {
      this.titleEl.textContent = "Speaker Notes";
      this.textareaEl.value = "";
      this.textareaEl.disabled = true;
      return;
    }

    this.currentSlideIndex = slide.index;
    this.titleEl.textContent = `Speaker Notes — Slide ${slide.slideNumber}`;
    this.textareaEl.disabled = false;
    this.textareaEl.value = slide.notes || "";
  }
}
