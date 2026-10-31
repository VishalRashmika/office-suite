import { setIcon } from "obsidian";
import { PptxSlide } from "../../domain/pptx";

export interface PptxNotesDelegate {
  onClose(): void;
}

export class PptxNotes {
  private containerEl: HTMLElement | null = null;
  private delegate: PptxNotesDelegate | null = null;
  private contentEl: HTMLElement | null = null;
  private titleEl: HTMLElement | null = null;

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
    this.titleEl = titleWrap.createSpan({ cls: "office-pptx-notes-title", text: "Speaker Notes" });

    const closeBtn = headerEl.createEl("button", {
      cls: "office-pptx-toolbar-btn office-pptx-notes-close",
      attr: { title: "Close Notes (S)" },
    });
    setIcon(closeBtn, "x");
    closeBtn.onclick = () => this.delegate?.onClose();

    // Notes content body
    this.contentEl = this.containerEl.createDiv({ cls: "office-pptx-notes-content" });
  }

  updateSlide(slide: PptxSlide | undefined): void {
    if (!this.titleEl || !this.contentEl) return;

    if (!slide) {
      this.titleEl.textContent = "Speaker Notes";
      this.contentEl.innerHTML = `<div class="office-pptx-notes-empty">No slide selected</div>`;
      return;
    }

    this.titleEl.textContent = `Speaker Notes — Slide ${slide.slideNumber}`;

    if (slide.notes && slide.notes.trim()) {
      this.contentEl.empty();
      for (const line of slide.notes.trim().split("\n")) {
        const p = this.contentEl.createEl("p", { text: line });
        p.className = "office-pptx-notes-line";
      }
    } else {
      this.contentEl.innerHTML = `<div class="office-pptx-notes-empty">No speaker notes recorded for this slide.</div>`;
    }
  }
}
