/**
 * Floating Popover Card for Viewing and Editing Sticky Notes and Highlight Comments.
 * Includes bi-directional linking with Obsidian Vault Markdown notes, author, and color.
 * Supports dynamic repositioning on scroll, text selection, and keyboard shortcuts.
 */

import { App, TFile, Notice } from "obsidian";
import {
  PdfAnnotation,
  PDF_COLOR_PRESETS,
  rgbToHex,
  hexToRgb,
  PdfColorSchemeProfile,
  findTagForColor,
  normalizeHexColor,
} from "../../domain/pdf";

export interface PdfNotePopoverDelegate {
  onUpdateAnnotation(annotation: PdfAnnotation): void;
  onDeleteAnnotation(id: string): void;
  onOpenLinkedNote?(vaultPath: string): void;
}

export class PdfNotePopover {
  private el: HTMLElement | null = null;
  private currentAnnotation: PdfAnnotation | null = null;
  private currentTargetEl: HTMLElement | SVGElement | null = null;
  private scrollContainerEl: HTMLElement | null = null;
  private delegate: PdfNotePopoverDelegate | null = null;
  private colorScheme?: PdfColorSchemeProfile;
  private app: App;

  constructor(app: App) {
    this.app = app;
  }

  setColorScheme(scheme: PdfColorSchemeProfile): void {
    this.colorScheme = scheme;
  }

  setScrollContainer(container: HTMLElement | null): void {
    this.scrollContainerEl = container;
  }

  show(
    targetEl: HTMLElement | SVGElement,
    annotation: PdfAnnotation,
    delegate: PdfNotePopoverDelegate,
    colorScheme?: PdfColorSchemeProfile
  ): void {
    this.hide();
    if (colorScheme) {
      this.colorScheme = colorScheme;
    }
    this.currentAnnotation = annotation;
    this.currentTargetEl = targetEl;
    this.delegate = delegate;

    const popover = document.createElement("div");
    popover.className = "pdf-note-popover";
    this.el = popover;

    // Header: Type, Author, Date, Close
    const header = popover.createDiv({ cls: "pdf-popover-header" });
    const titleWrap = header.createDiv({ cls: "pdf-popover-title-wrap" });

    const colorDot = titleWrap.createSpan({ cls: "pdf-popover-color-dot" });
    const initialHex = rgbToHex(annotation.color);
    colorDot.style.backgroundColor = initialHex;

    const typeTitle =
      annotation.type === "highlight"
        ? "Highlight Note"
        : annotation.type === "underline"
        ? "Underline Note"
        : annotation.type === "text"
        ? "Sticky Note"
        : "Annotation";
    titleWrap.createSpan({ cls: "pdf-popover-type", text: typeTitle });

    // Semantic Tag Badge
    const tagBadge = titleWrap.createSpan({ cls: "pdf-popover-tag-badge" });
    const initialTag = findTagForColor(this.colorScheme, annotation.color);
    if (initialTag?.tag) {
      tagBadge.setText(initialTag.tag);
      tagBadge.style.color = initialTag.hex;
      tagBadge.style.borderColor = initialTag.hex;
    } else {
      tagBadge.addClass("is-hidden");
    }

    const authorText = annotation.author || "User";
    const dateText = annotation.modifiedAt
      ? new Date(annotation.modifiedAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
      : "Just now";
    header.createSpan({ cls: "pdf-popover-meta", text: `${authorText} · ${dateText}` });

    const closeBtn = header.createEl("button", { cls: "pdf-popover-close-btn", text: "✕", title: "Close (Esc)" });
    closeBtn.onclick = (e) => {
      e.stopPropagation();
      this.hide();
    };

    // Highlight / Underline text quote if attached to a selection
    const quotedSnippet = annotation.highlightedText || (annotation as any).highlightedText;
    if (quotedSnippet) {
      const quoteWrap = popover.createDiv({ cls: "pdf-popover-quote-wrap" });
      const quoteEl = quoteWrap.createDiv({ cls: "pdf-popover-quote" });
      quoteEl.setText(`“${quotedSnippet}”`);

      const copyBtn = quoteWrap.createEl("button", {
        cls: "pdf-popover-copy-btn",
        text: "Copy",
        attr: { title: "Copy quoted text" },
      });
      copyBtn.onclick = async (e) => {
        e.stopPropagation();
        await navigator.clipboard.writeText(quotedSnippet);
        new Notice("Copied text to clipboard");
      };
    }

    // Body: Textarea for comments
    const body = popover.createDiv({ cls: "pdf-popover-body" });
    const textarea = body.createEl("textarea", {
      cls: "pdf-popover-textarea",
      attr: { placeholder: "Add a note or comment...", rows: "3" },
    });
    const existingUserComment =
      annotation.contents && annotation.contents !== annotation.highlightedText
        ? annotation.contents
        : "";
    textarea.value = existingUserComment;
    textarea.oninput = () => {
      if (this.currentAnnotation && this.delegate) {
        const val = textarea.value.trim();
        this.currentAnnotation.contents = val ? textarea.value : (this.currentAnnotation.highlightedText || "");
        this.delegate.onUpdateAnnotation(this.currentAnnotation);
      }
    };

    // Color Swatches from active Scheme
    const colorRow = body.createDiv({ cls: "pdf-popover-colors" });
    const availableColors =
      this.colorScheme?.colors && this.colorScheme.colors.length > 0
        ? this.colorScheme.colors
        : PDF_COLOR_PRESETS.map((p) => ({
            id: p.name.toLowerCase(),
            hex: p.hex,
            name: p.name,
            tag: "",
            rgb: p.rgb,
          }));

    for (const item of availableColors) {
      const swatch = colorRow.createDiv({ cls: "pdf-popover-swatch" });
      swatch.style.backgroundColor = item.hex;
      const tagText = item.tag ? ` (${item.tag})` : "";
      swatch.title = `${item.name}${tagText}`;

      if (rgbToHex(annotation.color).toLowerCase() === normalizeHexColor(item.hex)) {
        swatch.addClass("is-selected");
      }

      swatch.onclick = (e) => {
        e.stopPropagation();
        colorRow.querySelectorAll(".pdf-popover-swatch").forEach((s) => s.removeClass("is-selected"));
        swatch.addClass("is-selected");
        colorDot.style.backgroundColor = item.hex;

        // Update tag badge
        if (item.tag) {
          tagBadge.setText(item.tag);
          tagBadge.style.color = item.hex;
          tagBadge.style.borderColor = item.hex;
          tagBadge.removeClass("is-hidden");
        } else {
          tagBadge.addClass("is-hidden");
        }

        if (this.currentAnnotation && this.delegate) {
          this.currentAnnotation.color = item.rgb || hexToRgb(item.hex);
          this.delegate.onUpdateAnnotation(this.currentAnnotation);
        }
      };
    }

    // Linked Obsidian Note Section
    const linkSection = body.createDiv({ cls: "pdf-popover-link-section" });
    const linkHeader = linkSection.createDiv({ cls: "pdf-popover-link-header" });
    linkHeader.createSpan({ text: "Linked Obsidian Note:" });

    const linkRow = linkSection.createDiv({ cls: "pdf-popover-link-row" });
    const linkInput = linkRow.createEl("input", {
      cls: "pdf-popover-link-input",
      type: "text",
      value: annotation.linkedNote || "",
      attr: { placeholder: "e.g. Notes/Summary.md" },
    });

    linkInput.onchange = () => {
      const val = linkInput.value.trim();
      if (this.currentAnnotation && this.delegate) {
        this.currentAnnotation.linkedNote = val || undefined;
        this.delegate.onUpdateAnnotation(this.currentAnnotation);
      }
    };

    if (annotation.linkedNote) {
      const openNoteBtn = linkRow.createEl("button", {
        cls: "pdf-popover-btn mod-primary",
        text: "Open Note",
      });
      openNoteBtn.onclick = (e) => {
        e.stopPropagation();
        this.openLinkedNote(annotation.linkedNote!);
      };
    } else {
      const createNoteBtn = linkRow.createEl("button", {
        cls: "pdf-popover-btn",
        text: "Create Note",
      });
      createNoteBtn.onclick = async (e) => {
        e.stopPropagation();
        await this.createNoteFromAnnotation(annotation, linkInput);
      };
    }

    // Footer: Delete button
    const footer = popover.createDiv({ cls: "pdf-popover-footer" });
    const deleteBtn = footer.createEl("button", {
      cls: "pdf-popover-delete-btn",
      text: "Delete Annotation",
    });
    deleteBtn.onclick = (e) => {
      e.stopPropagation();
      if (this.currentAnnotation && this.delegate) {
        this.delegate.onDeleteAnnotation(this.currentAnnotation.id);
      }
      this.hide();
    };

    // Position popover relative to target element
    document.body.appendChild(popover);
    this.updatePosition();

    // Auto-focus textarea with cursor at end
    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(textarea.value.length, textarea.value.length);
    }, 60);

    // Event listeners
    setTimeout(() => {
      window.addEventListener("pointerdown", this.onOutsidePointerDown);
      window.addEventListener("keydown", this.onKeyDown);
      if (this.scrollContainerEl) {
        this.scrollContainerEl.addEventListener("scroll", this.onScroll, { passive: true });
      }
    }, 80);
  }

  private onScroll = (): void => {
    if (this.el && this.currentTargetEl) {
      this.updatePosition();
    }
  };

  private onKeyDown = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      this.hide();
    }
  };

  public updatePosition(): void {
    if (!this.el || !this.currentTargetEl) return;
    const targetRect = this.currentTargetEl.getBoundingClientRect();
    const popoverRect = this.el.getBoundingClientRect();

    let left = targetRect.left + (targetRect.width > 0 ? targetRect.width + 12 : 28);
    let top = targetRect.top - 12;

    // Boundary checks against window viewport
    const margin = 16;
    if (left + popoverRect.width > window.innerWidth - margin) {
      // Place to the left of target if not enough room on right
      left = targetRect.left - popoverRect.width - 12;
      if (left < margin) {
        left = window.innerWidth - popoverRect.width - margin;
      }
    }
    if (left < margin) left = margin;

    if (top + popoverRect.height > window.innerHeight - margin) {
      top = window.innerHeight - popoverRect.height - margin;
    }
    if (top < margin) top = margin;

    this.el.style.left = `${Math.round(left)}px`;
    this.el.style.top = `${Math.round(top)}px`;
  }

  private onOutsidePointerDown = (e: PointerEvent): void => {
    if (this.el && !this.el.contains(e.target as Node)) {
      // Don't close if clicking target element itself
      if (this.currentTargetEl && this.currentTargetEl.contains(e.target as Node)) {
        return;
      }
      this.hide();
    }
  };

  private async openLinkedNote(path: string): Promise<void> {
    const file = this.app.vault.getAbstractFileByPath(path);
    if (file instanceof TFile) {
      await this.app.workspace.openLinkText(file.path, "", "split");
    } else {
      new Notice(`Note not found: ${path}`);
    }
  }

  private async createNoteFromAnnotation(
    annotation: PdfAnnotation,
    inputEl: HTMLInputElement
  ): Promise<void> {
    const pageNum = annotation.pageIndex + 1;
    const sanitizedTitle = `Note - p${pageNum} ${annotation.id.slice(0, 6)}`;
    const folder = "Notes";
    const notePath = `${folder}/${sanitizedTitle}.md`;

    try {
      // Ensure folder exists
      const folderExists = this.app.vault.getAbstractFileByPath(folder);
      if (!folderExists) {
        await this.app.vault.createFolder(folder);
      }

      const quote = annotation.highlightedText || (annotation as any).highlightedText || annotation.contents || "";
      const content = [
        `---`,
        `page: ${pageNum}`,
        `annotation_id: "${annotation.id}"`,
        `date: ${new Date().toISOString()}`,
        `tags:`,
        `  - pdf-note`,
        `---`,
        ``,
        `# Notes on Page ${pageNum}`,
        ``,
        quote ? `> "${quote}"\n\n` : "",
        `## User Notes`,
        annotation.contents ? `${annotation.contents}\n` : "",
        ``,
      ].join("\n");

      let created = this.app.vault.getAbstractFileByPath(notePath);
      if (!created) {
        created = await this.app.vault.create(notePath, content);
      }

      annotation.linkedNote = notePath;
      inputEl.value = notePath;
      if (this.delegate) {
        this.delegate.onUpdateAnnotation(annotation);
      }
      new Notice(`Created linked note: ${notePath}`);
      await this.app.workspace.openLinkText(notePath, "", "split");
    } catch (err) {
      console.error("Failed to create linked note:", err);
      new Notice(`Could not create note: ${(err as Error).message}`);
    }
  }

  hide(): void {
    window.removeEventListener("pointerdown", this.onOutsidePointerDown);
    window.removeEventListener("keydown", this.onKeyDown);
    if (this.scrollContainerEl) {
      this.scrollContainerEl.removeEventListener("scroll", this.onScroll);
    }
    if (this.el) {
      this.el.remove();
      this.el = null;
    }
    this.currentAnnotation = null;
    this.currentTargetEl = null;
    this.delegate = null;
  }
}
