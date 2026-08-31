import { App } from "obsidian";
import { Node as PMNode } from "prosemirror-model";
import { EditorView, NodeView } from "prosemirror-view";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { TextPromptModal } from "../common/text-prompt-modal";

export class DocxImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private wrapper: HTMLElement;
  private toolbar: HTMLElement | null = null;
  private tooltip: HTMLElement | null = null;
  private handles: HTMLElement[] = [];
  private isResizing = false;
  private isSelected = false;

  constructor(
    private node: PMNode,
    private view: EditorView,
    private getPos: (() => number | undefined) | boolean,
    private docxDoc: DocxDocument | null,
    private app?: App
  ) {
    this.dom = createSpan({
      cls: "docx-image-wrapper-outer",
      attr: { contenteditable: "false" },
    });

    const align = (this.node.attrs.align as string) || "center";
    this.wrapper = this.dom.createSpan({
      cls: `docx-image-container docx-image-align-${align}`,
    });

    const src = (this.node.attrs.src as string) || "";
    const alt = (this.node.attrs.alt as string) || "";
    this.img = this.wrapper.createEl("img", {
      cls: "docx-image-element",
      attr: { src, alt },
    });

    const width = this.node.attrs.width as number | null;
    const height = this.node.attrs.height as number | null;
    if (width) {
      this.img.setCssStyles({ width: `${width}px` });
    }
    if (height) {
      this.img.setCssStyles({ height: `${height}px` });
    }

    // Natural dimension capture if not yet stored
    this.img.onload = () => {
      if (!this.node.attrs.naturalWidth || !this.node.attrs.naturalHeight) {
        const nw = this.img.naturalWidth;
        const nh = this.img.naturalHeight;
        if (nw && nh) {
          const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
          if (pos !== undefined) {
            const currentWidth = (this.node.attrs.width as number) || Math.min(nw, 680);
            const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
              ...this.node.attrs,
              naturalWidth: nw,
              naturalHeight: nh,
              width: currentWidth,
              height:
                (this.node.attrs.height as number) ||
                Math.round((currentWidth / nw) * nh),
            });
            this.view.dispatch(tr);
          }
        }
      }
    };

    this.createHandles();
    this.setupEvents();
  }

  private createHandles(): void {
    const corners = ["nw", "ne", "se", "sw"] as const;
    for (const corner of corners) {
      const handle = this.wrapper.createSpan({
        cls: `docx-image-resize-handle handle-${corner}`,
        attr: { "data-corner": corner },
      });
      handle.addEventListener("mousedown", (e) => this.onHandleMouseDown(e, corner));
      this.handles.push(handle);
    }
  }

  private setupEvents(): void {
    this.img.addEventListener("click", (e) => {
      e.stopPropagation();
      this.select();
    });

    document.addEventListener("click", (e) => {
      if (this.isSelected && !this.dom.contains(e.target as Node)) {
        this.deselect();
      }
    });
  }

  select(): void {
    if (this.isSelected) return;
    this.isSelected = true;
    this.wrapper.addClass("is-selected");
    this.showToolbar();
  }

  deselect(): void {
    if (!this.isSelected) return;
    this.isSelected = false;
    this.wrapper.removeClass("is-selected");
    this.hideToolbar();
  }

  selectNode(): void {
    this.select();
  }

  deselectNode(): void {
    this.deselect();
  }

  private showToolbar(): void {
    if (this.toolbar) return;

    this.toolbar = this.wrapper.createDiv({ cls: "docx-image-toolbar" });

    // Size presets
    const presetsGroup = this.toolbar.createDiv({ cls: "docx-img-toolbar-group" });

    const presets = [
      { label: "25%", val: 0.25 },
      { label: "50%", val: 0.5 },
      { label: "75%", val: 0.75 },
      { label: "100%", val: 1.0 },
      { label: "Original", val: "orig" },
    ] as const;

    for (const p of presets) {
      const btn = presetsGroup.createEl("button", {
        cls: "docx-img-btn",
        text: p.label,
        attr: { title: `Scale to ${p.label}` },
      });
      btn.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (p.val === "orig") {
          this.resetOriginalSize();
        } else {
          this.applyScalePreset(p.val);
        }
      };
    }

    // Separator
    this.toolbar.createSpan({ cls: "docx-img-toolbar-sep" });

    // Alignment buttons
    const alignGroup = this.toolbar.createDiv({ cls: "docx-img-toolbar-group" });

    const aligns = [
      { label: "Left", align: "left", title: "Align Left" },
      { label: "Center", align: "center", title: "Align Center" },
      { label: "Right", align: "right", title: "Align Right" },
      { label: "Inline", align: "inline", title: "Inline with text" },
    ] as const;

    for (const a of aligns) {
      const isActive = this.node.attrs.align === a.align;
      const btn = alignGroup.createEl("button", {
        cls: `docx-img-btn ${isActive ? "is-active" : ""}`.trim(),
        text: a.label,
        attr: { title: a.title },
      });
      btn.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.setAlignment(a.align);
      };
    }

    // Separator
    this.toolbar.createSpan({ cls: "docx-img-toolbar-sep" });

    // Action buttons (Alt text, Replace, Delete)
    const actionsGroup = this.toolbar.createDiv({ cls: "docx-img-toolbar-group" });

    const altBtn = actionsGroup.createEl("button", {
      cls: "docx-img-btn",
      text: "Alt Text",
      attr: { title: "Edit Alt Text / Caption" },
    });
    altBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.promptAltText();
    };

    const replaceBtn = actionsGroup.createEl("button", {
      cls: "docx-img-btn",
      text: "Replace",
      attr: { title: "Replace with another image" },
    });
    replaceBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.replaceImage();
    };

    const delBtn = actionsGroup.createEl("button", {
      cls: "docx-img-btn docx-img-btn-danger",
      text: "✕",
      attr: { title: "Delete image" },
    });
    delBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.deleteImage();
    };
  }

  private hideToolbar(): void {
    if (this.toolbar) {
      this.toolbar.remove();
      this.toolbar = null;
    }
  }

  private onHandleMouseDown(e: MouseEvent, corner: "nw" | "ne" | "se" | "sw"): void {
    e.preventDefault();
    e.stopPropagation();
    this.isResizing = true;

    const startX = e.clientX;
    const startWidth = this.img.offsetWidth;
    const startHeight = this.img.offsetHeight;
    const naturalWidth = (this.node.attrs.naturalWidth as number) || this.img.naturalWidth || startWidth;
    const naturalHeight = (this.node.attrs.naturalHeight as number) || this.img.naturalHeight || startHeight;
    const aspectRatio = naturalWidth / (naturalHeight || 1);

    this.showTooltip(`${startWidth} × ${startHeight}px`);

    const onMouseMove = (moveEvent: MouseEvent) => {
      moveEvent.preventDefault();
      const deltaX = moveEvent.clientX - startX;

      let newWidth = startWidth;

      if (corner === "se") {
        newWidth = Math.max(40, startWidth + deltaX);
      } else if (corner === "sw") {
        newWidth = Math.max(40, startWidth - deltaX);
      } else if (corner === "ne") {
        newWidth = Math.max(40, startWidth + deltaX);
      } else if (corner === "nw") {
        newWidth = Math.max(40, startWidth - deltaX);
      }

      // Constrain to container width
      newWidth = Math.min(newWidth, 800);
      const newHeight = Math.round(newWidth / aspectRatio);

      this.img.setCssStyles({
        width: `${newWidth}px`,
        height: `${newHeight}px`,
      });

      const pct = Math.round((newWidth / naturalWidth) * 100);
      this.showTooltip(`${newWidth} × ${newHeight}px (${pct}%)`);
    };

    const onMouseUp = () => {
      this.isResizing = false;
      this.hideTooltip();
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);

      const finalWidth = parseInt(this.img.style.width, 10) || startWidth;
      const finalHeight = parseInt(this.img.style.height, 10) || startHeight;

      const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
      if (pos !== undefined) {
        const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
          ...this.node.attrs,
          width: finalWidth,
          height: finalHeight,
        });
        this.view.dispatch(tr);
      }
    };

    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
  }

  private showTooltip(text: string): void {
    if (!this.tooltip) {
      this.tooltip = this.wrapper.createDiv({ cls: "docx-img-dimension-badge" });
    }
    this.tooltip.setText(text);
  }

  private hideTooltip(): void {
    if (this.tooltip) {
      this.tooltip.remove();
      this.tooltip = null;
    }
  }

  private applyScalePreset(factor: number): void {
    const parent = this.dom.closest(".ProseMirror") || this.dom.parentElement;
    const containerWidth = parent ? (parent as HTMLElement).clientWidth - 64 : 680;
    const targetWidth = Math.round(Math.min(containerWidth, Math.max(60, containerWidth * factor)));
    const nw = (this.node.attrs.naturalWidth as number) || this.img.naturalWidth || targetWidth;
    const nh = (this.node.attrs.naturalHeight as number) || this.img.naturalHeight || targetWidth;
    const targetHeight = Math.round((targetWidth / nw) * nh);

    const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
    if (pos !== undefined) {
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
        ...this.node.attrs,
        width: targetWidth,
        height: targetHeight,
      });
      this.view.dispatch(tr);
    }
  }

  private resetOriginalSize(): void {
    const nw = (this.node.attrs.naturalWidth as number) || this.img.naturalWidth;
    const nh = (this.node.attrs.naturalHeight as number) || this.img.naturalHeight;
    if (!nw || !nh) return;

    const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
    if (pos !== undefined) {
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
        ...this.node.attrs,
        width: nw,
        height: nh,
      });
      this.view.dispatch(tr);
    }
  }

  private setAlignment(align: "left" | "center" | "right" | "inline"): void {
    const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
    if (pos !== undefined) {
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
        ...this.node.attrs,
        align,
      });
      this.view.dispatch(tr);
    }
  }

  private promptAltText(): void {
    const currentAlt = (this.node.attrs.alt as string) || "";
    if (this.app) {
      new TextPromptModal(this.app, "Image Description / Alt Text", currentAlt, (newAlt) => {
        const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
        if (pos !== undefined) {
          const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
            ...this.node.attrs,
            alt: newAlt,
          });
          this.view.dispatch(tr);
        }
      }).open();
    }
  }

  private replaceImage(): void {
    if (!this.docxDoc) return;
    const input = createEl("input", {
      type: "file",
      attr: { accept: "image/*" },
      cls: "office-hidden-file-input",
    });
    input.setCssStyles({ display: "none" });
    input.onchange = () => {
      const file = input.files?.[0];
      if (!file) {
        input.remove();
        return;
      }

      void (async () => {
        try {
          const buffer = await file.arrayBuffer();
          const ext = file.name.split(".").pop() || "png";
          const { rId, dataUrl } = await this.docxDoc!.addMediaFile(buffer, ext);

          const tempImg = new Image();
          tempImg.onload = () => {
            const nw = tempImg.naturalWidth;
            const nh = tempImg.naturalHeight;
            const targetWidth = Math.min(nw, 680);
            const targetHeight = Math.round((targetWidth / nw) * nh);

            const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
            if (pos !== undefined) {
              const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
                ...this.node.attrs,
                rId,
                src: dataUrl,
                drawingXml: "", // will be regenerated cleanly
                width: targetWidth,
                height: targetHeight,
                naturalWidth: nw,
                naturalHeight: nh,
              });
              this.view.dispatch(tr);
            }
          };
          tempImg.src = dataUrl;
        } finally {
          input.remove();
        }
      })();
    };
    document.body.appendChild(input);
    input.click();
  }

  private deleteImage(): void {
    const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
    if (pos !== undefined) {
      const tr = this.view.state.tr.delete(pos, pos + 1);
      this.view.dispatch(tr);
    }
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;

    const src = (this.node.attrs.src as string) || "";
    const alt = (this.node.attrs.alt as string) || "";
    this.img.src = src;
    this.img.alt = alt;

    const width = this.node.attrs.width as number | null;
    const height = this.node.attrs.height as number | null;
    if (width) {
      this.img.setCssStyles({ width: `${width}px` });
    }
    if (height) {
      this.img.setCssStyles({ height: `${height}px` });
    }

    const align = (this.node.attrs.align as string) || "center";
    this.wrapper.className = `docx-image-container docx-image-align-${align}${
      this.isSelected ? " is-selected" : ""
    }`;

    return true;
  }

  destroy(): void {
    this.hideToolbar();
    this.hideTooltip();
  }
}
