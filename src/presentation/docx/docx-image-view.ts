import { Node as PMNode } from "prosemirror-model";
import { EditorView, NodeView } from "prosemirror-view";
import { DocxDocument } from "../../infrastructure/docx/docx-document";

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
    private getPos: () => number | undefined,
    private docxDoc: DocxDocument | null
  ) {
    this.dom = document.createElement("span");
    this.dom.className = "docx-image-wrapper-outer";
    this.dom.setAttribute("contenteditable", "false");

    this.wrapper = document.createElement("span");
    this.wrapper.className = `docx-image-container docx-image-align-${this.node.attrs.align || "center"}`;
    this.dom.appendChild(this.wrapper);

    this.img = document.createElement("img");
    this.img.className = "docx-image-element";
    this.img.src = this.node.attrs.src || "";
    this.img.alt = this.node.attrs.alt || "";

    if (this.node.attrs.width) {
      this.img.style.width = `${this.node.attrs.width}px`;
    }
    if (this.node.attrs.height) {
      this.img.style.height = `${this.node.attrs.height}px`;
    }

    this.wrapper.appendChild(this.img);

    // Natural dimension capture if not yet stored
    this.img.onload = () => {
      if (!this.node.attrs.naturalWidth || !this.node.attrs.naturalHeight) {
        const nw = this.img.naturalWidth;
        const nh = this.img.naturalHeight;
        if (nw && nh) {
          const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
          if (pos !== undefined) {
            const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
              ...this.node.attrs,
              naturalWidth: nw,
              naturalHeight: nh,
              width: this.node.attrs.width || Math.min(nw, 680),
              height:
                this.node.attrs.height ||
                Math.round(((this.node.attrs.width || Math.min(nw, 680)) / nw) * nh),
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
      const handle = document.createElement("span");
      handle.className = `docx-image-resize-handle handle-${corner}`;
      handle.setAttribute("data-corner", corner);
      handle.addEventListener("mousedown", (e) => this.onHandleMouseDown(e, corner));
      this.wrapper.appendChild(handle);
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
    this.wrapper.classList.add("is-selected");
    this.showToolbar();
  }

  deselect(): void {
    if (!this.isSelected) return;
    this.isSelected = false;
    this.wrapper.classList.remove("is-selected");
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

    this.toolbar = document.createElement("div");
    this.toolbar.className = "docx-image-toolbar";

    // Size presets
    const presetsGroup = document.createElement("div");
    presetsGroup.className = "docx-img-toolbar-group";

    const presets = [
      { label: "25%", val: 0.25 },
      { label: "50%", val: 0.5 },
      { label: "75%", val: 0.75 },
      { label: "100%", val: 1.0 },
      { label: "Original", val: "orig" },
    ] as const;

    for (const p of presets) {
      const btn = document.createElement("button");
      btn.className = "docx-img-btn";
      btn.textContent = p.label;
      btn.title = `Scale to ${p.label}`;
      btn.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (p.val === "orig") {
          this.resetOriginalSize();
        } else {
          this.applyScalePreset(p.val);
        }
      };
      presetsGroup.appendChild(btn);
    }
    this.toolbar.appendChild(presetsGroup);

    // Separator
    const sep1 = document.createElement("span");
    sep1.className = "docx-img-toolbar-sep";
    this.toolbar.appendChild(sep1);

    // Alignment buttons
    const alignGroup = document.createElement("div");
    alignGroup.className = "docx-img-toolbar-group";

    const aligns = [
      { label: "Left", align: "left", title: "Align Left" },
      { label: "Center", align: "center", title: "Align Center" },
      { label: "Right", align: "right", title: "Align Right" },
      { label: "Inline", align: "inline", title: "Inline with text" },
    ] as const;

    for (const a of aligns) {
      const btn = document.createElement("button");
      btn.className = `docx-img-btn ${this.node.attrs.align === a.align ? "is-active" : ""}`;
      btn.textContent = a.label;
      btn.title = a.title;
      btn.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.setAlignment(a.align);
      };
      alignGroup.appendChild(btn);
    }
    this.toolbar.appendChild(alignGroup);

    // Separator
    const sep2 = document.createElement("span");
    sep2.className = "docx-img-toolbar-sep";
    this.toolbar.appendChild(sep2);

    // Action buttons (Alt text, Replace, Delete)
    const actionsGroup = document.createElement("div");
    actionsGroup.className = "docx-img-toolbar-group";

    const altBtn = document.createElement("button");
    altBtn.className = "docx-img-btn";
    altBtn.textContent = "Alt Text";
    altBtn.title = "Edit Alt Text / Caption";
    altBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.promptAltText();
    };
    actionsGroup.appendChild(altBtn);

    const replaceBtn = document.createElement("button");
    replaceBtn.className = "docx-img-btn";
    replaceBtn.textContent = "Replace";
    replaceBtn.title = "Replace with another image";
    replaceBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.replaceImage();
    };
    actionsGroup.appendChild(replaceBtn);

    const delBtn = document.createElement("button");
    delBtn.className = "docx-img-btn docx-img-btn-danger";
    delBtn.textContent = "✕";
    delBtn.title = "Delete image";
    delBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.deleteImage();
    };
    actionsGroup.appendChild(delBtn);

    this.toolbar.appendChild(actionsGroup);
    this.wrapper.appendChild(this.toolbar);
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
    const startY = e.clientY;
    const startWidth = this.img.offsetWidth;
    const startHeight = this.img.offsetHeight;
    const naturalWidth = this.node.attrs.naturalWidth || this.img.naturalWidth || startWidth;
    const naturalHeight = this.node.attrs.naturalHeight || this.img.naturalHeight || startHeight;
    const aspectRatio = naturalWidth / (naturalHeight || 1);

    this.showTooltip(`${startWidth} × ${startHeight}px`);

    const onMouseMove = (moveEvent: MouseEvent) => {
      moveEvent.preventDefault();
      const deltaX = moveEvent.clientX - startX;
      const deltaY = moveEvent.clientY - startY;

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

      this.img.style.width = `${newWidth}px`;
      this.img.style.height = `${newHeight}px`;

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
      this.tooltip = document.createElement("div");
      this.tooltip.className = "docx-img-dimension-badge";
      this.wrapper.appendChild(this.tooltip);
    }
    this.tooltip.textContent = text;
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
    const nw = this.node.attrs.naturalWidth || this.img.naturalWidth || targetWidth;
    const nh = this.node.attrs.naturalHeight || this.img.naturalHeight || targetWidth;
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
    const nw = this.node.attrs.naturalWidth || this.img.naturalWidth;
    const nh = this.node.attrs.naturalHeight || this.img.naturalHeight;
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
    const currentAlt = this.node.attrs.alt || "";
    const newAlt = window.prompt("Image description / Alt text:", currentAlt);
    if (newAlt === null) return;

    const pos = typeof this.getPos === "function" ? this.getPos() : undefined;
    if (pos !== undefined) {
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
        ...this.node.attrs,
        alt: newAlt,
      });
      this.view.dispatch(tr);
    }
  }

  private replaceImage(): void {
    if (!this.docxDoc) return;
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.style.display = "none";
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return;

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
    };
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

    this.img.src = this.node.attrs.src || "";
    this.img.alt = this.node.attrs.alt || "";

    if (this.node.attrs.width) {
      this.img.style.width = `${this.node.attrs.width}px`;
    }
    if (this.node.attrs.height) {
      this.img.style.height = `${this.node.attrs.height}px`;
    }

    this.wrapper.className = `docx-image-container docx-image-align-${this.node.attrs.align || "center"}${
      this.isSelected ? " is-selected" : ""
    }`;

    return true;
  }

  destroy(): void {
    this.hideToolbar();
    this.hideTooltip();
  }
}
