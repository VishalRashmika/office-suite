/**
 * Individual PDF Page View component.
 * Combines High-DPI canvas rendering, native selectable text layer,
 * interactive SVG annotation overlay, real-time drag & move repositioning for all components,
 * and text-attached note annotations.
 */

import {
  PdfAnnotation,
  PdfNoteAnnotation,
  PdfHighlightAnnotation,
  PdfInkAnnotation,
  PdfFreeTextAnnotation,
  PdfShapeAnnotation,
  PdfCoordinateTransformer,
  PdfRgbColor,
  PdfRect,
  rgbToHex,
} from "../../domain/pdf";
import { PdfActiveTool } from "./pdf-toolbar";

export interface PdfPageViewDelegate {
  onAddAnnotation(annotation: PdfAnnotation): void;
  onUpdateAnnotation(annotation: PdfAnnotation): void;
  onDeleteAnnotation(id: string): void;
  onClickAnnotation(annotation: PdfAnnotation, element: HTMLElement | SVGElement): void;
  onTextSelected?(selectedText: string, screenRect: DOMRect, pageIndex: number): void;
}

interface DragState {
  annot: PdfAnnotation;
  startScreen: { x: number; y: number };
  startPdfRect: PdfRect;
  startPaths?: number[][][];
  startQuadPoints?: number[];
  startLinePoints?: [number, number, number, number];
  hasMoved: boolean;
}

export class PdfPageView {
  public readonly pageIndex: number;
  public readonly pageNumber: number;

  private containerEl: HTMLElement;
  private canvasEl: HTMLCanvasElement;
  private textLayerEl: HTMLElement;
  private svgOverlayEl: SVGSVGElement;
  private htmlOverlayEl: HTMLElement;
  private onKeyDownListener: ((e: KeyboardEvent) => void) | null = null;
  private delegate: PdfPageViewDelegate | null = null;

  private pdfjsPage: any = null;
  private coordinateTransformer: PdfCoordinateTransformer | null = null;
  private scale = 1.0;
  private currentTool: PdfActiveTool = "hand";
  private currentColor: PdfRgbColor = [0.99, 0.88, 0.28];
  private strokeWidth = 2;
  private isRendering = false;
  private annotations: PdfAnnotation[] = [];

  // Selection & Drag State
  private selectedAnnotId: string | null = null;
  private dragState: DragState | null = null;

  // Live drawing / shape creation state
  private isDrawing = false;
  private activePath: [number, number][] = [];
  private liveDrawEl: SVGPathElement | null = null;
  private shapeStartPoint: { x: number; y: number } | null = null;
  private liveShapeEl: SVGElement | null = null;

  constructor(pageIndex: number, parentEl: HTMLElement) {
    this.pageIndex = pageIndex;
    this.pageNumber = pageIndex + 1;

    this.containerEl = parentEl.createDiv({ cls: "pdf-page-view" });
    this.containerEl.dataset.pageIndex = `${pageIndex}`;

    this.canvasEl = this.containerEl.createEl("canvas", { cls: "pdf-page-canvas" });
    this.textLayerEl = this.containerEl.createDiv({ cls: "pdf-text-layer" });

    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("class", "pdf-annotation-overlay");
    this.containerEl.appendChild(svg);
    this.svgOverlayEl = svg;

    this.htmlOverlayEl = this.containerEl.createDiv({ cls: "pdf-html-annotation-overlay" });

    this.setupPointerEvents();
    this.setupTextSelection();
    this.setupKeyboardShortcuts();
  }

  get element(): HTMLElement {
    return this.containerEl;
  }

  setDelegate(delegate: PdfPageViewDelegate): void {
    this.delegate = delegate;
  }

  setTool(tool: PdfActiveTool, color: PdfRgbColor, width: number): void {
    this.currentTool = tool;
    this.currentColor = color;
    this.strokeWidth = width;

    this.containerEl.dataset.activeTool = tool;

    // In draw and shape creation modes, SVG overlay captures all pointer events to draw
    if (tool === "draw" || tool === "square" || tool === "circle" || tool === "line") {
      this.svgOverlayEl.style.pointerEvents = "auto";
      this.textLayerEl.style.pointerEvents = "none";
    } else {
      // In hand, highlight, note, freetext, eraser:
      // SVG root has pointer-events: none (so textLayer can be selected),
      // while child annotation elements have pointer-events: auto (so they can be clicked and dragged)
      this.svgOverlayEl.style.pointerEvents = "none";
      this.textLayerEl.style.pointerEvents = "auto";
    }
  }

  setAnnotations(annotations: PdfAnnotation[]): void {
    this.annotations = annotations;
    this.renderAnnotations();
  }

  getSelectedAnnotationId(): string | null {
    return this.selectedAnnotId;
  }

  selectAnnotation(id: string | null): void {
    this.selectedAnnotId = id;
    this.renderAnnotations();
  }

  /**
   * Renders the page using PDF.js with High-DPI support.
   */
  async render(pdfjsPage: any, scale: number): Promise<void> {
    if (this.isRendering) return;
    this.isRendering = true;
    this.pdfjsPage = pdfjsPage;
    this.scale = scale;

    try {
      const dpr = window.devicePixelRatio || 1;
      const viewport = pdfjsPage.getViewport({ scale: scale * dpr });
      const cssViewport = pdfjsPage.getViewport({ scale });

      this.coordinateTransformer = new PdfCoordinateTransformer(
        cssViewport.width / scale,
        cssViewport.height / scale,
        (pdfjsPage.rotate % 360) as any
      );

      // Set canvas pixel buffer size for retina displays
      this.canvasEl.width = viewport.width;
      this.canvasEl.height = viewport.height;

      // Set visible CSS dimensions
      const cssWidth = `${Math.floor(cssViewport.width)}px`;
      const cssHeight = `${Math.floor(cssViewport.height)}px`;
      this.containerEl.style.width = cssWidth;
      this.containerEl.style.height = cssHeight;
      this.canvasEl.style.width = cssWidth;
      this.canvasEl.style.height = cssHeight;
      this.textLayerEl.style.width = cssWidth;
      this.textLayerEl.style.height = cssHeight;
      this.svgOverlayEl.setAttribute("width", `${cssViewport.width}`);
      this.svgOverlayEl.setAttribute("height", `${cssViewport.height}`);
      this.htmlOverlayEl.style.width = cssWidth;
      this.htmlOverlayEl.style.height = cssHeight;

      const ctx = this.canvasEl.getContext("2d");
      if (ctx) {
        await pdfjsPage.render({
          canvasContext: ctx,
          viewport,
        }).promise;
      }

      // Render Text Layer for native selection
      await this.renderTextLayer(pdfjsPage, cssViewport);

      // Render Annotations Overlay
      this.renderAnnotations();
    } catch (err) {
      console.error(`Failed to render page ${this.pageNumber}:`, err);
    } finally {
      this.isRendering = false;
    }
  }

  private async renderTextLayer(pdfjsPage: any, viewport: any): Promise<void> {
    this.textLayerEl.empty();
    try {
      const textContent = await pdfjsPage.getTextContent();
      for (const item of textContent.items) {
        if (!("str" in item) || !item.str.trim()) continue;

        const tx = (window as any).pdfjsLib?.Util?.transform?.(viewport.transform, item.transform) || item.transform;
        const span = this.textLayerEl.createSpan({ text: item.str });
        span.style.left = `${tx[4]}px`;
        span.style.top = `${tx[5] - (item.height * this.scale)}px`;
        span.style.fontSize = `${item.height * this.scale}px`;
        span.style.fontFamily = item.fontName || "sans-serif";
      }
    } catch (err) {
      // TextContent extraction fallback
    }
  }

  // ==================== Annotation Overlay Rendering ====================

  renderAnnotations(): void {
    if (!this.coordinateTransformer) return;
    while (this.svgOverlayEl.firstChild) {
      this.svgOverlayEl.removeChild(this.svgOverlayEl.firstChild);
    }

    // Preserve active editing state in HTML overlay if user is actively typing
    const activeEl = document.activeElement;
    const editingBox = activeEl?.closest(".pdf-annot-freetext-box") as HTMLElement | null;
    const editingAnnotId = editingBox?.dataset.annotId || null;

    const existingBoxes = Array.from(this.htmlOverlayEl.querySelectorAll(".pdf-annot-freetext-box"));
    for (const box of existingBoxes) {
      if ((box as HTMLElement).dataset.annotId !== editingAnnotId) {
        box.remove();
      }
    }

    for (const annot of this.annotations) {
      this.renderSingleAnnotation(annot);
    }
  }

  private renderSingleAnnotation(annot: PdfAnnotation): void {
    if (!this.coordinateTransformer) return;
    const colorHex = rgbToHex(annot.color);
    const isSelected = this.selectedAnnotId === annot.id;

    switch (annot.type) {
      case "highlight": {
        this.renderHighlightAnnotation(annot as PdfHighlightAnnotation, colorHex, isSelected);
        break;
      }
      case "text": {
        this.renderNoteAnnotation(annot as PdfNoteAnnotation, colorHex, isSelected);
        break;
      }
      case "ink": {
        this.renderInkAnnotation(annot as PdfInkAnnotation, colorHex, isSelected);
        break;
      }
      case "freetext": {
        this.renderFreeTextAnnotation(annot as PdfFreeTextAnnotation, colorHex, isSelected);
        break;
      }
      case "square":
      case "circle":
      case "line": {
        this.renderShapeAnnotation(annot as PdfShapeAnnotation, colorHex, isSelected);
        break;
      }
    }
  }

  private renderHighlightAnnotation(annot: PdfHighlightAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", `pdf-annot-highlight-group ${isSelected ? "is-selected" : ""}`);
    g.dataset.annotId = annot.id;
    g.style.pointerEvents = "auto";
    g.style.cursor = "pointer";

    if (annot.quadPoints && annot.quadPoints.length >= 8) {
      for (let i = 0; i < annot.quadPoints.length; i += 8) {
        const p1 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i], annot.quadPoints[i + 1], this.scale);
        const p2 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i + 2], annot.quadPoints[i + 3], this.scale);
        const p4 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i + 6], annot.quadPoints[i + 7], this.scale);
        const p3 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i + 4], annot.quadPoints[i + 5], this.scale);

        const poly = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
        poly.setAttribute("points", `${p1.x},${p1.y} ${p2.x},${p2.y} ${p4.x},${p4.y} ${p3.x},${p3.y}`);
        poly.setAttribute("fill", colorHex);
        poly.setAttribute("fill-opacity", "0.38");
        poly.style.mixBlendMode = "multiply";
        g.appendChild(poly);
      }
    } else {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("x", `${screenRect.left}`);
      rect.setAttribute("y", `${screenRect.top}`);
      rect.setAttribute("width", `${screenRect.width}`);
      rect.setAttribute("height", `${screenRect.height}`);
      rect.setAttribute("fill", colorHex);
      rect.setAttribute("fill-opacity", "0.38");
      rect.style.mixBlendMode = "multiply";
      g.appendChild(rect);
    }

    if (isSelected) {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const selBox = this.createSelectionBox(screenRect.left, screenRect.top, screenRect.width, screenRect.height);
      g.appendChild(selBox);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private renderNoteAnnotation(annot: PdfNoteAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);

    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", `pdf-annot-pin-group ${isSelected ? "is-selected" : ""}`);
    g.setAttribute("transform", `translate(${screenRect.left}, ${screenRect.top})`);
    g.dataset.annotId = annot.id;
    g.style.pointerEvents = "auto";
    g.style.cursor = "grab";

    // Pin circle background
    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", "12");
    circle.setAttribute("cy", "12");
    circle.setAttribute("r", isSelected ? "13" : "11");
    circle.setAttribute("fill", colorHex);
    circle.setAttribute("stroke", isSelected ? "var(--interactive-accent)" : "#ffffff");
    circle.setAttribute("stroke-width", isSelected ? "2.5" : "2");
    circle.setAttribute("filter", "drop-shadow(0 2px 5px rgba(0,0,0,0.3))");
    g.appendChild(circle);

    // Comment icon path inside pin
    const icon = document.createElementNS("http://www.w3.org/2000/svg", "path");
    icon.setAttribute("d", "M8 9h8M8 12h5M6 16V7a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H8l-2 2z");
    icon.setAttribute("stroke", "#0f172a");
    icon.setAttribute("stroke-width", "1.5");
    icon.setAttribute("stroke-linecap", "round");
    icon.setAttribute("stroke-linejoin", "round");
    icon.setAttribute("fill", "none");
    g.appendChild(icon);

    if (isSelected) {
      const ring = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      ring.setAttribute("cx", "12");
      ring.setAttribute("cy", "12");
      ring.setAttribute("r", "16");
      ring.setAttribute("fill", "none");
      ring.setAttribute("stroke", "var(--interactive-accent)");
      ring.setAttribute("stroke-width", "1.5");
      ring.setAttribute("stroke-dasharray", "3 3");
      g.appendChild(ring);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private renderInkAnnotation(annot: PdfInkAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", `pdf-annot-ink-group ${isSelected ? "is-selected" : ""}`);
    g.dataset.annotId = annot.id;
    g.style.pointerEvents = "auto";
    g.style.cursor = "move";

    for (const stroke of annot.paths) {
      if (stroke.length < 2) continue;
      const d = this.pointsToSvgPath(stroke);
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", d);
      path.setAttribute("stroke", colorHex);
      path.setAttribute("stroke-width", `${(annot.strokeWidth || 2) * this.scale}`);
      path.setAttribute("stroke-linecap", "round");
      path.setAttribute("stroke-linejoin", "round");
      path.setAttribute("fill", "none");
      g.appendChild(path);
    }

    if (isSelected) {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const selBox = this.createSelectionBox(screenRect.left - 4, screenRect.top - 4, screenRect.width + 8, screenRect.height + 8);
      g.appendChild(selBox);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private renderFreeTextAnnotation(annot: PdfFreeTextAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);

    // If box already exists and is actively being edited, update its position without re-creating DOM
    const existingBox = this.htmlOverlayEl.querySelector(`.pdf-annot-freetext-box[data-annot-id="${annot.id}"]`) as HTMLElement | null;
    if (existingBox) {
      existingBox.style.left = `${screenRect.left}px`;
      existingBox.style.top = `${screenRect.top}px`;
      if (isSelected) {
        existingBox.addClass("is-selected");
      } else {
        existingBox.removeClass("is-selected");
      }
      return;
    }

    const box = this.htmlOverlayEl.createDiv({
      cls: `pdf-annot-freetext-box ${isSelected ? "is-selected" : ""}`,
    });
    box.dataset.annotId = annot.id;
    box.style.left = `${screenRect.left}px`;
    box.style.top = `${screenRect.top}px`;
    box.style.minWidth = `${Math.max(100, screenRect.width)}px`;

    // Handle bar for drag-to-move and delete
    const handleBar = box.createDiv({ cls: "pdf-annot-freetext-handle" });
    const grip = handleBar.createSpan({ cls: "pdf-annot-freetext-grip", text: "⋮⋮" });
    grip.title = "Drag to move text box";

    // Pointer down on handle starts drag
    handleBar.addEventListener("pointerdown", (e: PointerEvent) => {
      e.stopPropagation();
      e.preventDefault();
      this.selectedAnnotId = annot.id;
      const rect = this.containerEl.getBoundingClientRect();
      const sx = e.clientX - rect.left;
      const sy = e.clientY - rect.top;

      this.dragState = {
        annot,
        startScreen: { x: sx, y: sy },
        startPdfRect: [...annot.rect] as PdfRect,
        hasMoved: false,
      };
      this.renderAnnotations();
    });

    // Delete button on handle
    const deleteBtn = handleBar.createSpan({ cls: "pdf-annot-freetext-del", text: "×" });
    deleteBtn.title = "Delete text box";
    deleteBtn.addEventListener("pointerdown", (e: PointerEvent) => {
      e.stopPropagation();
      e.preventDefault();
      this.delegate?.onDeleteAnnotation(annot.id);
      this.selectedAnnotId = null;
    });

    // Content editable element
    const contentEl = box.createDiv({
      cls: "pdf-annot-freetext-content",
      attr: {
        contenteditable: "plaintext-only",
        spellcheck: "false",
        placeholder: "Type text here...",
      },
    });
    contentEl.textContent = annot.contents || "";
    contentEl.style.fontSize = `${Math.max(10, (annot.fontSize || 13) * this.scale)}px`;
    contentEl.style.color = rgbToHex(annot.textColor || [0, 0, 0]);

    // Clicking content element stops propagation so it doesn't trigger container pointerdown or canvas drag
    contentEl.addEventListener("pointerdown", (e: PointerEvent) => {
      e.stopPropagation();
      if (this.selectedAnnotId !== annot.id) {
        this.selectedAnnotId = annot.id;
      }
    });

    contentEl.addEventListener("focus", () => {
      box.addClass("is-editing");
    });

    contentEl.addEventListener("blur", () => {
      box.removeClass("is-editing");
      const text = contentEl.textContent?.trim() || "";
      const boxRect = box.getBoundingClientRect();
      const pageRect = this.containerEl.getBoundingClientRect();

      const screenBounds = {
        left: boxRect.left - pageRect.left,
        top: boxRect.top - pageRect.top,
        width: Math.max(60, boxRect.width),
        height: Math.max(24, boxRect.height),
      };

      if (this.coordinateTransformer) {
        annot.rect = this.coordinateTransformer.screenRectToPdfRect(screenBounds, this.scale);
      }
      annot.contents = text;
      this.delegate?.onUpdateAnnotation(annot);
    });
  }

  private renderShapeAnnotation(annot: PdfShapeAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
    const sw = (annot.strokeWidth || 2) * this.scale;
    const fillColor = annot.fillColor ? rgbToHex(annot.fillColor) : "none";

    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", `pdf-annot-shape-group ${isSelected ? "is-selected" : ""}`);
    g.dataset.annotId = annot.id;
    g.style.pointerEvents = "auto";
    g.style.cursor = "move";

    let shapeEl: SVGElement;
    if (annot.type === "circle") {
      shapeEl = document.createElementNS("http://www.w3.org/2000/svg", "ellipse");
      shapeEl.setAttribute("cx", `${screenRect.left + screenRect.width / 2}`);
      shapeEl.setAttribute("cy", `${screenRect.top + screenRect.height / 2}`);
      shapeEl.setAttribute("rx", `${screenRect.width / 2}`);
      shapeEl.setAttribute("ry", `${screenRect.height / 2}`);
    } else if (annot.type === "line") {
      shapeEl = document.createElementNS("http://www.w3.org/2000/svg", "line");
      shapeEl.setAttribute("x1", `${screenRect.left}`);
      shapeEl.setAttribute("y1", `${screenRect.top}`);
      shapeEl.setAttribute("x2", `${screenRect.left + screenRect.width}`);
      shapeEl.setAttribute("y2", `${screenRect.top + screenRect.height}`);
    } else {
      shapeEl = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      shapeEl.setAttribute("x", `${screenRect.left}`);
      shapeEl.setAttribute("y", `${screenRect.top}`);
      shapeEl.setAttribute("width", `${screenRect.width}`);
      shapeEl.setAttribute("height", `${screenRect.height}`);
    }

    shapeEl.setAttribute("stroke", colorHex);
    shapeEl.setAttribute("stroke-width", `${sw}`);
    shapeEl.setAttribute("fill", fillColor);
    shapeEl.setAttribute("class", "pdf-annot-shape");
    g.appendChild(shapeEl);

    if (isSelected) {
      const selBox = this.createSelectionBox(screenRect.left - 2, screenRect.top - 2, screenRect.width + 4, screenRect.height + 4);
      g.appendChild(selBox);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private createSelectionBox(left: number, top: number, width: number, height: number): SVGGElement {
    const selG = document.createElementNS("http://www.w3.org/2000/svg", "g");
    selG.setAttribute("class", "pdf-annot-selection-handles");

    const selRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    selRect.setAttribute("x", `${left}`);
    selRect.setAttribute("y", `${top}`);
    selRect.setAttribute("width", `${width}`);
    selRect.setAttribute("height", `${height}`);
    selRect.setAttribute("fill", "none");
    selRect.setAttribute("stroke", "var(--interactive-accent)");
    selRect.setAttribute("stroke-width", "1.5");
    selRect.setAttribute("stroke-dasharray", "4 4");
    selG.appendChild(selRect);

    return selG;
  }

  // ==================== Interactive Pointer, Drag & Move Events ====================

  private setupPointerEvents(): void {
    // Listen for pointerdown on both svg overlay and container
    this.containerEl.addEventListener("pointerdown", this.onPointerDown);
    window.addEventListener("pointermove", this.onPointerMove);
    window.addEventListener("pointerup", this.onPointerUp);
  }

  private onPointerDown = (e: PointerEvent): void => {
    if (!this.coordinateTransformer) return;
    const targetEl = e.target as HTMLElement | SVGElement;

    // Check if an existing annotation was clicked/pressed
    const annotEl = targetEl.closest("[data-annot-id]") as HTMLElement | SVGElement | null;

    const rect = this.containerEl.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    if (annotEl) {
      const annotId = annotEl.dataset.annotId;
      const annot = this.annotations.find((a) => a.id === annotId);

      if (annot) {
        // Eraser tool: delete immediately
        if (this.currentTool === "eraser") {
          e.stopPropagation();
          e.preventDefault();
          this.delegate?.onDeleteAnnotation(annot.id);
          this.selectedAnnotId = null;
          return;
        }

        // Existing annotation hit! DO NOT create a new note/shape/text.
        e.stopPropagation();
        this.selectedAnnotId = annot.id;

        // Prepare drag state
        this.dragState = {
          annot,
          startScreen: { x: sx, y: sy },
          startPdfRect: [...annot.rect] as PdfRect,
          startPaths: annot.type === "ink" ? JSON.parse(JSON.stringify((annot as PdfInkAnnotation).paths)) : undefined,
          startQuadPoints: annot.type === "highlight" ? [...(annot as PdfHighlightAnnotation).quadPoints] : undefined,
          startLinePoints: (annot as any).linePoints ? [...(annot as any).linePoints] : undefined,
          hasMoved: false,
        };

        this.renderAnnotations();
        return;
      }
    }

    // No existing annotation clicked: user clicked on empty space
    if (this.currentTool === "draw") {
      this.isDrawing = true;
      this.activePath = [[sx, sy]];

      this.liveDrawEl = document.createElementNS("http://www.w3.org/2000/svg", "path");
      this.liveDrawEl.setAttribute("stroke", rgbToHex(this.currentColor));
      this.liveDrawEl.setAttribute("stroke-width", `${this.strokeWidth * this.scale}`);
      this.liveDrawEl.setAttribute("stroke-linecap", "round");
      this.liveDrawEl.setAttribute("stroke-linejoin", "round");
      this.liveDrawEl.setAttribute("fill", "none");
      this.svgOverlayEl.appendChild(this.liveDrawEl);
    } else if (this.currentTool === "square" || this.currentTool === "circle" || this.currentTool === "line") {
      this.isDrawing = true;
      this.shapeStartPoint = { x: sx, y: sy };
      this.createLiveShapeElement(sx, sy);
    } else if (this.currentTool === "note") {
      // Place new sticky note on empty space
      this.placeStickyNote(sx, sy);
    } else if (this.currentTool === "freetext") {
      // Place new text box on empty space
      this.placeFreeTextBox(sx, sy);
    } else if (this.currentTool === "hand") {
      // Deselect
      if (this.selectedAnnotId) {
        this.selectedAnnotId = null;
        this.renderAnnotations();
      }
    }
  };

  private onPointerMove = (e: PointerEvent): void => {
    const rect = this.containerEl.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // 1. Dragging an existing annotation
    if (this.dragState && this.coordinateTransformer) {
      const dx = sx - this.dragState.startScreen.x;
      const dy = sy - this.dragState.startScreen.y;

      if (!this.dragState.hasMoved && Math.hypot(dx, dy) > 3) {
        this.dragState.hasMoved = true;
        document.body.style.cursor = "grabbing";
      }

      if (this.dragState.hasMoved) {
        const pStart = this.coordinateTransformer.screenToPdf(
          this.dragState.startScreen.x,
          this.dragState.startScreen.y,
          this.scale
        );
        const pCurrent = this.coordinateTransformer.screenToPdf(sx, sy, this.scale);
        const dPdfX = pCurrent.x - pStart.x;
        const dPdfY = pCurrent.y - pStart.y;

        const annot = this.dragState.annot;
        const startRect = this.dragState.startPdfRect;

        // Update PDF rect
        annot.rect = [
          startRect[0] + dPdfX,
          startRect[1] + dPdfY,
          startRect[2] + dPdfX,
          startRect[3] + dPdfY,
        ];

        // Update ink paths
        if (annot.type === "ink" && this.dragState.startPaths) {
          (annot as PdfInkAnnotation).paths = this.dragState.startPaths.map((stroke) =>
            stroke.map(([x, y]) => [x + dPdfX, y + dPdfY])
          );
        }

        // Update highlight quadpoints
        if (annot.type === "highlight" && this.dragState.startQuadPoints) {
          (annot as PdfHighlightAnnotation).quadPoints = this.dragState.startQuadPoints.map((val, idx) =>
            idx % 2 === 0 ? val + dPdfX : val + dPdfY
          );
        }

        // Update line points
        if (annot.type === "line" && this.dragState.startLinePoints) {
          (annot as any).linePoints = [
            this.dragState.startLinePoints[0] + dPdfX,
            this.dragState.startLinePoints[1] + dPdfY,
            this.dragState.startLinePoints[2] + dPdfX,
            this.dragState.startLinePoints[3] + dPdfY,
          ];
        }

        this.renderAnnotations();
      }
      return;
    }

    // 2. Active drawing or shape creation
    if (!this.isDrawing) return;

    if (this.currentTool === "draw" && this.liveDrawEl) {
      this.activePath.push([sx, sy]);
      this.liveDrawEl.setAttribute("d", this.screenPointsToSvgPath(this.activePath));
    } else if (this.shapeStartPoint && this.liveShapeEl) {
      this.updateLiveShapeElement(this.shapeStartPoint.x, this.shapeStartPoint.y, sx, sy);
    }
  };

  private onPointerUp = (e: PointerEvent): void => {
    document.body.style.cursor = "";

    // 1. Handling Drag End on existing annotation
    if (this.dragState) {
      const annot = this.dragState.annot;
      const hasMoved = this.dragState.hasMoved;
      this.dragState = null;

      if (hasMoved) {
        // Drag moved: commit updated coordinates
        this.delegate?.onUpdateAnnotation(annot);
        this.renderAnnotations();
        return;
      } else {
        // Just clicked on the existing annotation (without dragging)
        if (annot.type === "text" || annot.type === "highlight") {
          const annotEl = this.containerEl.querySelector(`[data-annot-id="${annot.id}"]`) as HTMLElement | SVGElement | null;
          this.delegate?.onClickAnnotation(annot, annotEl || this.containerEl);
        }
        return;
      }
    }

    // 2. Drawing or shape creation end
    if (!this.isDrawing) return;
    this.isDrawing = false;

    if (this.currentTool === "draw" && this.activePath.length > 1 && this.coordinateTransformer) {
      this.finishDrawing();
    } else if (this.shapeStartPoint && this.coordinateTransformer) {
      const rect = this.containerEl.getBoundingClientRect();
      const endX = e.clientX - rect.left;
      const endY = e.clientY - rect.top;
      this.finishShape(this.shapeStartPoint.x, this.shapeStartPoint.y, endX, endY);
    }

    if (this.liveDrawEl) {
      this.liveDrawEl.remove();
      this.liveDrawEl = null;
    }
    if (this.liveShapeEl) {
      this.liveShapeEl.remove();
      this.liveShapeEl = null;
    }
    this.activePath = [];
    this.shapeStartPoint = null;
  };

  private finishDrawing(): void {
    if (!this.coordinateTransformer) return;
    const pdfPath = this.activePath.map(([sx, sy]) => {
      const pt = this.coordinateTransformer!.screenToPdf(sx, sy, this.scale);
      return [pt.x, pt.y];
    });

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const [x, y] of pdfPath) {
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x > maxX) maxX = x;
      if (y > maxY) maxY = y;
    }

    const annot: PdfInkAnnotation = {
      id: `ink_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "ink",
      pageIndex: this.pageIndex,
      rect: [minX - 5, minY - 5, maxX + 5, maxY + 5],
      color: this.currentColor,
      paths: [pdfPath],
      strokeWidth: this.strokeWidth,
      author: "User",
      createdAt: new Date(),
    };

    this.delegate?.onAddAnnotation(annot);
  }

  private createLiveShapeElement(sx: number, sy: number): void {
    const sw = this.strokeWidth * this.scale;
    const color = rgbToHex(this.currentColor);

    if (this.currentTool === "circle") {
      this.liveShapeEl = document.createElementNS("http://www.w3.org/2000/svg", "ellipse");
    } else if (this.currentTool === "line") {
      this.liveShapeEl = document.createElementNS("http://www.w3.org/2000/svg", "line");
    } else {
      this.liveShapeEl = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    }

    this.liveShapeEl.setAttribute("stroke", color);
    this.liveShapeEl.setAttribute("stroke-width", `${sw}`);
    this.liveShapeEl.setAttribute("fill", "none");
    this.svgOverlayEl.appendChild(this.liveShapeEl);
  }

  private updateLiveShapeElement(x1: number, y1: number, x2: number, y2: number): void {
    if (!this.liveShapeEl) return;
    const left = Math.min(x1, x2);
    const top = Math.min(y1, y2);
    const width = Math.abs(x2 - x1);
    const height = Math.abs(y2 - y1);

    if (this.currentTool === "circle") {
      this.liveShapeEl.setAttribute("cx", `${left + width / 2}`);
      this.liveShapeEl.setAttribute("cy", `${top + height / 2}`);
      this.liveShapeEl.setAttribute("rx", `${width / 2}`);
      this.liveShapeEl.setAttribute("ry", `${height / 2}`);
    } else if (this.currentTool === "line") {
      this.liveShapeEl.setAttribute("x1", `${x1}`);
      this.liveShapeEl.setAttribute("y1", `${y1}`);
      this.liveShapeEl.setAttribute("x2", `${x2}`);
      this.liveShapeEl.setAttribute("y2", `${y2}`);
    } else {
      this.liveShapeEl.setAttribute("x", `${left}`);
      this.liveShapeEl.setAttribute("y", `${top}`);
      this.liveShapeEl.setAttribute("width", `${width}`);
      this.liveShapeEl.setAttribute("height", `${height}`);
    }
  }

  private finishShape(x1: number, y1: number, x2: number, y2: number): void {
    if (!this.coordinateTransformer) return;
    const left = Math.min(x1, x2);
    const top = Math.min(y1, y2);
    const width = Math.max(12, Math.abs(x2 - x1));
    const height = Math.max(12, Math.abs(y2 - y1));

    const pdfRect = this.coordinateTransformer.screenRectToPdfRect({ left, top, width, height }, this.scale);
    const shapeType = this.currentTool as "square" | "circle" | "line";

    const annot: PdfShapeAnnotation = {
      id: `shape_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: shapeType,
      pageIndex: this.pageIndex,
      rect: pdfRect,
      color: this.currentColor,
      strokeColor: this.currentColor,
      strokeWidth: this.strokeWidth,
      author: "User",
      createdAt: new Date(),
    };

    this.delegate?.onAddAnnotation(annot);
  }

  private placeStickyNote(sx: number, sy: number): void {
    if (!this.coordinateTransformer) return;
    const pt = this.coordinateTransformer.screenToPdf(sx, sy, this.scale);
    const pdfRect: PdfRect = [pt.x, pt.y - 24, pt.x + 24, pt.y];

    const annot: PdfNoteAnnotation = {
      id: `note_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "text",
      pageIndex: this.pageIndex,
      rect: pdfRect,
      color: this.currentColor,
      contents: "",
      author: "User",
      iconName: "Comment",
      createdAt: new Date(),
    };

    this.delegate?.onAddAnnotation(annot);
    setTimeout(() => {
      const pinEl = this.containerEl.querySelector(`[data-annot-id="${annot.id}"]`) as HTMLElement | SVGElement | null;
      this.delegate?.onClickAnnotation(annot, pinEl || this.containerEl);
    }, 60);
  }

  private placeFreeTextBox(sx: number, sy: number): void {
    if (!this.coordinateTransformer) return;
    const pt = this.coordinateTransformer.screenToPdf(sx, sy, this.scale);
    const pdfRect: PdfRect = [pt.x, pt.y - 28, pt.x + 140, pt.y];

    const annot: PdfFreeTextAnnotation = {
      id: `freetext_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "freetext",
      pageIndex: this.pageIndex,
      rect: pdfRect,
      color: this.currentColor,
      textColor: [0, 0, 0],
      fontSize: 13,
      contents: "",
      author: "User",
      createdAt: new Date(),
    };

    this.selectedAnnotId = annot.id;
    this.delegate?.onAddAnnotation(annot);

    // Auto-focus immediately so user can type right away!
    setTimeout(() => {
      const box = this.htmlOverlayEl.querySelector(`.pdf-annot-freetext-box[data-annot-id="${annot.id}"]`);
      if (box) {
        const contentEl = box.querySelector(".pdf-annot-freetext-content") as HTMLElement | null;
        if (contentEl) {
          contentEl.focus();
        }
      }
    }, 50);
  }

  // ==================== Text Selection Support ====================

  private setupTextSelection(): void {
    this.textLayerEl.addEventListener("mouseup", () => {
      const sel = window.getSelection();
      if (!sel || sel.isCollapsed || !sel.rangeCount) return;

      const text = sel.toString().trim();
      if (!text) return;

      const range = sel.getRangeAt(0);
      const rect = range.getBoundingClientRect();
      this.delegate?.onTextSelected?.(text, rect, this.pageIndex);
    });
  }

  private setupKeyboardShortcuts(): void {
    this.onKeyDownListener = (e: KeyboardEvent) => {
      if ((e.key === "Delete" || e.key === "Backspace") && this.selectedAnnotId) {
        const activeTag = document.activeElement?.tagName.toLowerCase();
        if (
          activeTag === "input" ||
          activeTag === "textarea" ||
          (document.activeElement as HTMLElement)?.isContentEditable
        ) {
          return;
        }
        this.delegate?.onDeleteAnnotation(this.selectedAnnotId);
        this.selectedAnnotId = null;
        this.renderAnnotations();
      }
    };
    window.addEventListener("keydown", this.onKeyDownListener);
  }

  createHighlightFromSelection(selectedText: string, color: PdfRgbColor): void {
    if (!this.coordinateTransformer) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    const range = sel.getRangeAt(0);
    const clientRects = Array.from(range.getClientRects());
    if (clientRects.length === 0) return;

    const pageRect = this.containerEl.getBoundingClientRect();
    const quadPoints: number[] = [];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    for (const r of clientRects) {
      const screenLeft = r.left - pageRect.left;
      const screenTop = r.top - pageRect.top;
      const screenRight = r.right - pageRect.left;
      const screenBottom = r.bottom - pageRect.top;

      const pTopLeft = this.coordinateTransformer.screenToPdf(screenLeft, screenTop, this.scale);
      const pTopRight = this.coordinateTransformer.screenToPdf(screenRight, screenTop, this.scale);
      const pBottomLeft = this.coordinateTransformer.screenToPdf(screenLeft, screenBottom, this.scale);
      const pBottomRight = this.coordinateTransformer.screenToPdf(screenRight, screenBottom, this.scale);

      quadPoints.push(
        pTopLeft.x, pTopLeft.y,
        pTopRight.x, pTopRight.y,
        pBottomLeft.x, pBottomLeft.y,
        pBottomRight.x, pBottomRight.y
      );

      minX = Math.min(minX, pTopLeft.x, pBottomLeft.x);
      minY = Math.min(minY, pBottomLeft.y, pBottomRight.y);
      maxX = Math.max(maxX, pTopRight.x, pBottomRight.x);
      maxY = Math.max(maxY, pTopLeft.y, pTopRight.y);
    }

    const annot: PdfHighlightAnnotation = {
      id: `highlight_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "highlight",
      pageIndex: this.pageIndex,
      rect: [minX, minY, maxX, maxY],
      quadPoints,
      color,
      contents: selectedText,
      highlightedText: selectedText,
      author: "User",
      createdAt: new Date(),
    };

    sel.removeAllRanges();
    this.delegate?.onAddAnnotation(annot);
  }

  /**
   * Creates a text-attached annotated note directly on the selected sentence or phrase.
   * Immediately opens the note popover so the user can type their note!
   */
  addAnnotatedNoteFromSelection(selectedText: string, color: PdfRgbColor): void {
    if (!this.coordinateTransformer) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    const range = sel.getRangeAt(0);
    const clientRects = Array.from(range.getClientRects());
    if (clientRects.length === 0) return;

    const pageRect = this.containerEl.getBoundingClientRect();
    const quadPoints: number[] = [];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    for (const r of clientRects) {
      const screenLeft = r.left - pageRect.left;
      const screenTop = r.top - pageRect.top;
      const screenRight = r.right - pageRect.left;
      const screenBottom = r.bottom - pageRect.top;

      const pTopLeft = this.coordinateTransformer.screenToPdf(screenLeft, screenTop, this.scale);
      const pTopRight = this.coordinateTransformer.screenToPdf(screenRight, screenTop, this.scale);
      const pBottomLeft = this.coordinateTransformer.screenToPdf(screenLeft, screenBottom, this.scale);
      const pBottomRight = this.coordinateTransformer.screenToPdf(screenRight, screenBottom, this.scale);

      quadPoints.push(
        pTopLeft.x, pTopLeft.y,
        pTopRight.x, pTopRight.y,
        pBottomLeft.x, pBottomLeft.y,
        pBottomRight.x, pBottomRight.y
      );

      minX = Math.min(minX, pTopLeft.x, pBottomLeft.x);
      minY = Math.min(minY, pBottomLeft.y, pBottomRight.y);
      maxX = Math.max(maxX, pTopRight.x, pBottomRight.x);
      maxY = Math.max(maxY, pTopLeft.y, pTopRight.y);
    }

    const annot: PdfHighlightAnnotation = {
      id: `note_hl_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "highlight",
      pageIndex: this.pageIndex,
      rect: [minX, minY, maxX, maxY],
      quadPoints,
      color,
      contents: "",
      highlightedText: selectedText,
      author: "User",
      createdAt: new Date(),
    };

    sel.removeAllRanges();
    this.delegate?.onAddAnnotation(annot);

    // Open note popover anchored to the newly created highlight
    setTimeout(() => {
      const annotEl = this.containerEl.querySelector(`[data-annot-id="${annot.id}"]`) as HTMLElement | SVGElement | null;
      this.delegate?.onClickAnnotation(annot, annotEl || this.containerEl);
    }, 60);
  }

  // ==================== Path Smoothing Helpers ====================

  private pointsToSvgPath(stroke: number[][]): string {
    if (!this.coordinateTransformer || stroke.length === 0) return "";
    const screenPts = stroke.map(([x, y]) => this.coordinateTransformer!.pdfToScreen(x, y, this.scale));
    return this.screenPointsToSvgPath(screenPts.map((p) => [p.x, p.y]));
  }

  private screenPointsToSvgPath(pts: [number, number][]): string {
    if (pts.length === 0) return "";
    if (pts.length === 1) return `M ${pts[0][0]} ${pts[0][1]} L ${pts[0][0]} ${pts[0][1]}`;

    let d = `M ${pts[0][0]} ${pts[0][1]}`;
    for (let i = 1; i < pts.length; i++) {
      const midX = (pts[i - 1][0] + pts[i][0]) / 2;
      const midY = (pts[i - 1][1] + pts[i][1]) / 2;
      d += ` Q ${pts[i - 1][0]} ${pts[i - 1][1]}, ${midX} ${midY}`;
    }
    d += ` L ${pts[pts.length - 1][0]} ${pts[pts.length - 1][1]}`;
    return d;
  }

  pulseAnnotation(annotId: string): void {
    const el = this.containerEl.querySelector(`[data-annot-id="${annotId}"]`);
    if (el) {
      el.classList.add("is-pulsing");
      setTimeout(() => el.classList.remove("is-pulsing"), 1500);
    }
  }

  dispose(): void {
    this.containerEl.removeEventListener("pointerdown", this.onPointerDown);
    window.removeEventListener("pointermove", this.onPointerMove);
    window.removeEventListener("pointerup", this.onPointerUp);
    if (this.onKeyDownListener) {
      window.removeEventListener("keydown", this.onKeyDownListener);
      this.onKeyDownListener = null;
    }
    this.containerEl.remove();
    this.delegate = null;
    this.pdfjsPage = null;
  }
}
