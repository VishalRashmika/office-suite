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
  PdfUnderlineAnnotation,
  PdfInkAnnotation,
  PdfFreeTextAnnotation,
  PdfShapeAnnotation,
  PdfCoordinateTransformer,
  PdfRgbColor,
  PdfRect,
  PDF_COLOR_PRESETS,
  rgbToHex,
  hexToRgb,
  PdfColorSchemeProfile,
  normalizeHexColor,
} from "../../domain/pdf";
import { PdfActiveTool } from "./pdf-toolbar";

export interface PdfPageViewDelegate {
  onAddAnnotation(annotation: PdfAnnotation): void;
  onUpdateAnnotation(annotation: PdfAnnotation): void;
  onDeleteAnnotation(id: string): void;
  onClickAnnotation(annotation: PdfAnnotation, element: HTMLElement | SVGElement): void;
  onTextSelected?(selectedText: string, screenRect: DOMRect, pageIndex: number): void;
  getPageView?(pageIndex: number): PdfPageView | undefined;
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
  private colorScheme?: PdfColorSchemeProfile;

  // Selection & Drag State
  private selectedAnnotId: string | null = null;
  private dragState: DragState | null = null;

  // Live drawing / shape creation state
  private isDrawing = false;
  private activePath: [number, number][] = [];
  private liveDrawEl: SVGPathElement | null = null;
  private shapeStartPoint: { x: number; y: number } | null = null;
  private liveShapeEl: SVGElement | null = null;

  // FreeText creation state
  private isCreatingFreeText = false;
  private freeTextStartPoint: { x: number; y: number } | null = null;
  private liveFreeTextEl: HTMLElement | null = null;

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

  getCoordinateTransformer(): PdfCoordinateTransformer | null {
    return this.coordinateTransformer;
  }

  setDelegate(delegate: PdfPageViewDelegate): void {
    this.delegate = delegate;
  }

  setColorScheme(scheme: PdfColorSchemeProfile): void {
    this.colorScheme = scheme;
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
    const pdfjsLib = (window as any).pdfjsLib;

    try {
      const textContent = await pdfjsPage.getTextContent();

      // Try official PDF.js text layer if available
      if (pdfjsLib?.renderTextLayer) {
        try {
          await pdfjsLib.renderTextLayer({
            textContent,
            container: this.textLayerEl,
            viewport,
            textDivs: [],
          }).promise;
          return;
        } catch (e) {
          // fall through to precise manual renderer
        }
      } else if (pdfjsLib?.TextLayer) {
        try {
          const tl = new pdfjsLib.TextLayer({
            textContentSource: textContent,
            container: this.textLayerEl,
            viewport,
          });
          await tl.render();
          return;
        } catch (e) {
          // fall through to precise manual renderer
        }
      }

      // Precise manual renderer with subpixel glyph matching
      for (const item of textContent.items) {
        if (!("str" in item) || !item.str) continue;

        const tx = pdfjsLib?.Util?.transform?.(viewport.transform, item.transform) || item.transform;
        const fontHeight = Math.hypot(tx[2], tx[3]) || (item.height * this.scale) || 12;
        const fontWidth = item.width * this.scale;

        const span = this.textLayerEl.createSpan({ text: item.str });
        span.style.left = `${tx[4]}px`;
        span.style.top = `${tx[5] - fontHeight}px`;
        span.style.fontSize = `${fontHeight}px`;
        span.style.fontFamily = item.fontName ? `"${item.fontName}", sans-serif` : "sans-serif";

        const naturalWidth = span.getBoundingClientRect().width;
        if (naturalWidth > 0 && fontWidth > 0) {
          const scaleX = fontWidth / naturalWidth;
          span.style.transform = `scaleX(${scaleX})`;
          span.style.transformOrigin = `0% 0%`;
          span.style.width = `${naturalWidth}px`;
        }
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

    // Clean up FreeText DOM boxes that no longer belong to this page's annotations
    const currentIds = new Set(this.annotations.map((a) => a.id));
    const existingBoxes = Array.from(this.htmlOverlayEl.querySelectorAll(".pdf-annot-freetext-box")) as HTMLElement[];
    for (const box of existingBoxes) {
      const annotId = box.dataset.annotId;
      if (!annotId || !currentIds.has(annotId)) {
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
      case "underline": {
        this.renderUnderlineAnnotation(annot as PdfUnderlineAnnotation, colorHex, isSelected);
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

    // Subtle Zotero Bookmark indicator if this highlight has an attached note or comment
    const hasNote = Boolean(
      (annot.contents && annot.contents.trim() !== "" && annot.contents.trim() !== annot.highlightedText?.trim()) ||
      annot.linkedNote
    );
    if (hasNote) {
      let bookmarkX = 0;
      let bookmarkY = 0;
      if (annot.quadPoints && annot.quadPoints.length >= 8) {
        const p1 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[0], annot.quadPoints[1], this.scale);
        bookmarkX = p1.x - 3;
        bookmarkY = p1.y - 13;
      } else {
        const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
        bookmarkX = screenRect.left - 3;
        bookmarkY = screenRect.top - 13;
      }
      const bookmarkG = this.createBookmarkIndicator(bookmarkX, bookmarkY, colorHex, annot);
      g.appendChild(bookmarkG);
    }

    if (isSelected) {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const selBox = this.createSelectionBox(screenRect.left, screenRect.top, screenRect.width, screenRect.height);
      g.appendChild(selBox);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private renderUnderlineAnnotation(annot: PdfUnderlineAnnotation, colorHex: string, isSelected: boolean): void {
    if (!this.coordinateTransformer) return;
    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", `pdf-annot-underline-group ${isSelected ? "is-selected" : ""}`);
    g.dataset.annotId = annot.id;
    g.style.pointerEvents = "auto";
    g.style.cursor = "pointer";

    if (annot.quadPoints && annot.quadPoints.length >= 8) {
      for (let i = 0; i < annot.quadPoints.length; i += 8) {
        const pBottomLeft = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i + 4], annot.quadPoints[i + 5], this.scale);
        const pBottomRight = this.coordinateTransformer.pdfToScreen(annot.quadPoints[i + 6], annot.quadPoints[i + 7], this.scale);

        // Invisible wider stroke for easy click hit detection
        const hitLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
        hitLine.setAttribute("x1", `${pBottomLeft.x}`);
        hitLine.setAttribute("y1", `${pBottomLeft.y}`);
        hitLine.setAttribute("x2", `${pBottomRight.x}`);
        hitLine.setAttribute("y2", `${pBottomRight.y}`);
        hitLine.setAttribute("stroke", "transparent");
        hitLine.setAttribute("stroke-width", "14");
        hitLine.style.pointerEvents = "stroke";
        g.appendChild(hitLine);

        // Visible underline stroke
        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", `${pBottomLeft.x}`);
        line.setAttribute("y1", `${pBottomLeft.y}`);
        line.setAttribute("x2", `${pBottomRight.x}`);
        line.setAttribute("y2", `${pBottomRight.y}`);
        line.setAttribute("stroke", colorHex);
        line.setAttribute("stroke-width", "2.5");
        line.setAttribute("stroke-linecap", "round");
        g.appendChild(line);
      }
    } else {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      line.setAttribute("x1", `${screenRect.left}`);
      line.setAttribute("y1", `${screenRect.top + screenRect.height}`);
      line.setAttribute("x2", `${screenRect.left + screenRect.width}`);
      line.setAttribute("y2", `${screenRect.top + screenRect.height}`);
      line.setAttribute("stroke", colorHex);
      line.setAttribute("stroke-width", "2.5");
      line.setAttribute("stroke-linecap", "round");
      g.appendChild(line);
    }

    // Subtle Zotero Bookmark indicator if this underline has an attached note or comment
    const hasNote = Boolean(
      (annot.contents && annot.contents.trim() !== "" && annot.contents.trim() !== annot.highlightedText?.trim()) ||
      annot.linkedNote
    );
    if (hasNote) {
      let bookmarkX = 0;
      let bookmarkY = 0;
      if (annot.quadPoints && annot.quadPoints.length >= 8) {
        const p1 = this.coordinateTransformer.pdfToScreen(annot.quadPoints[0], annot.quadPoints[1], this.scale);
        bookmarkX = p1.x - 3;
        bookmarkY = p1.y - 13;
      } else {
        const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
        bookmarkX = screenRect.left - 3;
        bookmarkY = screenRect.top - 13;
      }
      const bookmarkG = this.createBookmarkIndicator(bookmarkX, bookmarkY, colorHex, annot);
      g.appendChild(bookmarkG);
    }

    if (isSelected) {
      const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
      const selBox = this.createSelectionBox(screenRect.left, screenRect.top, screenRect.width, screenRect.height);
      g.appendChild(selBox);
    }

    this.svgOverlayEl.appendChild(g);
  }

  private createBookmarkIndicator(x: number, y: number, colorHex: string, annot: PdfAnnotation): SVGGElement {
    const bg = document.createElementNS("http://www.w3.org/2000/svg", "g");
    bg.setAttribute("class", "pdf-annot-bookmark-badge");
    bg.setAttribute("transform", `translate(${Math.round(x)}, ${Math.round(y)})`);
    bg.setAttribute("cursor", "pointer");
    bg.setAttribute("pointer-events", "auto");

    const notePreview =
      annot.contents && annot.contents !== (annot as any).highlightedText
        ? annot.contents
        : "View Note";
    bg.setAttribute("title", `Note: ${notePreview}`);

    // Sleek Zotero bookmark ribbon shape
    const ribbon = document.createElementNS("http://www.w3.org/2000/svg", "path");
    ribbon.setAttribute("d", "M 0,0 L 12,0 C 13,0 13.5,0.5 13.5,1.5 L 13.5,15 L 6.75,11 L 0,15 L 0,1.5 C 0,0.5 0.5,0 0,0 Z");
    ribbon.setAttribute("fill", colorHex);
    ribbon.setAttribute("stroke", "rgba(0, 0, 0, 0.2)");
    ribbon.setAttribute("stroke-width", "0.8");
    ribbon.setAttribute("filter", "drop-shadow(0 2px 4px rgba(0,0,0,0.35))");
    bg.appendChild(ribbon);

    // Subtle inner accent line
    const accentLine = document.createElementNS("http://www.w3.org/2000/svg", "path");
    accentLine.setAttribute("d", "M 3.5,3.5 L 10,3.5 M 3.5,6 L 8.5,6");
    accentLine.setAttribute("stroke", "#ffffff");
    accentLine.setAttribute("stroke-width", "1.2");
    accentLine.setAttribute("stroke-linecap", "round");
    accentLine.setAttribute("opacity", "0.85");
    bg.appendChild(accentLine);

    bg.addEventListener("pointerdown", (e: PointerEvent) => {
      e.stopPropagation();
      e.preventDefault();
      this.selectedAnnotId = annot.id;
      this.delegate?.onClickAnnotation(annot, bg);
    });

    return bg;
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

    const colors =
      this.colorScheme?.colors && this.colorScheme.colors.length > 0
        ? this.colorScheme.colors
        : PDF_COLOR_PRESETS.map((p) => ({
            id: p.name.toLowerCase(),
            hex: p.hex,
            name: p.name,
            tag: "",
            rgb: p.rgb,
          }));

    const existingBox = this.htmlOverlayEl.querySelector(`.pdf-annot-freetext-box[data-annot-id="${annot.id}"]`) as HTMLElement | null;

    if (existingBox) {
      // If actively being transformed (dragged or resized), avoid overriding live transform
      if (existingBox.dataset.isTransforming !== "true") {
        existingBox.style.left = `${screenRect.left}px`;
        existingBox.style.top = `${screenRect.top}px`;
        existingBox.style.width = `${Math.max(60, screenRect.width)}px`;
        existingBox.style.height = `${Math.max(24, screenRect.height)}px`;
        if (screenRect.top < 38) {
          existingBox.addClass("is-top-edge");
        } else {
          existingBox.removeClass("is-top-edge");
        }
      }

      if (isSelected) {
        existingBox.addClass("is-selected");
      } else {
        existingBox.removeClass("is-selected");
      }

      const contentEl = existingBox.querySelector(".pdf-annot-freetext-content") as HTMLElement | null;
      if (contentEl) {
        if (document.activeElement !== contentEl) {
          contentEl.textContent = annot.contents || "";
        }
        contentEl.style.fontSize = `${Math.max(8, (annot.fontSize || 13) * this.scale)}px`;
        contentEl.style.color = rgbToHex(annot.textColor || [0.1, 0.1, 0.1]);
        contentEl.style.fontWeight = annot.isBold ? "bold" : "normal";
        contentEl.style.fontStyle = annot.isItalic ? "italic" : "normal";
        contentEl.style.textAlign = annot.textAlign || "left";
      }

      // Update background and border
      existingBox.style.backgroundColor = annot.fillColor ? rgbToHex(annot.fillColor) : "transparent";
      if (annot.borderColor && annot.borderStyle && annot.borderStyle !== "none") {
        existingBox.style.border = `${annot.borderWidth || 1}px ${annot.borderStyle} ${rgbToHex(annot.borderColor)}`;
      } else {
        existingBox.style.border = isSelected ? "1.5px solid var(--interactive-accent)" : "1px dashed transparent";
      }

      // Update toolbar indicators
      const sizeLabel = existingBox.querySelector(".pdf-annot-freetext-size-label");
      if (sizeLabel) sizeLabel.textContent = `${annot.fontSize || 13}`;

      const boldBtn = existingBox.querySelector(".pdf-annot-freetext-btn-bold");
      if (boldBtn) boldBtn.classList.toggle("is-active", !!annot.isBold);

      const italicBtn = existingBox.querySelector(".pdf-annot-freetext-btn-italic");
      if (italicBtn) italicBtn.classList.toggle("is-active", !!annot.isItalic);

      const alignBtn = existingBox.querySelector(".pdf-annot-freetext-btn-align");
      if (alignBtn) {
        const align = annot.textAlign || "left";
        alignBtn.textContent = align === "center" ? "⇋" : align === "right" ? "⇥" : "⇤";
      }

      const colorDot = existingBox.querySelector(".pdf-annot-freetext-color-dot") as HTMLElement | null;
      if (colorDot) {
        colorDot.style.backgroundColor = rgbToHex(annot.textColor || [0.1, 0.1, 0.1]);
      }

      return;
    }

    // Build fresh FreeText box DOM element
    const box = this.htmlOverlayEl.createDiv({
      cls: `pdf-annot-freetext-box ${isSelected ? "is-selected" : ""}`,
    });
    box.dataset.annotId = annot.id;
    box.style.left = `${screenRect.left}px`;
    box.style.top = `${screenRect.top}px`;
    const boxW = Math.max(60, screenRect.width);
    box.style.width = `${boxW}px`;
    const boxH = Math.max(24, screenRect.height);
    box.style.height = `${boxH}px`;
    if (screenRect.top < 38) {
      box.addClass("is-top-edge");
    }

    // Background & Border styling
    box.style.backgroundColor = annot.fillColor ? rgbToHex(annot.fillColor) : "transparent";
    if (annot.borderColor && annot.borderStyle && annot.borderStyle !== "none") {
      box.style.border = `${annot.borderWidth || 1}px ${annot.borderStyle} ${rgbToHex(annot.borderColor)}`;
    } else {
      box.style.border = isSelected ? "1.5px solid var(--interactive-accent)" : "1px dashed transparent";
    }

    // 1. Floating Action & Format Toolbar
    const handleBar = box.createDiv({ cls: "pdf-annot-freetext-handle" });

    // Drag-to-move Grip
    const grip = handleBar.createSpan({
      cls: "pdf-annot-freetext-grip",
      text: "⋮⋮",
      attr: { title: "Drag to move text box across PDF" },
    });

    const formatWrap = handleBar.createDiv({ cls: "pdf-annot-freetext-format-wrap" });

    // Font size -
    const fontMinus = formatWrap.createSpan({
      cls: "pdf-annot-freetext-btn",
      text: "A-",
      attr: { title: "Decrease font size" },
    });
    fontMinus.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const currentSize = annot.fontSize || 13;
      annot.fontSize = Math.max(8, currentSize - 1);
      contentEl.style.fontSize = `${Math.max(8, annot.fontSize * this.scale)}px`;
      fontSizeLabel.setText(`${annot.fontSize}`);
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Font size label
    const fontSizeLabel = formatWrap.createSpan({
      cls: "pdf-annot-freetext-size-label",
      text: `${annot.fontSize || 13}`,
    });

    // Font size +
    const fontPlus = formatWrap.createSpan({
      cls: "pdf-annot-freetext-btn",
      text: "A+",
      attr: { title: "Increase font size" },
    });
    fontPlus.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const currentSize = annot.fontSize || 13;
      annot.fontSize = Math.min(72, currentSize + 1);
      contentEl.style.fontSize = `${Math.max(8, annot.fontSize * this.scale)}px`;
      fontSizeLabel.setText(`${annot.fontSize}`);
      this.delegate?.onUpdateAnnotation(annot);
    });

    formatWrap.createSpan({ cls: "pdf-annot-freetext-divider" });

    // Bold button
    const boldBtn = formatWrap.createSpan({
      cls: `pdf-annot-freetext-btn pdf-annot-freetext-btn-bold ${annot.isBold ? "is-active" : ""}`,
      text: "B",
      attr: { title: "Bold" },
    });
    boldBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      annot.isBold = !annot.isBold;
      boldBtn.classList.toggle("is-active", annot.isBold);
      contentEl.style.fontWeight = annot.isBold ? "bold" : "normal";
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Italic button
    const italicBtn = formatWrap.createSpan({
      cls: `pdf-annot-freetext-btn pdf-annot-freetext-btn-italic ${annot.isItalic ? "is-active" : ""}`,
      text: "I",
      attr: { title: "Italic" },
    });
    italicBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      annot.isItalic = !annot.isItalic;
      italicBtn.classList.toggle("is-active", annot.isItalic);
      contentEl.style.fontStyle = annot.isItalic ? "italic" : "normal";
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Alignment toggle button
    const alignBtn = formatWrap.createSpan({
      cls: "pdf-annot-freetext-btn pdf-annot-freetext-btn-align",
      text: annot.textAlign === "center" ? "⇋" : annot.textAlign === "right" ? "⇥" : "⇤",
      attr: { title: "Text Alignment (Left / Center / Right)" },
    });
    alignBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const current = annot.textAlign || "left";
      const next = current === "left" ? "center" : current === "center" ? "right" : "left";
      annot.textAlign = next;
      alignBtn.setText(next === "center" ? "⇋" : next === "right" ? "⇥" : "⇤");
      contentEl.style.textAlign = next;
      this.delegate?.onUpdateAnnotation(annot);
    });

    formatWrap.createSpan({ cls: "pdf-annot-freetext-divider" });

    // Color Swatch Dot
    const colorBtn = formatWrap.createSpan({
      cls: "pdf-annot-freetext-color-dot",
      attr: { title: "Change text color" },
    });
    colorBtn.style.backgroundColor = rgbToHex(annot.textColor || [0.1, 0.1, 0.1]);
    colorBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const currentHex = normalizeHexColor(rgbToHex(annot.textColor || [0.1, 0.1, 0.1]));
      const idx = colors.findIndex((c) => normalizeHexColor(c.hex) === currentHex);
      const nextItem = colors[(idx + 1) % colors.length];
      annot.textColor = nextItem.rgb || hexToRgb(nextItem.hex);
      colorBtn.style.backgroundColor = nextItem.hex;
      contentEl.style.color = nextItem.hex;
      colorBtn.title = `Change text color (${nextItem.name}${nextItem.tag ? `: ${nextItem.tag}` : ""})`;
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Background fill toggle button
    const fillBtn = formatWrap.createSpan({
      cls: `pdf-annot-freetext-btn ${annot.fillColor ? "is-active" : ""}`,
      text: "▨",
      attr: { title: "Toggle background fill (transparent / tinted / white)" },
    });
    fillBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (!annot.fillColor) {
        annot.fillColor = [1.0, 0.98, 0.82]; // Subtle warm yellow tint
      } else if (annot.fillColor[0] === 1.0 && annot.fillColor[1] === 0.98) {
        annot.fillColor = [1.0, 1.0, 1.0]; // Solid white
      } else {
        annot.fillColor = undefined; // Transparent
      }
      fillBtn.classList.toggle("is-active", !!annot.fillColor);
      box.style.backgroundColor = annot.fillColor ? rgbToHex(annot.fillColor) : "transparent";
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Border toggle button
    const borderBtn = formatWrap.createSpan({
      cls: `pdf-annot-freetext-btn ${annot.borderStyle && annot.borderStyle !== "none" ? "is-active" : ""}`,
      text: "▢",
      attr: { title: "Toggle border (none / solid / dashed)" },
    });
    borderBtn.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (!annot.borderStyle || annot.borderStyle === "none") {
        annot.borderStyle = "solid";
        annot.borderWidth = 1;
        annot.borderColor = annot.textColor || [0.2, 0.2, 0.2];
      } else if (annot.borderStyle === "solid") {
        annot.borderStyle = "dashed";
        annot.borderWidth = 1;
        annot.borderColor = annot.textColor || [0.2, 0.2, 0.2];
      } else {
        annot.borderStyle = "none";
        annot.borderColor = undefined;
      }
      borderBtn.classList.toggle("is-active", annot.borderStyle !== "none");
      if (annot.borderColor && annot.borderStyle !== "none") {
        box.style.border = `${annot.borderWidth || 1}px ${annot.borderStyle} ${rgbToHex(annot.borderColor)}`;
      } else {
        box.style.border = isSelected ? "1.5px solid var(--interactive-accent)" : "1px dashed transparent";
      }
      this.delegate?.onUpdateAnnotation(annot);
    });

    // Delete button
    const deleteBtn = handleBar.createSpan({
      cls: "pdf-annot-freetext-del",
      text: "×",
      attr: { title: "Delete text box (Del)" },
    });
    deleteBtn.addEventListener("pointerdown", (e: PointerEvent) => {
      e.stopPropagation();
      e.preventDefault();
      this.delegate?.onDeleteAnnotation(annot.id);
      if (this.selectedAnnotId === annot.id) {
        this.selectedAnnotId = null;
      }
      box.remove();
    });

    // 2. Content Editable Element
    const contentEl = box.createDiv({
      cls: "pdf-annot-freetext-content",
      attr: {
        spellcheck: "false",
        placeholder: "Type text here...",
      },
    });
    contentEl.contentEditable = "false";
    contentEl.textContent = annot.contents || "";
    contentEl.style.fontSize = `${Math.max(8, (annot.fontSize || 13) * this.scale)}px`;
    contentEl.style.color = rgbToHex(annot.textColor || [0.1, 0.1, 0.1]);
    contentEl.style.fontWeight = annot.isBold ? "bold" : "normal";
    contentEl.style.fontStyle = annot.isItalic ? "italic" : "normal";
    contentEl.style.textAlign = annot.textAlign || "left";

    const enterEditMode = () => {
      box.addClass("is-editing");
      contentEl.contentEditable = "plaintext-only";
      contentEl.focus();
    };

    contentEl.addEventListener("input", () => {
      annot.contents = contentEl.textContent || "";
      this.delegate?.onUpdateAnnotation(annot);
    });

    contentEl.addEventListener("blur", () => {
      box.removeClass("is-editing");
      contentEl.contentEditable = "false";
      const text = contentEl.textContent?.trim() || "";

      // Cleanup empty abandoned boxes
      if (!text && !annot.fillColor && (!annot.borderStyle || annot.borderStyle === "none")) {
        box.remove();
        this.delegate?.onDeleteAnnotation(annot.id);
        if (this.selectedAnnotId === annot.id) {
          this.selectedAnnotId = null;
        }
        return;
      }

      annot.contents = contentEl.textContent || "";
      const boxRect = box.getBoundingClientRect();
      const pageRect = this.containerEl.getBoundingClientRect();
      const screenBounds = {
        left: boxRect.left - pageRect.left,
        top: boxRect.top - pageRect.top,
        width: Math.max(50, boxRect.width),
        height: Math.max(22, boxRect.height),
      };
      if (this.coordinateTransformer) {
        annot.rect = this.coordinateTransformer.screenRectToPdfRect(screenBounds, this.scale);
      }
      this.delegate?.onUpdateAnnotation(annot);
    });

    // 3. Move Across PDF Listener (Grip or Box Drag)
    const setupMove = (triggerEl: HTMLElement) => {
      triggerEl.addEventListener("pointerdown", (e: PointerEvent) => {
        if ((e.target as HTMLElement).closest(".pdf-annot-freetext-btn, .pdf-annot-freetext-del, .pdf-annot-freetext-color-dot, .pdf-annot-resize-handle")) {
          return;
        }

        e.stopPropagation();

        // If clicking inside contentEl and already selected: second click enters editing
        if (triggerEl === contentEl && box.hasClass("is-selected") && !box.hasClass("is-editing")) {
          enterEditMode();
          return;
        }

        if (box.hasClass("is-editing")) {
          return; // Allow text selection / cursor placement inside editing mode
        }

        // Select text box
        if (this.selectedAnnotId !== annot.id) {
          this.selectAnnotation(annot.id);
        }

        e.preventDefault();
        box.dataset.isTransforming = "true";
        document.body.style.cursor = "grabbing";

        const startX = e.clientX;
        const startY = e.clientY;
        const startLeft = parseFloat(box.style.left) || screenRect.left;
        const startTop = parseFloat(box.style.top) || screenRect.top;
        let hasMoved = false;

        const onMove = (me: PointerEvent) => {
          const dx = me.clientX - startX;
          const dy = me.clientY - startY;
          if (!hasMoved && Math.hypot(dx, dy) > 2) {
            hasMoved = true;
          }
          const curLeft = startLeft + dx;
          const curTop = startTop + dy;
          box.style.left = `${curLeft}px`;
          box.style.top = `${curTop}px`;
          if (curTop < 38) {
            box.addClass("is-top-edge");
          } else {
            box.removeClass("is-top-edge");
          }
        };

        const onUp = (ue: PointerEvent) => {
          window.removeEventListener("pointermove", onMove);
          window.removeEventListener("pointerup", onUp);
          delete box.dataset.isTransforming;
          document.body.style.cursor = "";

          if (!hasMoved) {
            return;
          }

          // Check if dropped across to another page
          const dropTarget = document.elementFromPoint(ue.clientX, ue.clientY);
          const targetPageEl = dropTarget?.closest(".pdf-page-view") as HTMLElement | null;
          const targetIdxStr = targetPageEl?.dataset.pageIndex;
          const targetPageIndex = targetIdxStr !== undefined ? parseInt(targetIdxStr, 10) : NaN;

          if (!isNaN(targetPageIndex) && targetPageIndex !== this.pageIndex) {
            // Transfer annotation to target page!
            const targetPageView = this.delegate?.getPageView?.(targetPageIndex);
            const targetTransformer = targetPageView?.getCoordinateTransformer();
            if (targetTransformer && targetPageEl) {
              const targetRect = targetPageEl.getBoundingClientRect();
              const boxRect = box.getBoundingClientRect();
              const screenBoundsOnTarget = {
                left: boxRect.left - targetRect.left,
                top: boxRect.top - targetRect.top,
                width: box.offsetWidth,
                height: box.offsetHeight,
              };
              annot.pageIndex = targetPageIndex;
              annot.rect = targetTransformer.screenRectToPdfRect(screenBoundsOnTarget, this.scale);
              box.remove();
              this.delegate?.onUpdateAnnotation(annot);
              return;
            }
          }

          // Dropped on current page
          const boxRect = box.getBoundingClientRect();
          const pageRect = this.containerEl.getBoundingClientRect();
          const screenBounds = {
            left: boxRect.left - pageRect.left,
            top: boxRect.top - pageRect.top,
            width: boxRect.width,
            height: boxRect.height,
          };
          if (this.coordinateTransformer) {
            annot.rect = this.coordinateTransformer.screenRectToPdfRect(screenBounds, this.scale);
          }
          this.delegate?.onUpdateAnnotation(annot);
        };

        window.addEventListener("pointermove", onMove);
        window.addEventListener("pointerup", onUp);
      });
    };

    setupMove(grip);
    setupMove(box);
    setupMove(contentEl);

    // Double-click to edit
    box.addEventListener("dblclick", (e) => {
      e.stopPropagation();
      enterEditMode();
    });

    // 4. 8-Handle Transform Tools (NW, N, NE, E, SE, S, SW, W)
    const handles: { dir: "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w"; cls: string; title: string }[] = [
      { dir: "nw", cls: "pdf-annot-resize-nw", title: "Resize" },
      { dir: "n",  cls: "pdf-annot-resize-n",  title: "Resize height" },
      { dir: "ne", cls: "pdf-annot-resize-ne", title: "Resize" },
      { dir: "e",  cls: "pdf-annot-resize-e",  title: "Resize width" },
      { dir: "se", cls: "pdf-annot-resize-se", title: "Resize" },
      { dir: "s",  cls: "pdf-annot-resize-s",  title: "Resize height" },
      { dir: "sw", cls: "pdf-annot-resize-sw", title: "Resize" },
      { dir: "w",  cls: "pdf-annot-resize-w",  title: "Resize width" },
    ];

    const setupResize = (targetHandle: HTMLElement, dir: "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w") => {
      targetHandle.addEventListener("pointerdown", (e: PointerEvent) => {
        e.stopPropagation();
        e.preventDefault();
        box.dataset.isTransforming = "true";

        const startX = e.clientX;
        const startY = e.clientY;
        const startLeft = parseFloat(box.style.left) || screenRect.left;
        const startTop = parseFloat(box.style.top) || screenRect.top;
        const startW = box.offsetWidth;
        const startH = box.offsetHeight;

        const minW = 50;
        const minH = 22;

        let lastBounds = { left: startLeft, top: startTop, width: startW, height: startH };

        const onMove = (me: PointerEvent) => {
          const dx = me.clientX - startX;
          const dy = me.clientY - startY;

          let newLeft = startLeft;
          let newTop = startTop;
          let newW = startW;
          let newH = startH;

          // Horizontal calculation
          if (dir === "e" || dir === "ne" || dir === "se") {
            newW = Math.max(minW, startW + dx);
          } else if (dir === "w" || dir === "nw" || dir === "sw") {
            newW = Math.max(minW, startW - dx);
            newLeft = startLeft + (startW - newW);
          }

          // Vertical calculation
          if (dir === "s" || dir === "se" || dir === "sw") {
            newH = Math.max(minH, startH + dy);
          } else if (dir === "n" || dir === "nw" || dir === "ne") {
            newH = Math.max(minH, startH - dy);
            newTop = startTop + (startH - newH);
          }

          box.style.left = `${newLeft}px`;
          box.style.top = `${newTop}px`;
          box.style.width = `${newW}px`;
          box.style.height = `${newH}px`;

          if (newTop < 38) {
            box.addClass("is-top-edge");
          } else {
            box.removeClass("is-top-edge");
          }

          lastBounds = { left: newLeft, top: newTop, width: newW, height: newH };
        };

        const onUp = () => {
          window.removeEventListener("pointermove", onMove);
          window.removeEventListener("pointerup", onUp);
          delete box.dataset.isTransforming;

          if (this.coordinateTransformer) {
            annot.rect = this.coordinateTransformer.screenRectToPdfRect(lastBounds, this.scale);
          }
          this.delegate?.onUpdateAnnotation(annot);
        };

        window.addEventListener("pointermove", onMove);
        window.addEventListener("pointerup", onUp);
      });
    };

    for (const h of handles) {
      const handleEl = box.createDiv({
        cls: `pdf-annot-resize-handle ${h.cls}`,
        attr: { title: h.title },
      });
      setupResize(handleEl, h.dir);
    }
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
      if (annot.linePoints && annot.linePoints.length >= 4 && this.coordinateTransformer) {
        const p1 = this.coordinateTransformer.pdfToScreen(annot.linePoints[0], annot.linePoints[1], this.scale);
        const p2 = this.coordinateTransformer.pdfToScreen(annot.linePoints[2], annot.linePoints[3], this.scale);
        shapeEl.setAttribute("x1", `${p1.x}`);
        shapeEl.setAttribute("y1", `${p1.y}`);
        shapeEl.setAttribute("x2", `${p2.x}`);
        shapeEl.setAttribute("y2", `${p2.y}`);
      } else {
        shapeEl.setAttribute("x1", `${screenRect.left}`);
        shapeEl.setAttribute("y1", `${screenRect.top}`);
        shapeEl.setAttribute("x2", `${screenRect.left + screenRect.width}`);
        shapeEl.setAttribute("y2", `${screenRect.top + screenRect.height}`);
      }
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

        // If it's a freetext annotation, FreeText handles its own drag and transform
        if (annot.type === "freetext") {
          this.renderAnnotations();
          return;
        }

        // Prepare drag state
        this.dragState = {
          annot,
          startScreen: { x: sx, y: sy },
          startPdfRect: [...annot.rect] as PdfRect,
          startPaths: annot.type === "ink" ? JSON.parse(JSON.stringify((annot as PdfInkAnnotation).paths)) : undefined,
          startQuadPoints: (annot.type === "highlight" || annot.type === "underline") ? [...(annot as any).quadPoints] : undefined,
          startLinePoints: (annot as any).linePoints ? ([...(annot as any).linePoints] as [number, number, number, number]) : undefined,
          hasMoved: false,
        };

        this.renderAnnotations();
        return;
      }
    }

    // No existing annotation clicked: user clicked on empty space
    if (this.selectedAnnotId && this.currentTool !== "draw" && this.currentTool !== "square" && this.currentTool !== "circle" && this.currentTool !== "line") {
      this.selectedAnnotId = null;
      this.renderAnnotations();
    }

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
      // Start drag-to-create or click-to-place text box
      this.isCreatingFreeText = true;
      this.freeTextStartPoint = { x: sx, y: sy };
      this.liveFreeTextEl = this.htmlOverlayEl.createDiv({ cls: "pdf-annot-freetext-preview" });
      this.liveFreeTextEl.style.left = `${sx}px`;
      this.liveFreeTextEl.style.top = `${sy}px`;
      this.liveFreeTextEl.style.width = "0px";
      this.liveFreeTextEl.style.height = "0px";
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

    // 0. Live FreeText drag-to-create rectangle
    if (this.isCreatingFreeText && this.freeTextStartPoint && this.liveFreeTextEl) {
      const left = Math.min(this.freeTextStartPoint.x, sx);
      const top = Math.min(this.freeTextStartPoint.y, sy);
      const width = Math.abs(sx - this.freeTextStartPoint.x);
      const height = Math.abs(sy - this.freeTextStartPoint.y);
      this.liveFreeTextEl.style.left = `${left}px`;
      this.liveFreeTextEl.style.top = `${top}px`;
      this.liveFreeTextEl.style.width = `${width}px`;
      this.liveFreeTextEl.style.height = `${height}px`;
      return;
    }

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

        // Update highlight / underline quadpoints
        if ((annot.type === "highlight" || annot.type === "underline") && this.dragState.startQuadPoints) {
          (annot as any).quadPoints = this.dragState.startQuadPoints.map((val, idx) =>
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
        if (annot.type === "text" || annot.type === "highlight" || annot.type === "underline") {
          const annotEl = this.containerEl.querySelector(`[data-annot-id="${annot.id}"]`) as HTMLElement | SVGElement | null;
          this.delegate?.onClickAnnotation(annot, annotEl || this.containerEl);
        }
        return;
      }
    }

    // 2. FreeText creation end
    if (this.isCreatingFreeText && this.freeTextStartPoint && this.coordinateTransformer) {
      this.isCreatingFreeText = false;
      this.liveFreeTextEl?.remove();
      this.liveFreeTextEl = null;

      const rect = this.containerEl.getBoundingClientRect();
      const sx = e.clientX - rect.left;
      const sy = e.clientY - rect.top;

      const dx = Math.abs(sx - this.freeTextStartPoint.x);
      const dy = Math.abs(sy - this.freeTextStartPoint.y);

      let boxLeft: number, boxTop: number, boxW: number, boxH: number;
      if (dx > 15 || dy > 15) {
        boxLeft = Math.min(this.freeTextStartPoint.x, sx);
        boxTop = Math.min(this.freeTextStartPoint.y, sy);
        boxW = Math.max(80, dx);
        boxH = Math.max(28, dy);
      } else {
        boxLeft = this.freeTextStartPoint.x;
        boxTop = this.freeTextStartPoint.y;
        boxW = 160;
        boxH = 36;
      }

      const pdfRect = this.coordinateTransformer.screenRectToPdfRect(
        { left: boxLeft, top: boxTop, width: boxW, height: boxH },
        this.scale
      );

      const annot: PdfFreeTextAnnotation = {
        id: `freetext_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        type: "freetext",
        pageIndex: this.pageIndex,
        rect: pdfRect,
        color: this.currentColor,
        textColor: this.currentColor,
        fontSize: 13,
        contents: "",
        author: "User",
        createdAt: new Date(),
      };

      this.selectedAnnotId = annot.id;
      this.delegate?.onAddAnnotation(annot);

      window.setTimeout(() => {
        const box = this.htmlOverlayEl.querySelector(`.pdf-annot-freetext-box[data-annot-id="${annot.id}"]`) as HTMLElement | null;
        if (box) {
          box.addClass("is-editing");
          const contentEl = box.querySelector(".pdf-annot-freetext-content") as HTMLElement | null;
          if (contentEl) {
            contentEl.contentEditable = "plaintext-only";
            contentEl.focus();
          }
        }
      }, 40);
      return;
    }

    // 3. Drawing or shape creation end
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

    let linePoints: [number, number, number, number] | undefined;
    if (shapeType === "line") {
      const pStart = this.coordinateTransformer.screenToPdf(x1, y1, this.scale);
      const pEnd = this.coordinateTransformer.screenToPdf(x2, y2, this.scale);
      linePoints = [pStart.x, pStart.y, pEnd.x, pEnd.y];
    }

    const annot: PdfShapeAnnotation = {
      id: `shape_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: shapeType,
      pageIndex: this.pageIndex,
      rect: pdfRect,
      color: this.currentColor,
      strokeColor: this.currentColor,
      strokeWidth: this.strokeWidth,
      linePoints,
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
    window.setTimeout(() => {
      const pinEl = this.containerEl.querySelector(`[data-annot-id="${annot.id}"]`) as HTMLElement | SVGElement | null;
      this.delegate?.onClickAnnotation(annot, pinEl || this.containerEl);
    }, 60);
  }

  private placeFreeTextBox(sx: number, sy: number): void {
    if (!this.coordinateTransformer) return;
    const pdfRect = this.coordinateTransformer.screenRectToPdfRect(
      { left: sx, top: sy, width: 160, height: 36 },
      this.scale
    );

    const annot: PdfFreeTextAnnotation = {
      id: `freetext_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "freetext",
      pageIndex: this.pageIndex,
      rect: pdfRect,
      color: this.currentColor,
      textColor: this.currentColor,
      fontSize: 13,
      contents: "",
      author: "User",
      createdAt: new Date(),
    };

    this.selectedAnnotId = annot.id;
    this.delegate?.onAddAnnotation(annot);

    // Auto-focus immediately so user can type right away!
    window.setTimeout(() => {
      const box = this.htmlOverlayEl.querySelector(`.pdf-annot-freetext-box[data-annot-id="${annot.id}"]`) as HTMLElement | null;
      if (box) {
        box.addClass("is-editing");
        const contentEl = box.querySelector(".pdf-annot-freetext-content") as HTMLElement | null;
        if (contentEl) {
          contentEl.contentEditable = "plaintext-only";
          contentEl.focus();
        }
      }
    }, 40);
  }

  // ==================== Text Selection Support ====================

  private setupTextSelection(): void {
    const handleMouseUp = () => {
      const sel = window.getSelection();
      if (!sel || sel.isCollapsed || !sel.rangeCount) return;

      const text = sel.toString().trim();
      if (!text) return;

      // In Highlight tool mode: instantly create highlight!
      if (this.currentTool === "highlight") {
        this.createHighlightFromSelection(text, this.currentColor);
        return;
      }

      // In Underline tool mode: instantly create underline!
      if (this.currentTool === "underline") {
        this.createUnderlineFromSelection(text, this.currentColor);
        return;
      }

      // In Hand tool mode: open sleek Zotero selection bar!
      if (this.currentTool === "hand") {
        const range = sel.getRangeAt(0);
        const rect = range.getBoundingClientRect();
        this.delegate?.onTextSelected?.(text, rect, this.pageIndex);
      }
    };

    this.textLayerEl.addEventListener("mouseup", handleMouseUp);
  }

  private setupKeyboardShortcuts(): void {
    this.onKeyDownListener = (e: KeyboardEvent) => {
      // 1. Escape key: if any text box is in editing mode, blur it to exit editing mode
      if (e.key === "Escape") {
        const editingBox = this.htmlOverlayEl.querySelector(".pdf-annot-freetext-box.is-editing") as HTMLElement | null;
        if (editingBox) {
          const contentEl = editingBox.querySelector(".pdf-annot-freetext-content") as HTMLElement | null;
          contentEl?.blur();
          return;
        }
      }

      // 2. Arrow keys nudge selected text box (when not editing text)
      if (["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key) && this.selectedAnnotId) {
        const activeTag = document.activeElement?.tagName.toLowerCase();
        const isContentEditable = (document.activeElement as HTMLElement)?.isContentEditable;
        if (activeTag !== "input" && activeTag !== "textarea" && !isContentEditable) {
          const annot = this.annotations.find((a) => a.id === this.selectedAnnotId);
          if (annot && annot.type === "freetext" && this.coordinateTransformer) {
            e.preventDefault();
            const delta = e.shiftKey ? 10 : 2;
            const dx = e.key === "ArrowLeft" ? -delta : e.key === "ArrowRight" ? delta : 0;
            const dy = e.key === "ArrowUp" ? -delta : e.key === "ArrowDown" ? delta : 0;

            const screenRect = this.coordinateTransformer.pdfRectToScreenRect(annot.rect, this.scale);
            const newBounds = {
              left: screenRect.left + dx,
              top: screenRect.top + dy,
              width: screenRect.width,
              height: screenRect.height,
            };
            annot.rect = this.coordinateTransformer.screenRectToPdfRect(newBounds, this.scale);
            this.delegate?.onUpdateAnnotation(annot);
            this.renderAnnotations();
            return;
          }
        }
      }

      // 3. Delete or Backspace
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

  /**
   * Groups selection client rects by line and merges overlapping/adjacent segments
   * into clean unified line rectangles. Prevents blocky artifacts and guarantees
   * single-word selection accuracy!
   */
  private getMergedLineRectsFromSelection(): { left: number; top: number; right: number; bottom: number }[] {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return [];

    const range = sel.getRangeAt(0);
    const rawRects = Array.from(range.getClientRects());
    if (rawRects.length === 0) return [];

    const pageRect = this.containerEl.getBoundingClientRect();

    // 1. Convert to page-relative coordinates & filter out zero-width/height rects
    const localRects = rawRects
      .filter((r) => r.width > 0.5 && r.height > 0.5)
      .map((r) => ({
        left: r.left - pageRect.left,
        top: r.top - pageRect.top,
        right: r.right - pageRect.left,
        bottom: r.bottom - pageRect.top,
        height: r.height,
      }));

    if (localRects.length === 0) return [];

    // 2. Merge fragments on the same line
    const lines: { left: number; top: number; right: number; bottom: number }[] = [];

    for (const r of localRects) {
      const existingLine = lines.find((l) => {
        const vOverlap = Math.min(l.bottom, r.bottom) - Math.max(l.top, r.top);
        const minH = Math.min(l.bottom - l.top, r.height);
        return vOverlap > minH * 0.4 || Math.abs(l.top - r.top) < 6;
      });

      if (existingLine) {
        existingLine.left = Math.min(existingLine.left, r.left);
        existingLine.right = Math.max(existingLine.right, r.right);
        existingLine.top = Math.min(existingLine.top, r.top);
        existingLine.bottom = Math.max(existingLine.bottom, r.bottom);
      } else {
        lines.push({ left: r.left, top: r.top, right: r.right, bottom: r.bottom });
      }
    }

    // 3. Sort lines top-to-bottom
    lines.sort((a, b) => a.top - b.top);
    return lines;
  }

  createHighlightFromSelection(selectedText: string, color: PdfRgbColor): void {
    if (!this.coordinateTransformer) return;
    const lines = this.getMergedLineRectsFromSelection();
    if (lines.length === 0) return;

    const quadPoints: number[] = [];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    for (const line of lines) {
      const pTopLeft = this.coordinateTransformer.screenToPdf(line.left, line.top, this.scale);
      const pTopRight = this.coordinateTransformer.screenToPdf(line.right, line.top, this.scale);
      const pBottomLeft = this.coordinateTransformer.screenToPdf(line.left, line.bottom, this.scale);
      const pBottomRight = this.coordinateTransformer.screenToPdf(line.right, line.bottom, this.scale);

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

    const sel = window.getSelection();
    sel?.removeAllRanges();
    this.delegate?.onAddAnnotation(annot);
  }

  createUnderlineFromSelection(selectedText: string, color: PdfRgbColor): void {
    if (!this.coordinateTransformer) return;
    const lines = this.getMergedLineRectsFromSelection();
    if (lines.length === 0) return;

    const quadPoints: number[] = [];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    for (const line of lines) {
      const pTopLeft = this.coordinateTransformer.screenToPdf(line.left, line.top, this.scale);
      const pTopRight = this.coordinateTransformer.screenToPdf(line.right, line.top, this.scale);
      const pBottomLeft = this.coordinateTransformer.screenToPdf(line.left, line.bottom, this.scale);
      const pBottomRight = this.coordinateTransformer.screenToPdf(line.right, line.bottom, this.scale);

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

    const annot: PdfUnderlineAnnotation = {
      id: `underline_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      type: "underline",
      pageIndex: this.pageIndex,
      rect: [minX, minY - 2, maxX, maxY],
      quadPoints,
      color,
      contents: selectedText,
      highlightedText: selectedText,
      author: "User",
      createdAt: new Date(),
    };

    const sel = window.getSelection();
    sel?.removeAllRanges();
    this.delegate?.onAddAnnotation(annot);
  }

  /**
   * Creates a text-attached annotated note directly on the selected sentence or phrase.
   * Immediately opens the note popover so the user can type their note!
   */
  addAnnotatedNoteFromSelection(selectedText: string, color: PdfRgbColor): void {
    if (!this.coordinateTransformer) return;
    const lines = this.getMergedLineRectsFromSelection();
    if (lines.length === 0) return;

    const quadPoints: number[] = [];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    for (const line of lines) {
      const pTopLeft = this.coordinateTransformer.screenToPdf(line.left, line.top, this.scale);
      const pTopRight = this.coordinateTransformer.screenToPdf(line.right, line.top, this.scale);
      const pBottomLeft = this.coordinateTransformer.screenToPdf(line.left, line.bottom, this.scale);
      const pBottomRight = this.coordinateTransformer.screenToPdf(line.right, line.bottom, this.scale);

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

    const sel = window.getSelection();
    sel?.removeAllRanges();
    this.delegate?.onAddAnnotation(annot);

    // Open note popover anchored to the newly created highlight
    window.setTimeout(() => {
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
      window.setTimeout(() => el.classList.remove("is-pulsing"), 1500);
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
