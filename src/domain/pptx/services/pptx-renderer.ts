import {
  PptxPresentation,
  PptxSlide,
  PptxElement,
  PptxParagraph,
  PptxRun,
  PptxTable,
} from "../models/slide";

export interface PptxRenderOptions {
  showSlideNumberBadge?: boolean;
  totalSlides?: number;
  scale?: number; // Zoom/scaling multiplier (default 1)
}

export class PptxRenderer {
  /**
   * Renders a single slide into a self-contained HTMLElement canvas.
   */
  static renderSlide(
    slide: PptxSlide,
    presentation: PptxPresentation,
    options: PptxRenderOptions = {}
  ): HTMLElement {
    const totalSlides = options.totalSlides ?? presentation.slides.length;
    const showBadge = options.showSlideNumberBadge ?? true;
    const width = presentation.width || 960;
    const height = presentation.height || 540;

    const slideEl = document.createElement("div");
    slideEl.className = "office-pptx-slide";
    slideEl.dataset.slideIndex = String(slide.index);
    slideEl.dataset.slideNumber = String(slide.slideNumber);
    slideEl.style.width = `${width}px`;
    slideEl.style.height = `${height}px`;
    slideEl.style.position = "relative";
    slideEl.style.overflow = "hidden";
    slideEl.style.boxSizing = "border-box";

    // Slide background
    if (slide.background) {
      if (slide.background.startsWith("#") || slide.background.startsWith("rgb")) {
        slideEl.style.backgroundColor = slide.background;
      } else if (slide.background.startsWith("linear-gradient") || slide.background.startsWith("radial-gradient")) {
        slideEl.style.backgroundImage = slide.background;
      } else if (slide.background.startsWith("data:") || slide.background.startsWith("blob:") || slide.background.startsWith("http")) {
        slideEl.style.backgroundImage = `url("${slide.background}")`;
        slideEl.style.backgroundSize = "cover";
        slideEl.style.backgroundPosition = "center center";
      } else {
        slideEl.style.backgroundColor = slide.background;
      }
    } else {
      slideEl.style.backgroundColor = "#ffffff";
    }

    // Sort elements by zIndex to preserve stacking order
    const sortedElements = [...slide.elements].sort((a, b) => a.zIndex - b.zIndex);

    for (const el of sortedElements) {
      const elNode = this.renderElement(el, slide, presentation);
      if (elNode) {
        slideEl.appendChild(elNode);
      }
    }

    // Prominent Slide Number Badge in lower-right or overlay
    if (showBadge && totalSlides > 0) {
      const badge = document.createElement("div");
      badge.className = "office-pptx-slide-badge";
      badge.textContent = `${slide.slideNumber} / ${totalSlides}`;
      slideEl.appendChild(badge);
    }

    return slideEl;
  }

  /**
   * Renders a PPTX element (shape, text box, image, table, group, line).
   */
  static renderElement(
    el: PptxElement,
    slide: PptxSlide,
    presentation: PptxPresentation
  ): HTMLElement | null {
    const container = document.createElement("div");
    container.className = `office-pptx-element office-pptx-${el.type}`;
    container.dataset.elementId = el.id;
    container.dataset.elementName = el.name;

    // Positioning
    container.style.position = "absolute";
    container.style.left = `${el.x}px`;
    container.style.top = `${el.y}px`;
    container.style.width = `${el.width}px`;
    container.style.height = `${el.height}px`;
    container.style.boxSizing = "border-box";
    container.style.zIndex = String(el.zIndex);

    if (el.rotation) {
      container.style.transform = `rotate(${el.rotation}deg)`;
      container.style.transformOrigin = "center center";
    }

    if (el.opacity !== undefined && el.opacity < 1) {
      container.style.opacity = String(el.opacity);
    }

    // 1. Image Element
    if (el.type === "image" && el.imageSrc) {
      const img = document.createElement("img");
      img.src = el.imageSrc;
      img.alt = el.imageAlt || el.name || "Slide Image";
      img.style.width = "100%";
      img.style.height = "100%";
      img.style.objectFit = "fill";
      img.style.display = "block";
      img.style.pointerEvents = "none";
      container.appendChild(img);
      return container;
    }

    // 2. Table Element
    if (el.type === "table" && el.table) {
      const tableNode = this.renderTable(el.table, slide);
      container.appendChild(tableNode);
      return container;
    }

    // 3. Group Element
    if (el.type === "group" && el.children) {
      for (const child of el.children) {
        const childNode = this.renderElement(child, slide, presentation);
        if (childNode) {
          container.appendChild(childNode);
        }
      }
      return container;
    }

    // 4. Shape / Textbox Element
    this.applyShapeStyles(container, el);

    // Padding & vertical alignment
    if (el.padding) {
      container.style.paddingTop = `${el.padding.top}px`;
      container.style.paddingRight = `${el.padding.right}px`;
      container.style.paddingBottom = `${el.padding.bottom}px`;
      container.style.paddingLeft = `${el.padding.left}px`;
    }

    if (el.verticalAlign === "middle") {
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.justifyContent = "center";
    } else if (el.verticalAlign === "bottom") {
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.justifyContent = "flex-end";
    }

    // Text Paragraphs
    if (el.paragraphs && el.paragraphs.length > 0) {
      const textWrapper = document.createElement("div");
      textWrapper.className = "office-pptx-text-content";
      textWrapper.style.width = "100%";

      for (const p of el.paragraphs) {
        const pNode = this.renderParagraph(p, slide);
        textWrapper.appendChild(pNode);
      }
      container.appendChild(textWrapper);
    }

    return container;
  }

  /**
   * Applies fills, borders, and shape geometry styling (rounded rect, ellipse, polygon clip-paths).
   */
  private static applyShapeStyles(container: HTMLElement, el: PptxElement): void {
    // Fill
    if (el.fill) {
      if (el.fill.type === "solid" && el.fill.color) {
        container.style.backgroundColor = el.fill.color;
      } else if (el.fill.type === "gradient" && el.fill.gradientStops && el.fill.gradientStops.length > 0) {
        const stops = el.fill.gradientStops
          .map((s) => `${s.color} ${Math.round(s.position * 100)}%`)
          .join(", ");
        container.style.backgroundImage = `linear-gradient(180deg, ${stops})`;
      }
    }

    // Border
    if (el.border && el.border.width > 0) {
      const style = el.border.style || "solid";
      container.style.border = `${el.border.width}px ${style} ${el.border.color}`;
    }

    // Shape geometry
    const geom = (el.shapeType || "rect").toLowerCase();
    switch (geom) {
      case "roundrect":
        container.style.borderRadius = "8px";
        break;
      case "ellipse":
        container.style.borderRadius = "50%";
        break;
      case "triangle":
        container.style.clipPath = "polygon(50% 0%, 0% 100%, 100% 100%)";
        break;
      case "diamond":
        container.style.clipPath = "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)";
        break;
      case "rightarrow":
      case "arrow":
        container.style.clipPath = "polygon(0% 25%, 65% 25%, 65% 0%, 100% 50%, 65% 100%, 65% 75%, 0% 75%)";
        break;
      case "line":
        container.style.borderBottom = `${Math.max(1, el.border?.width || 2)}px ${el.border?.style || "solid"} ${el.border?.color || "#000000"}`;
        container.style.borderTop = "none";
        container.style.borderLeft = "none";
        container.style.borderRight = "none";
        break;
      default:
        break;
    }
  }

  /**
   * Renders a single text paragraph with alignments, bullets, and runs.
   */
  private static renderParagraph(p: PptxParagraph, slide: PptxSlide): HTMLElement {
    const pEl = document.createElement("div");
    pEl.className = "office-pptx-p";
    pEl.style.textAlign = p.align || "left";
    pEl.style.margin = "0";
    pEl.style.padding = "0";

    if (p.spaceBefore) {
      pEl.style.marginTop = `${p.spaceBefore}px`;
    }
    if (p.spaceAfter) {
      pEl.style.marginBottom = `${p.spaceAfter}px`;
    }
    if (p.lineSpacing) {
      pEl.style.lineHeight = `${p.lineSpacing}`;
    }
    if (p.indent) {
      pEl.style.paddingLeft = `${p.indent}px`;
    }

    // Bullet styling
    if (p.bullet && p.bullet.type !== "none") {
      pEl.style.display = "flex";
      pEl.style.alignItems = "baseline";

      const bulletSpan = document.createElement("span");
      bulletSpan.className = "office-pptx-bullet";
      bulletSpan.style.marginRight = "6px";
      bulletSpan.style.flexShrink = "0";

      if (p.bullet.type === "char") {
        bulletSpan.textContent = p.bullet.char || "•";
      } else if (p.bullet.type === "number") {
        bulletSpan.textContent = `${p.bullet.number || 1}.`;
      }
      if (p.bullet.color) {
        bulletSpan.style.color = p.bullet.color;
      }
      if (p.bullet.font) {
        bulletSpan.style.fontFamily = p.bullet.font;
      }
      pEl.appendChild(bulletSpan);
    }

    const runsWrapper = document.createElement("span");
    runsWrapper.className = "office-pptx-runs";

    for (const run of p.runs) {
      const runNode = this.renderRun(run, slide);
      runsWrapper.appendChild(runNode);
    }

    // Ensure empty paragraphs still occupy height
    if (p.runs.length === 0 || p.runs.every((r) => !r.text)) {
      runsWrapper.innerHTML = "&nbsp;";
    }

    pEl.appendChild(runsWrapper);
    return pEl;
  }

  /**
   * Renders a text run with font sizes, formatting, colors, and slide number substitution.
   */
  private static renderRun(run: PptxRun, slide: PptxSlide): HTMLElement {
    let span: HTMLElement;

    if (run.hyperlink?.url) {
      const a = document.createElement("a");
      a.href = run.hyperlink.url;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      if (run.hyperlink.tooltip) {
        a.title = run.hyperlink.tooltip;
      }
      span = a;
    } else {
      span = document.createElement("span");
    }

    span.className = "office-pptx-run";

    // Handle slide number placeholder
    if (run.isSlideNumber) {
      span.textContent = String(slide.slideNumber);
      span.classList.add("office-pptx-slide-number-field");
    } else {
      span.textContent = run.text;
    }

    if (run.fontSize) {
      span.style.fontSize = `${run.fontSize}pt`;
    }
    if (run.fontFamily) {
      span.style.fontFamily = run.fontFamily;
    }
    if (run.color) {
      span.style.color = run.color;
    }
    if (run.bold) {
      span.style.fontWeight = "bold";
    }
    if (run.italic) {
      span.style.fontStyle = "italic";
    }
    if (run.underline) {
      span.style.textDecoration = "underline";
    }
    if (run.strikethrough) {
      span.style.textDecoration = span.style.textDecoration
        ? `${span.style.textDecoration} line-through`
        : "line-through";
    }

    return span;
  }

  /**
   * Renders an OOXML table.
   */
  private static renderTable(table: PptxTable, slide: PptxSlide): HTMLElement {
    const tableEl = document.createElement("table");
    tableEl.className = "office-pptx-table";
    tableEl.style.width = "100%";
    tableEl.style.height = "100%";
    tableEl.style.borderCollapse = "collapse";
    tableEl.style.tableLayout = "fixed";

    for (const row of table.rows) {
      const tr = document.createElement("tr");
      if (row.height > 0) {
        tr.style.height = `${row.height}px`;
      }

      for (const cell of row.cells) {
        const td = document.createElement("td");
        td.className = "office-pptx-cell";
        td.style.boxSizing = "border-box";
        td.style.padding = "4px 8px";

        if (cell.backgroundColor) {
          td.style.backgroundColor = cell.backgroundColor;
        }
        if (cell.borderColor && cell.borderWidth) {
          td.style.border = `${cell.borderWidth}px solid ${cell.borderColor}`;
        } else {
          td.style.border = "1px solid rgba(0, 0, 0, 0.15)";
        }

        if (cell.rowSpan && cell.rowSpan > 1) {
          td.rowSpan = cell.rowSpan;
        }
        if (cell.colSpan && cell.colSpan > 1) {
          td.colSpan = cell.colSpan;
        }

        if (cell.verticalAlign === "middle") {
          td.style.verticalAlign = "middle";
        } else if (cell.verticalAlign === "bottom") {
          td.style.verticalAlign = "bottom";
        } else {
          td.style.verticalAlign = "top";
        }

        if (cell.paragraphs && cell.paragraphs.length > 0) {
          for (const p of cell.paragraphs) {
            td.appendChild(this.renderParagraph(p, slide));
          }
        } else if (cell.text) {
          td.textContent = cell.text;
        }

        tr.appendChild(td);
      }
      tableEl.appendChild(tr);
    }

    return tableEl;
  }
}
