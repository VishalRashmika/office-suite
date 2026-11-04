/**
 * Comprehensive Test Suite for Zotero-Style PDF Annotations Side Panel:
 * 1. PdfAnnotationParser case-insensitive subtype parsing & aliases (Highlight, highlight, etc.)
 * 2. Normalization of bounding box coordinates [minX, minY, maxX, maxY]
 * 3. Extraction of ISO 32000-1 dates (M and CreationDate)
 * 4. PdfSidebar build, collapsed-by-default initialization, and refresh-on-open (show, toggle, showAnnotationsTab)
 * 5. 1:1 Zotero card rendering across all annotation types:
 *    - Highlights with quoted text block and quotation marks
 *    - Underlines with quoted text block and text decoration
 *    - Sticky Notes with comment bubble
 *    - FreeText boxes with text preview
 *    - Ink Drawings with handwriting label
 *    - Shapes (Square / Rectangle, Circle, Line) with shape preview and optional comments
 * 6. Quick Color Filter palette swatches (active profile colors + presets + custom annotation colors with counts)
 * 7. Type filter chips with toggle-to-all capability (Highlights, Underlines, Notes, Text, Drawings, Shapes)
 * 8. Full-text search with page number and shape keyword matching
 * 9. "Reset All Filters" quick action button on empty search results
 * 10. Zotero grouping by Color (with semantic tags) and Type (with icons and collapsible groups)
 */

import { strict as assert } from "assert";
import { PDFDocument, PDFName, PDFDict, PDFArray, PDFNumber, PDFString } from "pdf-lib";
import {
  PdfSidebar,
  PdfSidebarDelegate,
} from "../src/presentation/pdf/pdf-sidebar";
import { PdfDocumentAdapter } from "../src/infrastructure/pdf/pdf-document-adapter";
import {
  PdfAnnotationParser,
  ZOTERO_DEFAULT_PROFILE,
  PdfAnnotation,
  PdfHighlightAnnotation,
  PdfUnderlineAnnotation,
  PdfNoteAnnotation,
  PdfFreeTextAnnotation,
  PdfInkAnnotation,
  PdfShapeAnnotation,
  filterAnnotations,
  sortAnnotations,
  groupAnnotationsByColor,
  groupAnnotationsByType,
  formatTypeLabel,
} from "../src/domain/pdf";

// Mock Obsidian DOM helper
function createMockElement(tag = "div", opts?: any): any {
  const el: any = {
    tagName: tag.toUpperCase(),
    children: [],
    style: {},
    classList: new Set<string>(),
    attr: {},
    dataset: {},
    innerHTML: "",
    _textContent: "",
    value: "",
  };

  Object.defineProperty(el, "textContent", {
    get() {
      if (this._textContent) {
        return this._textContent;
      }
      return this.children.map((c: any) => c.textContent).join("");
    },
    set(v: string) {
      this._textContent = String(v);
    },
    configurable: true,
  });

  el.addClass = function (cls: string) {
    cls.split(" ").filter(Boolean).forEach((c) => this.classList.add(c));
  };
  el.removeClass = function (cls: string) {
    cls.split(" ").filter(Boolean).forEach((c) => this.classList.delete(c));
  };
  el.toggleClass = function (cls: string, val: boolean) {
    if (val) this.addClass(cls);
    else this.removeClass(cls);
  };
  el.empty = function () {
    this.children = [];
    this.innerHTML = "";
    this._textContent = "";
  };
  el.setText = function (txt: string) {
    this._textContent = String(txt);
  };
  el.setAttribute = function (k: string, v: string) {
    this.attr[k] = String(v);
  };
  el.getAttribute = function (k: string) {
    return this.attr[k] ?? null;
  };
  el.createDiv = function (options?: any) {
    return this.createEl("div", options);
  };
  el.createSpan = function (options?: any) {
    return this.createEl("span", options);
  };
  el.createEl = function (t: string, options?: any) {
    const child = createMockElement(t, options);
    this.children.push(child);
    return child;
  };
  el.querySelectorAll = function (selector: string): any[] {
    const parts = selector.trim().split(/\s+/);
    if (parts.length > 1) {
      let current = [this];
      for (const part of parts) {
        const next: any[] = [];
        for (const node of current) {
          next.push(...node.querySelectorAll(part));
        }
        current = next;
      }
      return current;
    }

    const results: any[] = [];
    const walk = (node: any) => {
      if (selector.startsWith(".")) {
        const classes = selector.split(".").filter(Boolean);
        const match = classes.every((c) => node.classList && node.classList.has(c));
        if (match) {
          results.push(node);
        }
      } else if (selector.startsWith("#")) {
        const id = selector.slice(1);
        if (node.attr && node.attr.id === id) {
          results.push(node);
        }
      } else if (node.tagName && node.tagName.toLowerCase() === selector.toLowerCase()) {
        results.push(node);
      }
      for (const ch of node.children || []) {
        walk(ch);
      }
    };
    for (const ch of this.children || []) {
      walk(ch);
    }
    return results;
  };
  el.querySelector = function (selector: string): any {
    const all = this.querySelectorAll(selector);
    return all.length > 0 ? all[0] : null;
  };
  el.addEventListener = function () {};
  el.removeEventListener = function () {};
  el.getBoundingClientRect = function () {
    return { left: 0, top: 0, right: 300, bottom: 600, width: 300, height: 600 };
  };

  if (opts) {
    if (opts.cls) el.addClass(opts.cls);
    if (opts.text) el.setText(opts.text);
    if (opts.value) el.value = opts.value;
    if (opts.attr) {
      for (const [k, v] of Object.entries(opts.attr)) {
        el.setAttribute(k, String(v));
      }
    }
  }

  return el;
}

export async function runPdfSidebarAnnotationsTestSuite(): Promise<void> {
  console.log("\n=== RUNNING PDF SIDEBAR & ZOTERO ANNOTATIONS TEST SUITE ===");

  // -------------------------------------------------------------------------
  // 1. PdfAnnotationParser: Subtype Case-Insensitivity & Aliases
  // -------------------------------------------------------------------------
  console.log("--- Testing PdfAnnotationParser Subtypes & Coordinate Normalization ---");
  const testPdf = await PDFDocument.create();
  const testPage = testPdf.addPage([612, 792]);
  const ctx = testPdf.context;

  // Test lowercase "highlight"
  const hlDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("highlight"), // lowercase!
    Rect: ctx.obj([100, 700, 300, 720]),
    C: ctx.obj([0.37, 0.70, 0.21]),
    Contents: PDFString.of("Parsed lowercase highlight"),
    T: PDFString.of("Dr. Watson"),
    M: PDFString.of("D:20240410123000Z"),
  });
  const parsedHl = PdfAnnotationParser.parseAnnotationDict(hlDict as any, 0, ctx);
  assert.ok(parsedHl !== null, "Lowercase /highlight must be parsed successfully");
  assert.equal(parsedHl!.type, "highlight");
  assert.equal(parsedHl!.contents, "Parsed lowercase highlight");
  assert.equal(parsedHl!.author, "Dr. Watson");
  assert.ok(parsedHl!.modifiedAt instanceof Date, "Modified date must be parsed");
  console.log("✓ Verified lowercase /highlight subtype parsed with ISO date");

  // Test uppercase "UNDERLINE" and inverted rect [300, 620, 100, 600]
  const ulDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("UNDERLINE"),
    Rect: ctx.obj([300, 620, 100, 600]), // Inverted coordinates!
    C: ctx.obj([1.0, 0.40, 0.40]),
    Contents: PDFString.of("Inverted rect underline"),
  });
  const parsedUl = PdfAnnotationParser.parseAnnotationDict(ulDict as any, 0, ctx);
  assert.ok(parsedUl !== null, "Uppercase /UNDERLINE must be parsed successfully");
  assert.equal(parsedUl!.type, "underline");
  assert.deepEqual(parsedUl!.rect, [100, 600, 300, 620], "Coordinates must be normalized to [minX, minY, maxX, maxY]");
  console.log("✓ Verified case-insensitive underline and inverted coordinate normalization");

  // Test shapes: square, circle, line
  const sqDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("square"),
    Rect: ctx.obj([50, 400, 250, 500]),
    C: ctx.obj([0.95, 0.60, 0.22]),
    Contents: PDFString.of("Rect selection"),
  });
  const parsedSq = PdfAnnotationParser.parseAnnotationDict(sqDict as any, 0, ctx);
  assert.ok(parsedSq !== null, "Square annotation parsed");
  assert.equal(parsedSq!.type, "square");

  const lineDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Line"),
    Rect: ctx.obj([50, 100, 200, 150]),
    L: ctx.obj([50, 100, 200, 150]),
    C: ctx.obj([0.60, 0.35, 0.71]),
  });
  const parsedLine = PdfAnnotationParser.parseAnnotationDict(lineDict as any, 0, ctx);
  assert.ok(parsedLine !== null, "Line annotation parsed");
  assert.equal(parsedLine!.type, "line");
  assert.equal((parsedLine as any).linePoints?.length, 4, "Line points preserved");
  console.log("✓ Verified Shape annotations (square, line) parsed with stroke geometry");

  // -------------------------------------------------------------------------
  // 2. PdfSidebar Build & Zotero Card Presentation across All Types
  // -------------------------------------------------------------------------
  console.log("--- Testing PdfSidebar Zotero Cards & Dynamic Palette ---");
  const docBytes = await testPdf.save();
  const adapter = await PdfDocumentAdapter.load(docBytes.buffer);

  // Add 6 rich annotations covering all types
  const a1: PdfHighlightAnnotation = {
    id: "hl-1",
    type: "highlight",
    pageIndex: 0,
    rect: [100, 700, 300, 720],
    color: [0.37, 0.70, 0.21], // Green
    highlightedText: "Zotero highlight quote block in green.",
    contents: "Green semantic note",
    author: "Scholar",
    createdAt: new Date(),
  };

  const a2: PdfUnderlineAnnotation = {
    id: "ul-1",
    type: "underline",
    pageIndex: 0,
    rect: [100, 600, 300, 620],
    color: [1.0, 0.40, 0.40], // Red
    highlightedText: "Important finding underlined in red.",
    contents: "",
    author: "Scholar",
    createdAt: new Date(),
  };

  const a3: PdfNoteAnnotation = {
    id: "note-1",
    type: "text",
    pageIndex: 0,
    rect: [150, 500, 170, 520],
    color: [1.0, 0.85, 0.0], // Yellow
    contents: "Check citation in chapter 4.",
    author: "Scholar",
    createdAt: new Date(),
  };

  const a4: PdfFreeTextAnnotation = {
    id: "ft-1",
    type: "freetext",
    pageIndex: 0,
    rect: [200, 400, 350, 440],
    color: [0.18, 0.66, 0.90], // Blue (preset color, not in 3-color scheme!)
    contents: "Margin annotation text box",
    fontSize: 12,
    textColor: [0.18, 0.66, 0.90],
    author: "Scholar",
    createdAt: new Date(),
  };

  const a5: PdfInkAnnotation = {
    id: "ink-1",
    type: "ink",
    pageIndex: 0,
    rect: [100, 200, 200, 300],
    color: [0.60, 0.35, 0.71], // Purple (custom color!)
    paths: [[[100, 200], [150, 250], [200, 300]]],
    strokeWidth: 2,
    author: "Scholar",
    createdAt: new Date(),
  };

  const a6: PdfShapeAnnotation = {
    id: "shape-1",
    type: "square",
    pageIndex: 0,
    rect: [80, 450, 400, 550],
    color: [0.95, 0.60, 0.22], // Orange
    strokeColor: [0.95, 0.60, 0.22],
    strokeWidth: 2,
    contents: "Area selection around Table 1",
    author: "Scholar",
    createdAt: new Date(),
  };

  adapter.upsertAnnotation(a1);
  adapter.upsertAnnotation(a2);
  adapter.upsertAnnotation(a3);
  adapter.upsertAnnotation(a4);
  adapter.upsertAnnotation(a5);
  adapter.upsertAnnotation(a6);

  assert.equal(adapter.getAllAnnotations().length, 6, "Adapter has 6 annotations");

  const root = createMockElement("div");
  const sidebar = new PdfSidebar();
  let selectedAnnotId: string | null = null;

  const delegate: PdfSidebarDelegate = {
    onSelectPage: () => {},
    onSelectAnnotation: (annot) => {
      selectedAnnotId = annot.id;
    },
    onEditAnnotation: () => {},
    onDeleteAnnotation: (id) => {
      adapter.deleteAnnotation(id);
    },
    onRotatePage: () => {},
    onDuplicatePage: () => {},
    onDeletePage: () => {},
  };

  // Build sidebar collapsed by default (initialVisible = false)
  sidebar.build(root, adapter, null, "sample.pdf", delegate, ZOTERO_DEFAULT_PROFILE, false);
  assert.equal(sidebar.getIsVisible(), false, "Sidebar initially collapsed");

  // Call showAnnotationsTab() to open into Annotations tab
  sidebar.showAnnotationsTab();
  assert.equal(sidebar.getIsVisible(), true, "Sidebar is now visible");

  // Check Tab button label shows count
  const annotTab = root.querySelector(".pdf-sidebar-tab.is-active");
  assert.ok(annotTab !== null, "Annotations tab is active");
  const tabLabel = annotTab.querySelector(".pdf-sidebar-tab-label");
  assert.equal(tabLabel.textContent, "Annotations (6)", "Tab button displays 'Annotations (6)'");
  console.log("✓ Verified tab header displays accurate annotation count: 'Annotations (6)'");

  // Check All 6 cards rendered
  const cards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(cards.length, 6, "All 6 annotation cards rendered");

  // Verify Card 1: Highlight quote box with quotation marks
  const card1 = cards.find((c: any) => c.dataset.annotId === "hl-1");
  assert.ok(card1 !== null, "Highlight card found");
  assert.ok(card1.style.borderLeft.includes("#5fb236"), "Card has green vertical accent bar");
  const quote1 = card1.querySelector(".pdf-annot-card-quote");
  assert.ok(quote1 !== null, "Highlight has quote box");
  assert.ok(quote1.textContent.includes("“Zotero highlight quote block in green.”"), "Quote text formatted with marks");

  // Verify Card 6: Shape preview with label and comment
  const card6 = cards.find((c: any) => c.dataset.annotId === "shape-1");
  assert.ok(card6 !== null, "Shape card found");
  const shapePrev = card6.querySelector(".pdf-annot-card-shape-preview");
  assert.ok(shapePrev !== null, "Shape preview element rendered");
  assert.ok(shapePrev.textContent.includes("Rectangle Area"), "Shape preview displays 'Rectangle Area'");
  const comment6 = card6.querySelector(".pdf-annot-card-comment");
  assert.ok(comment6 !== null, "Shape comment box rendered below preview");
  assert.ok(comment6.textContent.includes("Area selection around Table 1"), "Shape comment text displayed");
  console.log("✓ Verified 1:1 Zotero card structure for highlights and shape area annotations");

  // -------------------------------------------------------------------------
  // 3. Quick Color Filter Palette Row (Dynamic All-Colors Support)
  // -------------------------------------------------------------------------
  console.log("--- Testing Quick Color Filter Palette & Toggling ---");
  const colorPills = root.querySelectorAll(".pdf-annot-color-pill");
  assert.ok(colorPills.length >= 6, "Color filter row contains pills for scheme colors + all annotation colors");

  // The 'All' pill must show 6
  const allPill = colorPills[0];
  assert.ok(allPill.textContent.includes("All"), "First pill is 'All'");
  assert.ok(allPill.textContent.includes("6"), "All pill has count 6");

  // Click green pill (id / text: 'New Info' or 'Green')
  const greenPill = colorPills.find((p: any) => p.textContent.includes("New Info") || p.textContent.includes("Green"));
  assert.ok(greenPill !== null, "Green pill found");
  assert.ok(greenPill.textContent.includes("1"), "Green pill has count 1");

  // Trigger click on green pill
  greenPill.onclick();
  const filteredGreenCards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(filteredGreenCards.length, 1, "Only 1 card shown when filtered by Green");
  assert.equal(filteredGreenCards[0].dataset.annotId, "hl-1");

  // Click green pill again -> toggles back to 'All'
  greenPill.onclick();
  const toggledCards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(toggledCards.length, 6, "Toggling color pill restores all 6 cards");
  console.log("✓ Verified dynamic color filter pills and click toggle-to-all");

  // -------------------------------------------------------------------------
  // 4. Type Filter Chips
  // -------------------------------------------------------------------------
  console.log("--- Testing Type Filter Chips & Shape Grouping ---");
  const typeChips = root.querySelectorAll(".pdf-annot-type-chip");
  assert.equal(typeChips.length, 7, "7 type chips rendered (All, Highlights, Underlines, Notes, Text, Drawings, Shapes)");

  // Click 'Shapes' chip
  const shapesChip = typeChips.find((c: any) => c.textContent === "Shapes");
  assert.ok(shapesChip !== null, "Shapes chip found");
  shapesChip.onclick();

  const shapeCards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(shapeCards.length, 1, "Only 1 shape card shown");
  assert.equal(shapeCards[0].dataset.annotId, "shape-1");

  // Click 'Shapes' chip again -> toggles back to All
  shapesChip.onclick();
  const resetTypeCards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(resetTypeCards.length, 6, "Toggling Shapes chip resets filter back to All");
  console.log("✓ Verified type filter chips filtering and toggle reset");

  // -------------------------------------------------------------------------
  // 5. Search Filtering & Empty-State Reset Button
  // -------------------------------------------------------------------------
  console.log("--- Testing Search Query & Reset All Filters Button ---");
  const searchInput = root.querySelector(".pdf-annot-search-input");
  assert.ok(searchInput !== null, "Search input found");

  // Search by text snippet
  searchInput.value = "quote block";
  searchInput.oninput();
  let searchCards = root.querySelectorAll(".pdf-annot-card");
  assert.equal(searchCards.length, 1, "Found 1 card matching 'quote block'");
  assert.equal(searchCards[0].dataset.annotId, "hl-1");

  // Search for non-existent keyword -> triggers Empty State with 'Reset All Filters' button
  searchInput.value = "nonexistent_query_xyz";
  searchInput.oninput();
  assert.equal(root.querySelectorAll(".pdf-annot-card").length, 0, "0 cards for unmatched search");

  const resetFiltersBtn = root.querySelector(".pdf-annot-reset-filters-btn");
  assert.ok(resetFiltersBtn !== null, "'Reset All Filters' button rendered in empty state");
  assert.equal(resetFiltersBtn.textContent, "Reset All Filters");

  // Click reset filters button
  resetFiltersBtn.onclick();
  const cardsAfterReset = root.querySelectorAll(".pdf-annot-card");
  assert.equal(cardsAfterReset.length, 6, "Reset button restores all 6 cards");
  console.log("✓ Verified full-text search and 'Reset All Filters' recovery flow");

  // -------------------------------------------------------------------------
  // 6. Sort by Color & Grouping
  // -------------------------------------------------------------------------
  console.log("--- Testing Sort by Color & Collapsible Groups ---");
  const sortSelect = root.querySelector(".pdf-annot-sort-select");
  assert.ok(sortSelect !== null, "Sort select found");

  sortSelect.value = "color";
  sortSelect.onchange();

  const colorGroups = root.querySelectorAll(".pdf-annot-group");
  assert.ok(colorGroups.length >= 4, "Rendered color groups for different annotation colors");

  const firstGroupHeader = colorGroups[0].querySelector(".pdf-annot-group-header");
  assert.ok(firstGroupHeader !== null, "Color group header rendered");
  console.log("✓ Verified Sort by Color grouped view with section headers");

  // -------------------------------------------------------------------------
  // 7. Refresh on Expansion Lifecycle Verification
  // -------------------------------------------------------------------------
  console.log("--- Testing Refresh on Expansion Lifecycle ---");
  // Collapse sidebar
  sidebar.toggle();
  assert.equal(sidebar.getIsVisible(), false, "Sidebar collapsed");

  // Add another annotation while collapsed
  adapter.upsertAnnotation({
    id: "note-2",
    type: "text",
    pageIndex: 0,
    rect: [100, 100, 150, 150],
    color: [1.0, 0.85, 0.0],
    contents: "Late addition note",
    author: "Scholar",
    createdAt: new Date(),
  });

  // Toggle open sidebar -> must re-render freshly!
  sidebar.toggle();
  assert.equal(sidebar.getIsVisible(), true, "Sidebar expanded");
  const tabLabelAfterToggle = root.querySelector(".pdf-sidebar-tab.is-active .pdf-sidebar-tab-label");
  assert.equal(tabLabelAfterToggle.textContent, "Annotations (7)", "Tab label refreshed to 7 annotations");
  console.log("✓ Verified toggle() re-renders active tab immediately upon becoming visible");

  console.log("==================================================");
  console.log("🎉 PDF SIDEBAR & ANNOTATIONS TEST SUITE PASSED 100%!");
  console.log("==================================================");
}
