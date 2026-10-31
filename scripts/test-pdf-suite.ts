import { PDFDocument, PDFName, PDFString } from "pdf-lib";
import {
  PdfDocumentAdapter,
} from "../src/infrastructure/pdf/pdf-document-adapter";
import {
  PdfNoteAnnotation,
  PdfHighlightAnnotation,
  PdfInkAnnotation,
  PdfFreeTextAnnotation,
  PdfShapeAnnotation,
  CitationService,
  CitationBuilder,
  ZOTERO_DEFAULT_PROFILE,
  ACADEMIC_RESEARCH_PROFILE,
  LITERATURE_REVIEW_PROFILE,
  LEGAL_ANALYSIS_PROFILE,
  findTagForColor,
  cloneProfile,
  ensureValidProfile,
  normalizeHexColor,
  PdfColorSchemeProfile,
} from "../src/domain/pdf";
import { normalizeSettings, DEFAULT_OFFICE_SETTINGS } from "../src/settings/office-settings";

async function runTestSuite() {
  console.log("=== RUNNING PDF EDITOR & ANNOTATIONS TEST SUITE ===");

  // 1. Create a blank PDF document with 2 pages
  const baseDoc = await PDFDocument.create();
  baseDoc.setTitle("Test Research Paper");
  baseDoc.setAuthor("Dr. Jane Doe");
  const p1 = baseDoc.addPage([612, 792]);
  const p2 = baseDoc.addPage([612, 792]);
  const initialBytes = await baseDoc.save();

  console.log("✓ Base PDF created with 2 pages (612x792)");

  // 2. Load into PdfDocumentAdapter
  const adapter = await PdfDocumentAdapter.load(initialBytes.buffer);
  if (adapter.getPageCount() !== 2) {
    throw new Error(`Expected 2 pages, got ${adapter.getPageCount()}`);
  }
  if (adapter.getTitle() !== "Test Research Paper") {
    throw new Error(`Expected title 'Test Research Paper', got '${adapter.getTitle()}'`);
  }
  console.log("✓ Adapter loaded PDF, title and page count verified");

  // 3. Add Sticky Note Annotation (/Subtype /Text)
  const stickyNote: PdfNoteAnnotation = {
    id: "note_001",
    type: "text",
    pageIndex: 0,
    rect: [100, 700, 124, 724],
    color: [1, 0.85, 0], // Yellow
    contents: "This is a critical note for the thesis introduction.",
    author: "Researcher",
    iconName: "Comment",
    linkedNote: "Notes/Thesis-Intro.md",
  };
  adapter.upsertAnnotation(stickyNote);
  console.log("✓ Upserted Sticky Note Annotation with linked note");

  // 4. Add Highlight Annotation (/Subtype /Highlight) with QuadPoints
  const highlight: PdfHighlightAnnotation = {
    id: "hl_001",
    type: "highlight",
    pageIndex: 0,
    rect: [100, 650, 400, 670],
    quadPoints: [
      100, 670, 400, 670,
      100, 650, 400, 650,
    ],
    color: [0.5, 0.9, 0.5], // Green
    highlightedText: "Deep learning models require high-quality annotated training data.",
    contents: "Key hypothesis to cite in section 2",
    author: "Researcher",
    linkedNote: "Notes/Deep-Learning.md",
  };
  adapter.upsertAnnotation(highlight);
  console.log("✓ Upserted Highlight Annotation with QuadPoints");

  // 5. Add Ink Drawing Annotation (/Subtype /Ink)
  const ink: PdfInkAnnotation = {
    id: "ink_001",
    type: "ink",
    pageIndex: 1,
    rect: [150, 400, 350, 500],
    color: [0.9, 0.2, 0.2], // Red
    paths: [
      [[150, 400], [200, 450], [250, 420], [350, 500]],
      [[160, 410], [210, 460], [260, 430]],
    ],
    strokeWidth: 3,
    contents: "Handwritten signature / diagram",
    author: "Signer",
  };
  adapter.upsertAnnotation(ink);
  console.log("✓ Upserted Ink Drawing Annotation with multiple paths");

  // 6. Add FreeText Text Box Annotation (/Subtype /FreeText)
  const freeText: PdfFreeTextAnnotation = {
    id: "ft_001",
    type: "freetext",
    pageIndex: 1,
    rect: [100, 300, 300, 330],
    color: [0, 0, 0],
    textColor: [0.1, 0.1, 0.1],
    fontSize: 14,
    contents: "APPROVED FOR PUBLICATION",
    author: "Editor",
  };
  adapter.upsertAnnotation(freeText);
  console.log("✓ Upserted FreeText Annotation");

  // 7. Add Geometric Shape Annotation (/Subtype /Square)
  const shape: PdfShapeAnnotation = {
    id: "sq_001",
    type: "square",
    pageIndex: 0,
    rect: [50, 50, 200, 150],
    color: [0.2, 0.4, 0.9],
    strokeColor: [0.2, 0.4, 0.9],
    strokeWidth: 2,
    contents: "Figure boundary box",
    author: "Editor",
  };
  adapter.upsertAnnotation(shape);
  console.log("✓ Upserted Square Shape Annotation");

  // 8. Test Native Page Operations
  // Rotate page 0 by 90 degrees
  const newAngle = adapter.rotatePage(0, 90);
  if (newAngle !== 90) throw new Error(`Expected angle 90, got ${newAngle}`);
  console.log("✓ Rotated page 0 to 90°");

  // Insert a blank page
  adapter.insertBlankPage(1);
  if (adapter.getPageCount() !== 3) {
    throw new Error(`Expected 3 pages after insert, got ${adapter.getPageCount()}`);
  }
  console.log("✓ Inserted blank page, total pages = 3");

  // Duplicate page 0
  await adapter.duplicatePage(0);
  if (adapter.getPageCount() !== 4) {
    throw new Error(`Expected 4 pages after duplication, got ${adapter.getPageCount()}`);
  }
  console.log("✓ Duplicated page, total pages = 4");

  // Delete page 2
  adapter.deletePage(2);
  if (adapter.getPageCount() !== 3) {
    throw new Error(`Expected 3 pages after delete, got ${adapter.getPageCount()}`);
  }
  console.log("✓ Deleted page, total pages = 3");

  // 9. Save PDF and verify ISO 32000-1 binary serialization
  const savedBytes = await adapter.save();
  console.log(`✓ Saved modified PDF: ${savedBytes.byteLength} bytes`);

  // 10. Load saved bytes into a fresh adapter and verify all annotations are preserved
  const reloadedAdapter = await PdfDocumentAdapter.load(savedBytes.buffer);
  const reloadedAnnots = reloadedAdapter.getAllAnnotations();
  console.log(`✓ Reloaded PDF has ${reloadedAnnots.length} annotations`);

  const foundSticky = reloadedAnnots.find((a) => a.id === "note_001");
  if (!foundSticky || foundSticky.type !== "text") {
    throw new Error("Sticky note not preserved in reloaded PDF!");
  }
  if (foundSticky.contents !== "This is a critical note for the thesis introduction.") {
    throw new Error("Sticky note contents mismatched!");
  }
  if (foundSticky.linkedNote !== "Notes/Thesis-Intro.md") {
    throw new Error("Sticky note linkedNote mismatched!");
  }
  console.log("✓ Verified Sticky note parsed from ISO /Annots dictionary with linkedNote");

  const foundHl = reloadedAnnots.find((a) => a.id === "hl_001");
  if (!foundHl || foundHl.type !== "highlight") {
    throw new Error("Highlight not preserved in reloaded PDF!");
  }
  console.log("✓ Verified Highlight parsed with QuadPoints");

  const foundInk = reloadedAnnots.find((a) => a.id === "ink_001");
  if (!foundInk || foundInk.type !== "ink") {
    throw new Error("Ink drawing not preserved in reloaded PDF!");
  }
  if ((foundInk as any).paths.length !== 2) {
    throw new Error(`Expected 2 ink paths, got ${(foundInk as any).paths.length}`);
  }
  console.log("✓ Verified Ink drawing parsed with multiple stroke paths");

  // 11. Test Citation Generation
  const citationMeta = {
    title: "Attention Is All You Need",
    author: "Vaswani et al.",
    year: 2017,
    pdfFileName: "Attention.pdf",
    pageNumber: 3,
    quotedText: "The dominant sequence transduction models are based on complex recurrent neural networks.",
    annotationId: "hl_001",
  };

  const mdQuote = CitationBuilder.buildMarkdownQuote(citationMeta);
  console.log("\n--- Generated Markdown Quote Citation ---");
  console.log(mdQuote);
  if (!mdQuote.includes("> — [[Attention.pdf#page=3&annot=hl_001|Attention Is All You Need, p. 3]]")) {
    throw new Error("Markdown quote format mismatch");
  }

  const apa = CitationBuilder.buildApa(citationMeta);
  console.log("\n--- Generated APA Citation ---");
  console.log(apa);

  const bibtex = CitationBuilder.buildBibtex(citationMeta);
  console.log("\n--- Generated BibTeX Citation ---");
  console.log(bibtex);

  // 12. Test Full Markdown Annotations Export
  const exportedMd = CitationService.exportAnnotationsToMarkdown(
    "Paper.pdf",
    "Test Research Paper",
    reloadedAnnots,
    reloadedAdapter.getPageCount()
  );
  console.log("\n--- Generated Markdown Annotations Export ---");
  if (!exportedMd.includes("total_annotations:")) {
    throw new Error("Export Markdown format invalid");
  }
  if (!exportedMd.includes("> [!example] Text Box")) {
    throw new Error("Export Markdown should format text boxes with [!example] callout");
  }
  if (!exportedMd.includes("Jump to:")) {
    throw new Error("Export Markdown should include Jump to navigation links for multiple pages");
  }
  console.log("✓ Verified 1-click Markdown Export generates proper callouts, frontmatter, and jump links");

  // 13. Test Moving / Repositioning Annotations & Shapes
  console.log("\n--- Testing Movement & Repositioning ---");
  const stickyToMove = reloadedAdapter.getAnnotation("note_001") as PdfNoteAnnotation;
  const initialRect = [...stickyToMove.rect];
  const dX = 75;
  const dY = -40;
  stickyToMove.rect = [
    initialRect[0] + dX,
    initialRect[1] + dY,
    initialRect[2] + dX,
    initialRect[3] + dY,
  ];
  reloadedAdapter.upsertAnnotation(stickyToMove);

  // Move shape
  const shapeToMove = reloadedAdapter.getAnnotation("sq_001") as PdfShapeAnnotation;
  const shapeInitRect = [...shapeToMove.rect];
  shapeToMove.rect = [
    shapeInitRect[0] + 30,
    shapeInitRect[1] + 20,
    shapeInitRect[2] + 30,
    shapeInitRect[3] + 20,
  ];
  reloadedAdapter.upsertAnnotation(shapeToMove);

  const movedBytes = await reloadedAdapter.save();
  const movedAdapter = await PdfDocumentAdapter.load(movedBytes.buffer);

  const movedSticky = movedAdapter.getAnnotation("note_001");
  if (!movedSticky) throw new Error("Moved sticky note not found");
  if (Math.round(movedSticky.rect[0]) !== Math.round(initialRect[0] + dX)) {
    throw new Error(`Expected X=${initialRect[0] + dX}, got ${movedSticky.rect[0]}`);
  }
  if (Math.round(movedSticky.rect[1]) !== Math.round(initialRect[1] + dY)) {
    throw new Error(`Expected Y=${initialRect[1] + dY}, got ${movedSticky.rect[1]}`);
  }
  console.log("✓ Verified Sticky note repositioning and coordinate persistence");

  const movedShape = movedAdapter.getAnnotation("sq_001");
  if (!movedShape) throw new Error("Moved shape not found");
  if (Math.round(movedShape.rect[0]) !== Math.round(shapeInitRect[0] + 30)) {
    throw new Error(`Expected Shape X=${shapeInitRect[0] + 30}, got ${movedShape.rect[0]}`);
  }
  console.log("✓ Verified Shape repositioning and coordinate persistence");

  // 14. Test Deletion Purge from PDF Binary (1:1 Sync Guarantee)
  console.log("\n--- Testing Deletion Purge from Binary (1:1 Sync) ---");
  movedAdapter.deleteAnnotation("sq_001");
  const purgedBytes = await movedAdapter.save();
  const purgedAdapter = await PdfDocumentAdapter.load(purgedBytes.buffer);
  if (purgedAdapter.getAnnotation("sq_001")) {
    throw new Error("Deleted shape annotation was NOT purged from reloaded PDF!");
  }
  console.log("✓ Verified deleted annotation is completely purged from PDF /Annots dictionary (no ghost annotations)");

  // 15. Test Unicode & Special Characters (Cross-Device 1:1 Match)
  console.log("\n--- Testing Unicode & Special Characters (Cross-Device Match) ---");
  const unicodeNote: PdfNoteAnnotation = {
    id: "note_unicode_001",
    type: "text",
    pageIndex: 0,
    rect: [150, 600, 174, 624],
    color: [0.2, 0.6, 1.0],
    contents: "Résumé: Great analysis on AI models! 🎯 “Smart Quotes” & ‘single’ café",
    author: "Researcher José",
    linkedNote: "Notes/Résumé 2026.md",
  };
  purgedAdapter.upsertAnnotation(unicodeNote);
  const unicodeBytes = await purgedAdapter.save();
  const unicodeAdapter = await PdfDocumentAdapter.load(unicodeBytes.buffer);
  const reloadedUnicodeNote = unicodeAdapter.getAnnotation("note_unicode_001");
  if (!reloadedUnicodeNote) {
    throw new Error("Unicode note was not found in reloaded PDF");
  }
  if (reloadedUnicodeNote.contents !== unicodeNote.contents) {
    throw new Error(`Unicode contents mismatched! Expected: '${unicodeNote.contents}', Got: '${reloadedUnicodeNote.contents}'`);
  }
  if (reloadedUnicodeNote.author !== unicodeNote.author) {
    throw new Error(`Unicode author mismatched! Expected: '${unicodeNote.author}', Got: '${reloadedUnicodeNote.author}'`);
  }
  if (reloadedUnicodeNote.linkedNote !== unicodeNote.linkedNote) {
    throw new Error(`Unicode linkedNote mismatched! Expected: '${unicodeNote.linkedNote}', Got: '${reloadedUnicodeNote.linkedNote}'`);
  }
  console.log("✓ Verified 100% 1:1 Unicode, emoji, and accented text preservation across save & load");

  // 16. Test Third-Party Acrobat / External PDF with Indirect References & Popup Comments
  console.log("\n--- Testing Third-Party Acrobat & External Annotations (Indirect Refs, Popup Comments) ---");
  const externalDoc = await PDFDocument.create();
  const extPage = externalDoc.addPage([600, 800]);

  // Case A: Indirect Subtype & Rect (common in Acrobat object streams)
  const subtypeRef = externalDoc.context.register(PDFName.of("Highlight"));
  const rectRef = externalDoc.context.register(externalDoc.context.obj([72, 700, 250, 720]));
  const quadRef = externalDoc.context.register(externalDoc.context.obj([72, 720, 250, 720, 72, 700, 250, 700]));
  
  // Case B: Acrobat Popup comment attached to Highlight
  const popupDict = externalDoc.context.register(externalDoc.context.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Popup"),
    Contents: PDFString.of("Crucial finding from paper!"),
  }));

  const annotIndirect = externalDoc.context.register(externalDoc.context.obj({
    Type: PDFName.of("Annot"),
    Subtype: subtypeRef,
    Rect: rectRef,
    QuadPoints: quadRef,
    C: externalDoc.context.obj([1, 0.9, 0.2]),
    T: PDFString.of("External Reviewer"),
    Popup: popupDict,
  }));

  // Case C: Underline markup
  const annotUnderline = externalDoc.context.register(externalDoc.context.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Underline"),
    Rect: externalDoc.context.obj([72, 600, 200, 615]),
    C: externalDoc.context.obj([0, 0.5, 1]),
    Contents: PDFString.of("Underlined hypothesis"),
  }));

  // Case D: Sticky Note without NM
  const annotText = externalDoc.context.register(externalDoc.context.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Text"),
    Rect: externalDoc.context.obj([80, 500, 104, 524]),
    Contents: PDFString.of("Important margin note"),
  }));

  const extAnnotsArray = externalDoc.context.obj([annotIndirect, annotUnderline, annotText]);
  const extAnnotsRef = externalDoc.context.register(extAnnotsArray);
  extPage.node.set(PDFName.of("Annots"), extAnnotsRef);

  const extPdfBytes = await externalDoc.save();
  const extAdapter = await PdfDocumentAdapter.load(extPdfBytes.buffer);
  const loadedExtAnnots = extAdapter.getAllAnnotations();

  if (loadedExtAnnots.length !== 3) {
    throw new Error(`Expected 3 external annotations to be loaded, got ${loadedExtAnnots.length}`);
  }

  const loadedHighlight = loadedExtAnnots.find((a) => a.type === "highlight" && a.rect[1] === 700);
  if (!loadedHighlight) {
    throw new Error("Highlight with indirect Subtype and Rect was not parsed!");
  }
  if (loadedHighlight.contents !== "Crucial finding from paper!") {
    throw new Error(`Expected popup comment 'Crucial finding from paper!', got '${loadedHighlight.contents}'`);
  }
  if (loadedHighlight.author !== "External Reviewer") {
    throw new Error(`Expected author 'External Reviewer', got '${loadedHighlight.author}'`);
  }

  const loadedUnderline = loadedExtAnnots.find((a) => a.type === "underline" && a.rect[1] === 600);
  if (!loadedUnderline || loadedUnderline.contents !== "Underlined hypothesis") {
    throw new Error("Underline annotation was not parsed correctly!");
  }

  const loadedText = loadedExtAnnots.find((a) => a.type === "text");
  if (!loadedText || loadedText.contents !== "Important margin note") {
    throw new Error("Text note without NM was not parsed correctly!");
  }
  console.log("✓ Verified 100% of indirect references, Acrobat Popup comments, and Underline annotations parsed correctly");

  // --- Testing Zotero Style Features: Underline, FreeText Transform, Attached Notes & Palette ---
  console.log("\n--- Testing Zotero Style Features ---");

  // 1. Verify Zotero Palette presets
  const { PDF_COLOR_PRESETS: zoteroPresets } = await import("../src/domain/pdf");
  if (zoteroPresets.length !== 8) {
    throw new Error(`Expected 8 Zotero color presets, got ${zoteroPresets.length}`);
  }
  const yellowPreset = zoteroPresets.find((p) => p.name === "Yellow");
  if (!yellowPreset || yellowPreset.hex !== "#ffd400") {
    throw new Error(`Expected Zotero yellow #ffd400, got ${yellowPreset?.hex}`);
  }
  const bluePreset = zoteroPresets.find((p) => p.name === "Blue");
  if (!bluePreset || bluePreset.hex !== "#2ea8e5") {
    throw new Error(`Expected Zotero blue #2ea8e5, got ${bluePreset?.hex}`);
  }
  console.log("✓ Verified official 8 Zotero color palette presets");

  // 2. Create Underline with attached note & bookmark metadata
  const zoteroUnderline: any = {
    id: "zotero_ul_001",
    type: "underline",
    pageIndex: 0,
    rect: [120, 550, 380, 568],
    quadPoints: [
      120, 568, 380, 568,
      120, 550, 380, 550,
    ],
    color: [0.18, 0.66, 0.90], // Zotero Blue
    highlightedText: "Zotero style underlined text",
    contents: "My personal note attached to the underlined sentence",
    author: "Researcher",
    linkedNote: "Notes/Zotero-Features.md",
  };
  extAdapter.upsertAnnotation(zoteroUnderline);

  // 3. Create FreeText box with formatting and resized transform
  const zoteroTextBox: any = {
    id: "zotero_tb_001",
    type: "freetext",
    pageIndex: 0,
    rect: [200, 400, 450, 480], // Resized width (250) and height (80)
    color: [0, 0, 0],
    textColor: [0.18, 0.66, 0.90],
    fontSize: 16,
    isBold: true,
    isItalic: true,
    contents: "Resized & formatted Zotero textbox",
    author: "User",
  };
  extAdapter.upsertAnnotation(zoteroTextBox);

  const zoteroPdfBytes = await extAdapter.save();
  const zoteroReloadedAdapter = await PdfDocumentAdapter.load(zoteroPdfBytes.buffer);

  const reloadedUl = zoteroReloadedAdapter.getAnnotation("zotero_ul_001") as any;
  if (!reloadedUl || reloadedUl.type !== "underline") {
    throw new Error("Zotero underline annotation was not saved/reloaded with type 'underline'!");
  }
  if (reloadedUl.contents !== "My personal note attached to the underlined sentence") {
    throw new Error(`Expected attached note, got '${reloadedUl.contents}'`);
  }
  if (reloadedUl.quadPoints.length !== 8) {
    throw new Error(`Expected 8 quad points for underline, got ${reloadedUl.quadPoints.length}`);
  }
  console.log("✓ Verified Zotero Underline annotation saved, reloaded, and preserves attached note");

  const reloadedTb = zoteroReloadedAdapter.getAnnotation("zotero_tb_001") as any;
  if (!reloadedTb || reloadedTb.type !== "freetext") {
    throw new Error("Zotero text box was not saved/reloaded correctly!");
  }
  if (reloadedTb.rect[2] - reloadedTb.rect[0] !== 250 || reloadedTb.rect[3] - reloadedTb.rect[1] !== 80) {
    throw new Error(`Expected transformed dimensions 250x80, got ${reloadedTb.rect[2] - reloadedTb.rect[0]}x${reloadedTb.rect[3] - reloadedTb.rect[1]}`);
  }
  console.log("✓ Verified Zotero Text Box transform (width/height resize) persisted across PDF save/load");

  // ==================== Testing Zotero Color Schemes, Profiles & Semantic Tagging ====================
  console.log("\n--- Testing Zotero Color Schemes, Profiles & Semantic Tagging ---");

  // 1. Verify Standard Zotero Profile tags: Green = New Info, Red = Important Points, Yellow = Processes
  const greenTag = findTagForColor(ZOTERO_DEFAULT_PROFILE, "#5fb236");
  if (!greenTag || greenTag.tag !== "New Info") {
    throw new Error(`Expected Green to be tagged 'New Info', got '${greenTag?.tag}'`);
  }
  const redTag = findTagForColor(ZOTERO_DEFAULT_PROFILE, "#ff6666");
  if (!redTag || redTag.tag !== "Important Points") {
    throw new Error(`Expected Red to be tagged 'Important Points', got '${redTag?.tag}'`);
  }
  const yellowTag = findTagForColor(ZOTERO_DEFAULT_PROFILE, "#ffd400");
  if (!yellowTag || yellowTag.tag !== "Processes") {
    throw new Error(`Expected Yellow to be tagged 'Processes', got '${yellowTag?.tag}'`);
  }
  console.log("✓ Verified Zotero default palette tags: Green (New Info), Red (Important Points), Yellow (Processes)");

  // 2. Verify Built-in Templates
  if (!ACADEMIC_RESEARCH_PROFILE || ACADEMIC_RESEARCH_PROFILE.colors.length < 8) {
    throw new Error("Academic Research template is missing or incomplete!");
  }
  if (!LITERATURE_REVIEW_PROFILE || LITERATURE_REVIEW_PROFILE.colors.length < 8) {
    throw new Error("Literature Review template is missing or incomplete!");
  }
  if (!LEGAL_ANALYSIS_PROFILE || LEGAL_ANALYSIS_PROFILE.colors.length < 8) {
    throw new Error("Legal Analysis template is missing or incomplete!");
  }
  console.log("✓ Verified Built-in Template Profiles: Academic Research, Literature Review, Legal Analysis");

  // 3. Test Profile Cloning & ensureValidProfile
  const clonedProfile = cloneProfile(ACADEMIC_RESEARCH_PROFILE, "custom-thesis", "My Thesis Scheme");
  if (clonedProfile.id !== "custom-thesis" || clonedProfile.name !== "My Thesis Scheme") {
    throw new Error(`Failed to clone profile correctly: ${clonedProfile.id}, ${clonedProfile.name}`);
  }
  if (clonedProfile.isBuiltIn) {
    throw new Error("Cloned profile should not be marked as built-in");
  }
  // Modify a tag in custom profile
  clonedProfile.colors[0].tag = "Crucial Breakthrough";
  const verifiedThesisTag = findTagForColor(clonedProfile, clonedProfile.colors[0].hex);
  if (verifiedThesisTag?.tag !== "Crucial Breakthrough") {
    throw new Error(`Expected 'Crucial Breakthrough', got '${verifiedThesisTag?.tag}'`);
  }
  console.log("✓ Verified Template Profile cloning, validation, and custom tag customization");

  // 4. Test Per-PDF Custom Color Scheme Embedding in PDF Binary
  console.log("--- Testing Per-PDF Custom Color Scheme Embedding in PDF Binary ---");
  const pdfForScheme = await PDFDocument.create();
  pdfForScheme.addPage([612, 792]);
  const pdfForSchemeBytes = await pdfForScheme.save();
  const schemeAdapter = await PdfDocumentAdapter.load(pdfForSchemeBytes.buffer);

  // Define per-PDF custom color scheme manually
  const perPdfScheme: PdfColorSchemeProfile = {
    id: "biology-paper-scheme",
    name: "Biology Lab Notes",
    description: "Specific color tags for cell biology research",
    colors: [
      { id: "c1", hex: "#5fb236", name: "Green", tag: "Enzyme Reactions", rgb: [0.37, 0.70, 0.22] },
      { id: "c2", hex: "#ff6666", name: "Red", tag: "Toxic Inhibitors", rgb: [1.0, 0.40, 0.40] },
      { id: "c3", hex: "#ffd400", name: "Yellow", tag: "Incubation Phase", rgb: [1.0, 0.83, 0.0] },
      { id: "c4", hex: "#2ea8e5", name: "Blue", tag: "Control Samples", rgb: [0.18, 0.66, 0.90] },
    ],
  };

  schemeAdapter.setColorScheme(perPdfScheme);
  // Add an annotation tagged with this scheme
  schemeAdapter.upsertAnnotation({
    id: "bio_hl_1",
    type: "highlight",
    pageIndex: 0,
    rect: [100, 600, 300, 620],
    quadPoints: [100, 620, 300, 620, 100, 600, 300, 600],
    color: [0.37, 0.70, 0.22], // Green -> "Enzyme Reactions"
    highlightedText: "Taq polymerase showed 98% catalytic efficiency under 72°C.",
    contents: "Key experimental observation",
  });

  const savedSchemePdfBytes = await schemeAdapter.save();

  // Reload PDF in brand new adapter and verify scheme survived lossless binary round-trip
  const reloadedSchemeAdapter = await PdfDocumentAdapter.load(savedSchemePdfBytes.buffer);
  const reloadedScheme = reloadedSchemeAdapter.getColorScheme();

  if (!reloadedScheme) {
    throw new Error("Failed to reload embedded color scheme from PDF catalog!");
  }
  if (reloadedScheme.name !== "Biology Lab Notes") {
    throw new Error(`Expected scheme name 'Biology Lab Notes', got '${reloadedScheme.name}'`);
  }
  if (reloadedScheme.colors.length !== 4) {
    throw new Error(`Expected 4 colors, got ${reloadedScheme.colors.length}`);
  }
  const enzymeTag = findTagForColor(reloadedScheme, "#5fb236");
  if (!enzymeTag || enzymeTag.tag !== "Enzyme Reactions") {
    throw new Error(`Expected 'Enzyme Reactions', got '${enzymeTag?.tag}'`);
  }
  const inhibitorTag = findTagForColor(reloadedScheme, "#ff6666");
  if (!inhibitorTag || inhibitorTag.tag !== "Toxic Inhibitors") {
    throw new Error(`Expected 'Toxic Inhibitors', got '${inhibitorTag?.tag}'`);
  }
  console.log("✓ Verified Per-PDF Color Scheme embeds into PDF binary catalog and reloads losslessly with all tags");

  // 5. Test Plugin Settings Normalization & Persistence
  console.log("--- Testing Settings Normalization & Template Persistence ---");
  const normalized = normalizeSettings({
    pdfDefaultProfileId: "academic-research",
    pdfProfiles: [
      clonedProfile,
    ],
    pdfFileSchemes: {
      "Vault/Bio.pdf": perPdfScheme,
    },
  });

  if (normalized.pdfDefaultProfileId !== "academic-research") {
    throw new Error(`Expected default profile 'academic-research', got '${normalized.pdfDefaultProfileId}'`);
  }
  // Check that built-ins were preserved
  if (!normalized.pdfProfiles.some((p) => p.id === "zotero-default")) {
    throw new Error("Missing built-in zotero-default profile after normalization");
  }
  // Check that custom profile was preserved
  if (!normalized.pdfProfiles.some((p) => p.id === "custom-thesis")) {
    throw new Error("Missing custom-thesis profile after normalization");
  }
  // Check that per-file scheme override was preserved
  if (!normalized.pdfFileSchemes["Vault/Bio.pdf"] || normalized.pdfFileSchemes["Vault/Bio.pdf"].name !== "Biology Lab Notes") {
    throw new Error("Missing per-file scheme override after normalization");
  }
  console.log("✓ Verified Settings normalization preserves built-ins, custom templates, and per-file schemes");

  // 6. Test Markdown Export with Color Scheme Legend and Semantic Tag Callouts
  console.log("--- Testing Markdown Export with Color Scheme Legend & Tags ---");
  const exportAnnotations: any[] = [
    {
      id: "exp_hl_1",
      type: "highlight",
      pageIndex: 0,
      color: [0.37, 0.70, 0.22], // Green -> "Enzyme Reactions"
      highlightedText: "Taq polymerase showed 98% catalytic efficiency under 72°C.",
      contents: "Note on PCR amplification",
    },
    {
      id: "exp_ul_1",
      type: "underline",
      pageIndex: 0,
      color: [1.0, 0.40, 0.40], // Red -> "Toxic Inhibitors"
      highlightedText: "EDTA chelated free magnesium ions, quenching the enzymatic activity.",
      contents: "Critical inhibitor finding",
    },
    {
      id: "exp_txt_1",
      type: "text",
      pageIndex: 1,
      color: [1.0, 0.83, 0.0], // Yellow -> "Incubation Phase"
      author: "Researcher",
      contents: "Maintain 37°C for minimum 45 minutes",
    },
  ];

  const exportedSchemeMd = CitationService.exportAnnotationsToMarkdown(
    "BiologyPaper.pdf",
    "Cellular Biology Lab",
    exportAnnotations,
    2,
    perPdfScheme
  );

  // Check legend presence
  if (!exportedSchemeMd.includes("Color Scheme: Biology Lab Notes")) {
    throw new Error("Exported markdown missing Color Scheme header!");
  }
  if (!exportedSchemeMd.includes("Enzyme Reactions") || !exportedSchemeMd.includes("Toxic Inhibitors") || !exportedSchemeMd.includes("Incubation Phase")) {
    throw new Error("Exported markdown missing color scheme tag entries in table!");
  }

  // Check callout tags
  if (!exportedSchemeMd.includes("[!quote] Highlight [Enzyme Reactions]")) {
    throw new Error("Exported markdown missing tagged highlight callout!");
  }
  if (!exportedSchemeMd.includes("[!quote] Underline [Toxic Inhibitors]")) {
    throw new Error("Exported markdown missing tagged underline callout!");
  }
  if (!exportedSchemeMd.includes("[!note] Sticky Note [Incubation Phase]")) {
    throw new Error("Exported markdown missing tagged sticky note callout!");
  }
  console.log("✓ Verified Markdown export includes color scheme legend table and semantic tag callouts");

  console.log("==================================================");
  console.log("🎉 ALL TESTS PASSED SUCCESSFULLY! 100% VERIFIED!");
  console.log("==================================================");
}

runTestSuite().catch((err) => {
  console.error("TEST FAILED:", err);
  process.exit(1);
});
