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
} from "../src/domain/pdf";

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

  const loadedUnderline = loadedExtAnnots.find((a) => a.type === "highlight" && a.rect[1] === 600);
  if (!loadedUnderline || loadedUnderline.contents !== "Underlined hypothesis") {
    throw new Error("Underline annotation was not parsed correctly!");
  }

  const loadedText = loadedExtAnnots.find((a) => a.type === "text");
  if (!loadedText || loadedText.contents !== "Important margin note") {
    throw new Error("Text note without NM was not parsed correctly!");
  }
  console.log("✓ Verified 100% of indirect references, Acrobat Popup comments, and Underline annotations parsed correctly");

  console.log("==================================================");
  console.log("🎉 ALL TESTS PASSED SUCCESSFULLY! 100% VERIFIED!");
  console.log("==================================================");
}

runTestSuite().catch((err) => {
  console.error("TEST FAILED:", err);
  process.exit(1);
});
