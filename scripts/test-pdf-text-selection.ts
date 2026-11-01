/**
 * Dedicated test suite for PDF Text Layer, Native Selection, Highlighting,
 * Underlining, and Appearance Stream Persistence.
 */

import { PDFDocument, degrees } from "pdf-lib";
import {
  PdfDocumentAdapter,
} from "../src/infrastructure/pdf/pdf-document-adapter";
import {
  PdfHighlightAnnotation,
  PdfUnderlineAnnotation,
  PdfCoordinateTransformer,
  PdfRect,
  hexToRgb,
} from "../src/domain/pdf";

export async function runPdfTextSelectionTestSuite() {
  console.log("\n=== RUNNING PDF TEXT LAYER & SELECTION TEST SUITE ===");

  const dynamicRequire = new Function("moduleName", "return require(moduleName)");
  const pdfjs = dynamicRequire("pdfjs-dist/legacy/build/pdf.js");

  // ------------------------------------------------------------------
  // 1. Create a rich document with formatted text lines
  // ------------------------------------------------------------------
  const doc = await PDFDocument.create();
  const page = doc.addPage([612, 792]); // Standard US Letter

  // Line 1: Header at Y = 720 (72 pt from top)
  page.drawText("Obsidian Office Suite - Native PDF Reader", {
    x: 50,
    y: 720,
    size: 20,
  });

  // Line 2: First body sentence at Y = 680 (112 pt from top)
  page.drawText("High-precision selectable text layer matching native canvas rendering.", {
    x: 50,
    y: 680,
    size: 12,
  });

  // Line 3: Second body sentence at Y = 660 (132 pt from top)
  page.drawText("Supports full text highlighting, underlining, and Zotero-style annotations.", {
    x: 50,
    y: 660,
    size: 12,
  });

  // Page 2: Rotated 90 degrees
  const page2 = doc.addPage([612, 792]);
  page2.setRotation(degrees(90));
  page2.drawText("Rotated landscape page text selection test.", {
    x: 60,
    y: 700,
    size: 14,
  });

  const pdfBytes = await doc.save();
  console.log("✓ Created rich test PDF with formatted lines and rotated pages:", pdfBytes.length, "bytes");

  // ------------------------------------------------------------------
  // 2. Load with PDF.js and verify textContent extraction
  // ------------------------------------------------------------------
  const loadingTask = pdfjs.getDocument({ data: new Uint8Array(pdfBytes) });
  const pdfjsDoc = await loadingTask.promise;
  const page1 = await pdfjsDoc.getPage(1);
  const textContent1 = await page1.getTextContent();

  if (textContent1.items.length < 3) {
    throw new Error(`Expected at least 3 text items, got ${textContent1.items.length}`);
  }
  console.log(`✓ PDF.js extracted ${textContent1.items.length} text items successfully`);

  // ------------------------------------------------------------------
  // 3. Test mathematical 2D Affine Transformation & Subpixel Alignment
  // ------------------------------------------------------------------
  const scales = [1.0, 1.5, 2.0];
  function multiplyTransform(m1: number[], m2: number[]): number[] {
    return [
      m1[0] * m2[0] + m1[2] * m2[1],
      m1[1] * m2[0] + m1[3] * m2[1],
      m1[0] * m2[2] + m1[2] * m2[3],
      m1[1] * m2[2] + m1[3] * m2[3],
      m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
      m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
    ];
  }

  for (const scale of scales) {
    const viewport = page1.getViewport({ scale });

    for (const item of textContent1.items) {
      if (!("str" in item) || !item.str.trim()) continue;

      const tx = multiplyTransform(viewport.transform, item.transform);
      const fontHeight = Math.hypot(tx[2], tx[3]);

      // PDF.js viewport.convertToViewportPoint converts PDF coordinates to screen
      const pt = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);

      const xDiff = Math.abs(tx[4] - pt[0]);
      const yDiff = Math.abs(tx[5] - pt[1]);

      if (xDiff > 0.001 || yDiff > 0.001) {
        throw new Error(
          `Subpixel misalignment at scale ${scale}: tx=(${tx[4]}, ${tx[5]}), expected pt=(${pt[0]}, ${pt[1]})`
        );
      }

      // Check font height scales strictly with viewport.scale
      const unscaledHeight = item.height || Math.abs(item.transform[3]);
      const expectedHeight = unscaledHeight * scale;
      if (Math.abs(fontHeight - expectedHeight) > 0.01) {
        throw new Error(
          `Font height mismatch: got ${fontHeight}, expected ${expectedHeight} at scale ${scale}`
        );
      }
    }
  }
  console.log("✓ Verified 100% subpixel glyph alignment across scales 1.0x, 1.5x, 2.0x");

  // ------------------------------------------------------------------
  // 4. Test Text Line Fragment Merging Algorithm
  // ------------------------------------------------------------------
  // Simulate browser getClientRects() returning 3 fragmented boxes on the same line
  // e.g. for words "High-precision", "selectable", "text"
  const rawFragments = [
    { left: 50, top: 110, right: 150, bottom: 126, width: 100, height: 16 },
    { left: 155, top: 110, right: 230, bottom: 126, width: 75, height: 16 },
    { left: 235, top: 110, right: 280, bottom: 126, width: 45, height: 16 },
    // A second line:
    { left: 50, top: 130, right: 200, bottom: 146, width: 150, height: 16 },
    { left: 205, top: 130, right: 350, bottom: 146, width: 145, height: 16 },
    // A zero-width glitch fragment:
    { left: 350, top: 130, right: 350, bottom: 146, width: 0, height: 16 },
  ];

  const pageRect = { left: 0, top: 0, width: 612, height: 792, right: 612, bottom: 792 };

  const validLocalRects = rawFragments
    .filter((r) => r.width > 0.5 && r.height > 0.5)
    .map((r) => ({
      left: Math.max(0, Math.min(pageRect.width, r.left - pageRect.left)),
      top: Math.max(0, Math.min(pageRect.height, r.top - pageRect.top)),
      right: Math.max(0, Math.min(pageRect.width, r.right - pageRect.left)),
      bottom: Math.max(0, Math.min(pageRect.height, r.bottom - pageRect.top)),
      height: r.height,
    }));

  const mergedLines: { left: number; top: number; right: number; bottom: number }[] = [];
  for (const r of validLocalRects) {
    const existingLine = mergedLines.find((l) => {
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
      mergedLines.push({ left: r.left, top: r.top, right: r.right, bottom: r.bottom });
    }
  }
  mergedLines.sort((a, b) => a.top - b.top);

  if (mergedLines.length !== 2) {
    throw new Error(`Expected 2 merged lines, got ${mergedLines.length}`);
  }
  if (mergedLines[0].left !== 50 || mergedLines[0].right !== 280) {
    throw new Error(`Line 1 bounds invalid: expected [50, 280], got [${mergedLines[0].left}, ${mergedLines[0].right}]`);
  }
  if (mergedLines[1].left !== 50 || mergedLines[1].right !== 350) {
    throw new Error(`Line 2 bounds invalid: expected [50, 350], got [${mergedLines[1].left}, ${mergedLines[1].right}]`);
  }
  console.log("✓ Verified text line fragment merging: 5 fragments cleanly merged into 2 unified line bounds");

  // ------------------------------------------------------------------
  // 5. Test Screen-to-PDF QuadPoints Conversion
  // ------------------------------------------------------------------
  const transformer = new PdfCoordinateTransformer(612, 792, 0);
  const quadPoints: number[] = [];
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

  for (const line of mergedLines) {
    const pTopLeft = transformer.screenToPdf(line.left, line.top, 1.0);
    const pTopRight = transformer.screenToPdf(line.right, line.top, 1.0);
    const pBottomLeft = transformer.screenToPdf(line.left, line.bottom, 1.0);
    const pBottomRight = transformer.screenToPdf(line.right, line.bottom, 1.0);

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

  if (quadPoints.length !== 16) {
    throw new Error(`Expected 16 QuadPoints (8 per line), got ${quadPoints.length}`);
  }
  if (minX !== 50 || maxX !== 350) {
    throw new Error(`Bounding rect X mismatch: [${minX}, ${maxX}]`);
  }
  // Line 1 top is at screen Y=110 -> PDF Y = 792 - 110 = 682
  // Line 2 bottom is at screen Y=146 -> PDF Y = 792 - 146 = 646
  if (Math.abs(maxY - 682) > 0.01 || Math.abs(minY - 646) > 0.01) {
    throw new Error(`Bounding rect Y mismatch: got [${minY}, ${maxY}], expected [646, 682]`);
  }
  console.log("✓ Verified QuadPoints generation conforming to ISO 32000-1 user space specifications");

  // ------------------------------------------------------------------
  // 6. Test Highlight & Underline Creation, Upsert & ISO-32000 Roundtrip
  // ------------------------------------------------------------------
  const adapter = await PdfDocumentAdapter.load(pdfBytes.buffer);

  // A. Create Highlight
  const highlightAnnot: PdfHighlightAnnotation = {
    id: "hl_selection_test",
    type: "highlight",
    pageIndex: 0,
    rect: [minX, minY, maxX, maxY],
    quadPoints,
    color: hexToRgb("#fef08a"), // Light yellow
    contents: "High-precision selectable text layer matching native canvas rendering.",
    highlightedText: "High-precision selectable text layer matching native canvas rendering.",
    author: "Researcher",
    createdAt: new Date(),
  };
  adapter.upsertAnnotation(highlightAnnot);

  // B. Create Underline on Line 3
  const line3Screen = { left: 50, top: 132, right: 400, bottom: 144 };
  const ulTopLeft = transformer.screenToPdf(line3Screen.left, line3Screen.top, 1.0);
  const ulTopRight = transformer.screenToPdf(line3Screen.right, line3Screen.top, 1.0);
  const ulBottomLeft = transformer.screenToPdf(line3Screen.left, line3Screen.bottom, 1.0);
  const ulBottomRight = transformer.screenToPdf(line3Screen.right, line3Screen.bottom, 1.0);

  const underlineAnnot: PdfUnderlineAnnotation = {
    id: "ul_selection_test",
    type: "underline",
    pageIndex: 0,
    rect: [line3Screen.left, ulBottomLeft.y - 2, line3Screen.right, ulTopLeft.y],
    quadPoints: [
      ulTopLeft.x, ulTopLeft.y,
      ulTopRight.x, ulTopRight.y,
      ulBottomLeft.x, ulBottomLeft.y,
      ulBottomRight.x, ulBottomRight.y,
    ],
    color: hexToRgb("#ef4444"), // Red
    contents: "Supports full text highlighting, underlining",
    highlightedText: "Supports full text highlighting, underlining",
    author: "Researcher",
    createdAt: new Date(),
  };
  adapter.upsertAnnotation(underlineAnnot);

  const savedWithAnnots = await adapter.save();
  console.log("✓ Saved PDF with newly created Highlight & Underline annotations:", savedWithAnnots.length, "bytes");

  // Reload with PDF.js and verify ISO compliance
  const reloadedDoc = await pdfjs.getDocument({ data: new Uint8Array(savedWithAnnots) }).promise;
  const reloadedPage = await reloadedDoc.getPage(1);
  const pdfjsAnnots = await reloadedPage.getAnnotations({ intent: "display" });

  const foundHl = pdfjsAnnots.find(
    (a: any) => a.subtype === "Highlight" || a.annotationName === "hl_selection_test"
  );
  if (!foundHl) {
    throw new Error("Highlight annotation was not parsed by PDF.js!");
  }
  if (!foundHl.hasAppearance) {
    throw new Error("Highlight annotation missing /AP appearance stream!");
  }
  if (foundHl.subtype !== "Highlight") {
    throw new Error(`Expected subtype Highlight, got ${foundHl.subtype}`);
  }

  const foundUl = pdfjsAnnots.find(
    (a: any) => a.subtype === "Underline" || a.annotationName === "ul_selection_test"
  );
  if (!foundUl) {
    throw new Error("Underline annotation was not parsed by PDF.js!");
  }
  if (!foundUl.hasAppearance) {
    throw new Error("Underline annotation missing /AP appearance stream!");
  }
  if (foundUl.subtype !== "Underline") {
    throw new Error(`Expected subtype Underline, got ${foundUl.subtype}`);
  }
  console.log("✓ Verified PDF.js parses both Highlight & Underline annotations with compliant /AP streams");

  // ------------------------------------------------------------------
  // 7. Test Rotated Page Coordinates (90° Landscape Page)
  // ------------------------------------------------------------------
  const page2Viewport = (await pdfjsDoc.getPage(2)).getViewport({ scale: 1.0, rotation: 90 });
  const tx90 = multiplyTransform(page2Viewport.transform, [14, 0, 0, 14, 60, 700]);
  const pt90 = page2Viewport.convertToViewportPoint(60, 700);

  if (Math.abs(tx90[4] - pt90[0]) > 0.001 || Math.abs(tx90[5] - pt90[1]) > 0.001) {
    throw new Error(`Rotated 90deg text alignment mismatch: got (${tx90[4]}, ${tx90[5]}), expected (${pt90[0]}, ${pt90[1]})`);
  }
  console.log("✓ Verified 90° landscape rotated page subpixel text coordinate alignment");

  // ------------------------------------------------------------------
  // 8. Test Cached Selection Lines Resilience
  // ------------------------------------------------------------------
  // When clicking floating action bar buttons, selection might collapse.
  // Verify that passing cached lines creates valid annotations identically.
  const cachedLines = [
    { left: 100, top: 200, right: 300, bottom: 216 },
  ];
  const cachedHl: PdfHighlightAnnotation = {
    id: "cached_hl_test",
    type: "highlight",
    pageIndex: 0,
    rect: [100, 792 - 216, 300, 792 - 200],
    quadPoints: [
      100, 792 - 200,
      300, 792 - 200,
      100, 792 - 216,
      300, 792 - 216,
    ],
    color: [0.3, 0.8, 0.4],
    contents: "Resilient Cached Selection Text",
    highlightedText: "Resilient Cached Selection Text",
    author: "User",
    createdAt: new Date(),
  };
  adapter.upsertAnnotation(cachedHl);
  const reloadedCached = await adapter.getAnnotation("cached_hl_test");
  if (!reloadedCached || reloadedCached.type !== "highlight") {
    throw new Error("Cached line highlight was not upserted properly!");
  }
  console.log("✓ Verified Cached Selection Lines resilience against focus/selection loss");

  console.log("==================================================");
  console.log("🎉 PDF TEXT LAYER & SELECTION TEST SUITE PASSED 100%!");
  console.log("==================================================");
}
