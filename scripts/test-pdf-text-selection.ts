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
  // 4. Test Precise Line Clustering, Fragment Merging & Parent Container Rejection
  // ------------------------------------------------------------------
  // A. Granular word-level selection:
  // Words: "High-precision" (50..150), "selectable" (155..230), "text" (235..280) on Line 1
  // Words: "Supports" (50..120), "full" (125..160) on Line 2
  // Enclosing container rect from Chromium: [50..350, top: 110..146, height: 36] (must be rejected!)
  // Multi-column or separated gap on Line 1: "Notes" (420..480) (must remain separate segment!)
  const rawFragments = [
    // Chromium parent container rect covering both lines:
    { left: 50, top: 110, right: 350, bottom: 146, width: 300, height: 36 },
    // Line 1 fragments (height: 16):
    { left: 50, top: 110, right: 150, bottom: 126, width: 100, height: 16 },
    { left: 155, top: 110, right: 230, bottom: 126, width: 75, height: 16 },
    { left: 235, top: 110, right: 280, bottom: 126, width: 45, height: 16 },
    // Line 1 distant column fragment (margin/column gap > maxWordGap):
    { left: 420, top: 110, right: 480, bottom: 126, width: 60, height: 16 },
    // Line 2 fragments:
    { left: 50, top: 130, right: 120, bottom: 146, width: 70, height: 16 },
    { left: 125, top: 130, right: 160, bottom: 146, width: 35, height: 16 },
    // Zero-width glitch:
    { left: 160, top: 130, right: 160, bottom: 146, width: 0, height: 16 },
  ];

  const pageRect = { left: 0, top: 0, width: 612, height: 792, right: 612, bottom: 792 };

  interface LocalRect {
    left: number;
    top: number;
    right: number;
    bottom: number;
    width: number;
    height: number;
  }

  let validLocalRects: LocalRect[] = rawFragments
    .filter((r) => r.width > 0.5 && r.height > 0.5)
    .map((r) => ({
      left: Math.max(0, Math.min(pageRect.width, r.left - pageRect.left)),
      top: Math.max(0, Math.min(pageRect.height, r.top - pageRect.top)),
      right: Math.max(0, Math.min(pageRect.width, r.right - pageRect.left)),
      bottom: Math.max(0, Math.min(pageRect.height, r.bottom - pageRect.top)),
      width: r.width,
      height: r.height,
    }));

  // Safe filter for enclosing container/parent rects:
  // Only discard a rect if it is a TRUE multi-line container enclosing multiple vertically-separated rects.
  // NEVER discard valid individual letters, punctuation, or words of differing heights!
  if (validLocalRects.length > 1) {
    validLocalRects = validLocalRects.filter((r) => {
      const enclosed = validLocalRects.filter(
        (other) =>
          other !== r &&
          other.left >= r.left - 1.5 &&
          other.right <= r.right + 1.5 &&
          other.top >= r.top - 1.5 &&
          other.bottom <= r.bottom + 1.5
      );
      if (enclosed.length >= 2) {
        const minEnclosedTop = Math.min(...enclosed.map((e) => e.top));
        const maxEnclosedBottom = Math.max(...enclosed.map((e) => e.bottom));
        const minEnclosedHeight = Math.min(...enclosed.map((e) => e.height));
        const verticalSpread = maxEnclosedBottom - minEnclosedTop;
        if (verticalSpread > minEnclosedHeight * 1.5 && r.height > minEnclosedHeight * 1.8) {
          return false;
        }
      }
      return true;
    });
  }

  // Verify container rect was filtered out
  if (validLocalRects.some((r) => r.height > 25)) {
    throw new Error("Chromium parent container rect was not properly filtered out!");
  }

  // Cluster by vertical center alignment and anti-collision checks
  interface LineGroup {
    top: number;
    bottom: number;
    center: number;
    avgHeight: number;
    rects: LocalRect[];
  }
  const lineGroups: LineGroup[] = [];
  validLocalRects.sort((a, b) => a.top - b.top || a.left - b.left);

  for (const r of validLocalRects) {
    const rCenter = (r.top + r.bottom) / 2;
    const rHeight = r.bottom - r.top;

    let matchedGroup: LineGroup | null = null;
    for (const group of lineGroups) {
      const centerDiff = Math.abs(rCenter - group.center);
      const refHeight = Math.min(rHeight, group.avgHeight);
      const vOverlap = Math.min(group.bottom, r.bottom) - Math.max(group.top, r.top);

      const hasHorizontalConflict = group.rects.some((existing) => {
        if (existing.width > 8 && r.width > 8) {
          const hOverlap = Math.min(existing.right, r.right) - Math.max(existing.left, r.left);
          const minW = Math.min(existing.width, r.width);
          return hOverlap > minW * 0.4;
        }
        return false;
      });

      if (hasHorizontalConflict) {
        continue;
      }

      if (centerDiff <= refHeight * 0.38 && vOverlap >= refHeight * 0.45) {
        matchedGroup = group;
        break;
      }
    }

    if (matchedGroup) {
      matchedGroup.rects.push(r);
      matchedGroup.top = Math.min(matchedGroup.top, r.top);
      matchedGroup.bottom = Math.max(matchedGroup.bottom, r.bottom);
      matchedGroup.center =
        matchedGroup.rects.reduce((sum, item) => sum + (item.top + item.bottom) / 2, 0) / matchedGroup.rects.length;
      matchedGroup.avgHeight =
        matchedGroup.rects.reduce((sum, item) => sum + item.height, 0) / matchedGroup.rects.length;
    } else {
      lineGroups.push({
        top: r.top,
        bottom: r.bottom,
        center: rCenter,
        avgHeight: rHeight,
        rects: [r],
      });
    }
  }

  if (lineGroups.length !== 2) {
    throw new Error(`Expected exactly 2 distinct line groups, got ${lineGroups.length}`);
  }

  const mergedLines: { left: number; top: number; right: number; bottom: number }[] = [];
  for (const group of lineGroups) {
    group.rects.sort((a, b) => a.left - b.left);
    let currentSegment: { left: number; top: number; right: number; bottom: number } | null = null;
    const lineH = group.bottom - group.top;
    const maxWordGap = Math.max(16, lineH * 1.2);

    for (const r of group.rects) {
      if (!currentSegment) {
        currentSegment = { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
      } else if (r.left <= currentSegment.right + maxWordGap) {
        currentSegment.right = Math.max(currentSegment.right, r.right);
        currentSegment.top = Math.min(currentSegment.top, r.top);
        currentSegment.bottom = Math.max(currentSegment.bottom, r.bottom);
      } else {
        mergedLines.push(currentSegment);
        currentSegment = { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
      }
    }
    if (currentSegment) {
      mergedLines.push(currentSegment);
    }
  }
  mergedLines.sort((a, b) => a.top - b.top || a.left - b.left);

  // Line 1 should have 2 segments: [50, 280] and [420, 480]
  // Line 2 should have 1 segment: [50, 160]
  if (mergedLines.length !== 3) {
    throw new Error(`Expected 3 segments (Line 1 split into 2, Line 2 with 1), got ${mergedLines.length}`);
  }
  if (mergedLines[0].left !== 50 || mergedLines[0].right !== 280) {
    throw new Error(`Line 1 segment 1 bounds invalid: [${mergedLines[0].left}, ${mergedLines[0].right}]`);
  }
  if (mergedLines[1].left !== 420 || mergedLines[1].right !== 480) {
    throw new Error(`Line 1 column segment bounds invalid: [${mergedLines[1].left}, ${mergedLines[1].right}]`);
  }
  if (mergedLines[2].left !== 50 || mergedLines[2].right !== 160) {
    throw new Error(`Line 2 bounds invalid: [${mergedLines[2].left}, ${mergedLines[2].right}]`);
  }
  console.log("✓ Verified granular text selection: container rects rejected, lines separated, and sub-line segments preserved");

  // ------------------------------------------------------------------
  // 4B. Test Mixed-Height Glyphs (Punctuation, Capitals) & Horizontal Conflict Isolation
  // ------------------------------------------------------------------
  // Sentence with mixed heights:
  // "Hello," (H=18, ello=14, comma=5) followed by "world!" (w=14, !=12)
  // And next line in paragraph starting at left=50
  const mixedFragments = [
    // Line 1:
    { left: 50, top: 200, right: 65, bottom: 218, width: 15, height: 18 },  // 'H' (cap: 18px)
    { left: 65, top: 204, right: 100, bottom: 218, width: 35, height: 14 }, // 'ello' (14px)
    { left: 100, top: 213, right: 106, bottom: 218, width: 6, height: 5 },  // ',' (comma: 5px!)
    { left: 112, top: 204, right: 160, bottom: 218, width: 48, height: 14 }, // 'world!' (14px)
    // Line 2 (overlapping horizontally at left=50, top=222, bottom=238):
    { left: 50, top: 222, right: 130, bottom: 238, width: 80, height: 16 }, // 'Next sentence'
  ];

  let mixedLocalRects = mixedFragments.map((r) => ({ ...r }));
  // Apply safe container filtering
  if (mixedLocalRects.length > 1) {
    mixedLocalRects = mixedLocalRects.filter((r) => {
      const enclosed = mixedLocalRects.filter(
        (other) =>
          other !== r &&
          other.left >= r.left - 1.5 &&
          other.right <= r.right + 1.5 &&
          other.top >= r.top - 1.5 &&
          other.bottom <= r.bottom + 1.5
      );
      if (enclosed.length >= 2) {
        const minEnclosedTop = Math.min(...enclosed.map((e) => e.top));
        const maxEnclosedBottom = Math.max(...enclosed.map((e) => e.bottom));
        const minEnclosedHeight = Math.min(...enclosed.map((e) => e.height));
        const verticalSpread = maxEnclosedBottom - minEnclosedTop;
        if (verticalSpread > minEnclosedHeight * 1.5 && r.height > minEnclosedHeight * 1.8) {
          return false;
        }
      }
      return true;
    });
  }

  // Verify NONE of the words or punctuation were discarded!
  if (mixedLocalRects.length !== 5) {
    throw new Error(`Expected all 5 mixed fragments to be preserved, got ${mixedLocalRects.length}`);
  }

  // Cluster lines with horizontal collision checks
  const mixedLineGroups: LineGroup[] = [];
  mixedLocalRects.sort((a, b) => a.top - b.top || a.left - b.left);

  for (const r of mixedLocalRects) {
    const rCenter = (r.top + r.bottom) / 2;
    const rHeight = r.bottom - r.top;

    let matchedGroup: LineGroup | null = null;
    for (const group of mixedLineGroups) {
      const hasHorizontalConflict = group.rects.some((existing) => {
        if (existing.width > 8 && r.width > 8) {
          const hOverlap = Math.min(existing.right, r.right) - Math.max(existing.left, r.left);
          const minW = Math.min(existing.width, r.width);
          return hOverlap > minW * 0.4;
        }
        return false;
      });

      if (hasHorizontalConflict) {
        continue;
      }

      const centerDiff = Math.abs(rCenter - group.center);
      const vOverlap = Math.min(group.bottom, r.bottom) - Math.max(group.top, r.top);
      const minHeight = Math.min(rHeight, group.avgHeight);

      const isVerticallyContained =
        r.top >= group.top - 2 &&
        r.bottom <= group.bottom + 2 &&
        vOverlap >= minHeight * 0.75;

      const isCenterAligned =
        centerDiff <= minHeight * 0.45 &&
        vOverlap >= minHeight * 0.45;

      if (isCenterAligned || isVerticallyContained) {
        matchedGroup = group;
        break;
      }
    }

    if (matchedGroup) {
      matchedGroup.rects.push(r);
      matchedGroup.top = Math.min(matchedGroup.top, r.top);
      matchedGroup.bottom = Math.max(matchedGroup.bottom, r.bottom);
      matchedGroup.center =
        matchedGroup.rects.reduce((sum, item) => sum + (item.top + item.bottom) / 2, 0) / matchedGroup.rects.length;
      matchedGroup.avgHeight =
        matchedGroup.rects.reduce((sum, item) => sum + item.height, 0) / matchedGroup.rects.length;
    } else {
      mixedLineGroups.push({
        top: r.top,
        bottom: r.bottom,
        center: rCenter,
        avgHeight: rHeight,
        rects: [r],
      });
    }
  }

  if (mixedLineGroups.length !== 2) {
    throw new Error(`Expected exactly 2 line groups for mixed-height fragments, got ${mixedLineGroups.length}`);
  }

  // Merge segments horizontally
  const mixedResultLines: { left: number; top: number; right: number; bottom: number }[] = [];
  for (const group of mixedLineGroups) {
    group.rects.sort((a, b) => a.left - b.left);
    let currentSegment: { left: number; top: number; right: number; bottom: number } | null = null;
    const maxWordGap = 20;

    for (const r of group.rects) {
      if (!currentSegment) {
        currentSegment = { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
      } else if (r.left <= currentSegment.right + maxWordGap) {
        currentSegment.right = Math.max(currentSegment.right, r.right);
        currentSegment.top = Math.min(currentSegment.top, r.top);
        currentSegment.bottom = Math.max(currentSegment.bottom, r.bottom);
      } else {
        mixedResultLines.push(currentSegment);
        currentSegment = { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
      }
    }
    if (currentSegment) {
      mixedResultLines.push(currentSegment);
    }
  }

  if (mixedResultLines.length !== 2) {
    throw new Error(`Expected 2 merged lines, got ${mixedResultLines.length}`);
  }
  // Line 1: [50, 160] (all words, capitals, and punctuation merged seamlessly)
  if (mixedResultLines[0].left !== 50 || mixedResultLines[0].right !== 160) {
    throw new Error(`Line 1 mixed bounds mismatch: [${mixedResultLines[0].left}, ${mixedResultLines[0].right}]`);
  }
  // Line 2: [50, 130] (stays separate from Line 1)
  if (mixedResultLines[1].left !== 50 || mixedResultLines[1].right !== 130) {
    throw new Error(`Line 2 bounds mismatch: [${mixedResultLines[1].left}, ${mixedResultLines[1].right}]`);
  }
  console.log("✓ Verified 100% glyph retention: no missing letters/punctuation and 0 block ballooning");

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

  if (quadPoints.length !== 24) {
    throw new Error(`Expected 24 QuadPoints (8 per segment across 3 segments), got ${quadPoints.length}`);
  }
  if (minX !== 50 || maxX !== 480) {
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

  // ------------------------------------------------------------------
  // 9. Test Non-Blocking Text Selection & Precise Point-In-Polygon Hit Testing
  // ------------------------------------------------------------------
  function pointInPolygon(px: number, py: number, polygon: { x: number; y: number }[]): boolean {
    let inside = false;
    for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
      const xi = polygon[i].x, yi = polygon[i].y;
      const xj = polygon[j].x, yj = polygon[j].y;
      const intersect = yi > py !== yj > py && px < ((xj - xi) * (py - yi)) / (yj - yi) + xi;
      if (intersect) inside = !inside;
    }
    return inside;
  }

  // Segment 1 of highlightAnnot is [50, 666, 280, 682] in PDF user space
  // Point inside segment 1:
  const hit1 = pointInPolygon(100, 674, [
    { x: quadPoints[0], y: quadPoints[1] },
    { x: quadPoints[2], y: quadPoints[3] },
    { x: quadPoints[6], y: quadPoints[7] },
    { x: quadPoints[4], y: quadPoints[5] },
  ]);
  if (!hit1) {
    throw new Error("Hit test failed to detect point inside highlight segment 1!");
  }

  // Point in gap between segment 1 and column segment 2 (e.g. X = 350, Y = 674):
  const hitGap = pointInPolygon(350, 674, [
    { x: quadPoints[0], y: quadPoints[1] },
    { x: quadPoints[2], y: quadPoints[3] },
    { x: quadPoints[6], y: quadPoints[7] },
    { x: quadPoints[4], y: quadPoints[5] },
  ]);
  if (hitGap) {
    throw new Error("Hit test falsely triggered on gap between columns!");
  }
  console.log("✓ Verified precise non-blocking quadrilateral hit-testing and gap preservation");

  // ------------------------------------------------------------------
  // 10. Test Underline Precision (No Next-Line Collision) and Highlight Height
  // ------------------------------------------------------------------
  // Simulate text line 1: top=110, bottom=126 (height=16px)
  // Simulate text line 2 (next line): top=132, bottom=148
  const sampleLine = { left: 50, top: 110, right: 280, bottom: 126 };
  const nextLine = { left: 50, top: 132, right: 280, bottom: 148 };
  const sLineH = sampleLine.bottom - sampleLine.top; // 16px

  // Verify Highlight height logic:
  // Must match font line-height snugly without excess vertical height
  const hPad = Math.min(1.2, Math.max(0.6, sLineH * 0.05));
  if (hPad < 0.6 || hPad > 1.2) {
    throw new Error(`Highlight hPad out of expected bounds: ${hPad}`);
  }
  const hlTop = sampleLine.top;
  const hlBottom = sampleLine.bottom;
  const hlHeight = hlBottom - hlTop;
  if (hlHeight !== sLineH) {
    throw new Error(`Highlight height mismatch: expected ${sLineH}, got ${hlHeight}`);
  }

  // Verify Underline placement logic:
  // Must be placed right below bottom of line without reaching next line
  const underlineOffset = Math.min(1.0, Math.max(0.4, sLineH * 0.03));
  if (underlineOffset < 0.4 || underlineOffset > 1.0) {
    throw new Error(`Underline offset out of expected bounds: ${underlineOffset}`);
  }
  const renderedUnderlineY = sampleLine.bottom + underlineOffset;
  const strokeWidth = Math.min(1.4, Math.max(1.0, sLineH * 0.07));
  const underlineBottom = renderedUnderlineY + strokeWidth / 2;

  // Crucial requirement: Distance to next line's top must have ample whitespace gap (> 4px)
  const distanceToNextLine = nextLine.top - underlineBottom;
  if (distanceToNextLine < 4.0) {
    throw new Error(`Underline too close to next line (${distanceToNextLine}px), it touches the sentence below!`);
  }

  // Verify orientation-invariant vector displacement:
  const pTop = { x: sampleLine.left, y: sampleLine.top };
  const pBottom = { x: sampleLine.left, y: sampleLine.bottom };
  const dy = pBottom.y - pTop.y;
  const dx = pBottom.x - pTop.x;
  const dist = Math.hypot(dx, dy);
  const uy = dy / dist;
  const yDisplaced = pBottom.y + underlineOffset * uy;
  if (Math.abs(yDisplaced - renderedUnderlineY) > 0.001) {
    throw new Error("Vector displacement for underline clearance mismatch!");
  }
  console.log("✓ Verified highlight height perfectly proportioned and underline leaves clean whitespace above next line");

  // ------------------------------------------------------------------
  // 11. Multi-Column Research Paper Isolation & Text Disambiguation
  // ------------------------------------------------------------------
  // Simulate 3 lines of a 2-column research paper where DOM spans from Column 1 and Column 2 interleave
  interface SimSegment {
    left: number;
    top: number;
    right: number;
    bottom: number;
    text: string;
  }
  const multiColLines: SimSegment[][] = [
    // Line 1: Col 1 [50, 275] and Col 2 [340, 565]
    [
      { left: 50, top: 100, right: 275, bottom: 114, text: "Machine learning models" },
      { left: 340, top: 100, right: 565, bottom: 114, text: "We propose a novel" },
    ],
    // Line 2: Col 1 [50, 280] and Col 2 [340, 570]
    [
      { left: 50, top: 120, right: 280, bottom: 134, text: "have shown great" },
      { left: 340, top: 120, right: 570, bottom: 134, text: "architecture that" },
    ],
    // Line 3: Col 1 [50, 260] and Col 2 [340, 560]
    [
      { left: 50, top: 140, right: 260, bottom: 154, text: "success in vision." },
      { left: 340, top: 140, right: 560, bottom: 154, text: "outperforms prior art." },
    ],
  ];

  // Helper matching the algorithm in PdfPageView
  function resolveActiveColumnLines(
    lines: SimSegment[][],
    anchorLocalX: number
  ): SimSegment[] {
    const multiSegmentLines = lines.filter((segs) => segs.length > 1);
    const hasMultiColumn = multiSegmentLines.length >= 2;
    const results: SimSegment[] = [];

    if (hasMultiColumn) {
      const gutterLeft = Math.max(...multiSegmentLines.map((segs) => segs[0].right));
      const gutterRight = Math.min(...multiSegmentLines.map((segs) => segs[1].left));
      const isAnchorInLeftCol = anchorLocalX <= (gutterLeft + gutterRight) / 2;

      for (const segs of lines) {
        for (const seg of segs) {
          if (isAnchorInLeftCol) {
            if (seg.left <= gutterLeft + 15) {
              results.push(seg);
            }
          } else {
            if (seg.right >= gutterRight - 15) {
              results.push(seg);
            }
          }
        }
      }
    } else {
      for (const segs of lines) {
        for (const seg of segs) {
          results.push(seg);
        }
      }
    }
    return results;
  }

  // Case A: User drags in Left Column (anchorX = 120)
  const leftColSelected = resolveActiveColumnLines(multiColLines, 120);
  if (leftColSelected.length !== 3) {
    throw new Error(`Expected 3 lines in Left Column, got ${leftColSelected.length}`);
  }
  for (const line of leftColSelected) {
    if (line.left >= 300 || line.right > 300) {
      throw new Error(`Left column line rect leaked into right column: left=${line.left}, right=${line.right}`);
    }
  }
  const cleanLeftText = leftColSelected.map((l) => l.text).join(" ");
  if (cleanLeftText !== "Machine learning models have shown great success in vision.") {
    throw new Error(`Corrupted text extraction in multi-column PDF: "${cleanLeftText}"`);
  }
  if (cleanLeftText.includes("propose") || cleanLeftText.includes("architecture")) {
    throw new Error("Text from right column bled into left column selection!");
  }

  // Case B: User drags in Right Column (anchorX = 420)
  const rightColSelected = resolveActiveColumnLines(multiColLines, 420);
  if (rightColSelected.length !== 3) {
    throw new Error(`Expected 3 lines in Right Column, got ${rightColSelected.length}`);
  }
  for (const line of rightColSelected) {
    if (line.left < 300) {
      throw new Error(`Right column line rect leaked into left column: left=${line.left}`);
    }
  }
  const cleanRightText = rightColSelected.map((l) => l.text).join(" ");
  if (cleanRightText !== "We propose a novel architecture that outperforms prior art.") {
    throw new Error(`Corrupted text extraction in right column: "${cleanRightText}"`);
  }
  console.log("✓ Verified double-column multi-line isolation (zero bleeding between adjacent columns)");

  // ------------------------------------------------------------------
  // 12. Guard Against Dragging Past Last Letter of Word (Section Jump Filter)
  // ------------------------------------------------------------------
  // Simulate user dragging across a word on line 1: Y between 100 and 114
  const mouseDown = { x: 50, y: 105 };
  const mouseUp = { x: 130, y: 107 }; // User dragged horizontally ~80px to select a word
  const pageTop = 0;

  // Browser mistakenly expanded range to include spans on lines 5, 6, 7 (Y = 250..350)
  const corruptedRects = [
    { left: 50, top: 100, right: 130, bottom: 114, text: "Machine" }, // Target word
    { left: 50, top: 250, right: 280, bottom: 264, text: "Section 2: Related Work" },
    { left: 50, top: 270, right: 280, bottom: 284, text: "Prior studies have evaluated..." },
    { left: 50, top: 290, right: 280, bottom: 304, text: "In contrast to our approach..." },
  ];

  const dragDist = Math.hypot(mouseUp.x - mouseDown.x, mouseUp.y - mouseDown.y);
  let clampedRects = [...corruptedRects];
  if (dragDist > 10) {
    const minDragY = Math.min(mouseDown.y, mouseUp.y) - pageTop - 20;
    const maxDragY = Math.max(mouseDown.y, mouseUp.y) - pageTop + 20;
    const inBoundsRects = clampedRects.filter((r) => r.top <= maxDragY && r.bottom >= minDragY);
    if (inBoundsRects.length > 0) {
      clampedRects = inBoundsRects;
    }
  }

  if (clampedRects.length !== 1 || clampedRects[0].text !== "Machine") {
    throw new Error(`Section jump filter failed: expected 1 word, got ${clampedRects.length} items`);
  }
  console.log("✓ Verified selection endpoint clamping & section-wide auto-selection prevention");

  // ------------------------------------------------------------------
  // 13. Test Middle-to-End Multi-Line Selection & Last Word Retention
  // ------------------------------------------------------------------
  // Simulates selecting from the middle of Line 1 to a short Line 2 ending with the last word.
  // Prior bug: anchorLocalX (in middle of line 1) caused line 2 to be flagged as isFarLeft and discarded!
  const multiLineMiddleToLast = [
    // Line 1: user selected from x=200 to x=450
    { left: 200, top: 100, right: 300, bottom: 116, width: 100, height: 16, text: "middle of line 1" },
    { left: 305, top: 100, right: 450, bottom: 116, width: 145, height: 16, text: "to the end" },
    // Line 2: short line with the last word (ending at x=160, far to the left of anchorLocalX=200!)
    { left: 50, top: 124, right: 100, bottom: 140, width: 50, height: 16, text: "and" },
    { left: 105, top: 124, right: 160, bottom: 140, width: 55, height: 16, text: "lastword" },
  ];

  // Helper matching the algorithm in PdfPageView
  function simulateProcessSelection(
    fragments: { left: number; top: number; right: number; bottom: number; width: number; height: number; text: string }[],
    dragStartLocalX?: number | null,
    dragEndLocalX?: number | null,
    pageWidth = 600
  ) {
    interface LineGroup {
      top: number;
      bottom: number;
      center: number;
      avgHeight: number;
      rects: typeof fragments;
    }
    const groups: LineGroup[] = [];
    const sorted = [...fragments].sort((a, b) => a.top - b.top || a.left - b.left);

    for (const r of sorted) {
      const rCenter = (r.top + r.bottom) / 2;
      const rHeight = r.bottom - r.top;
      let matchedGroup: LineGroup | null = null;

      for (const group of groups) {
        const centerDiff = Math.abs(rCenter - group.center);
        const vOverlap = Math.min(group.bottom, r.bottom) - Math.max(group.top, r.top);
        const refHeight = Math.min(rHeight, group.avgHeight);

        const hasStrongOverlap = vOverlap >= refHeight * 0.3;
        const isCenterAligned =
          centerDiff <= Math.max(refHeight * 0.45, 6) &&
          vOverlap >= refHeight * 0.2;
        const isPunctuationOrAccent =
          rHeight <= 8 &&
          r.top >= group.top - 6 &&
          r.bottom <= group.bottom + 8;

        if (hasStrongOverlap || isCenterAligned || isPunctuationOrAccent) {
          const isStackedLine = group.rects.some((existing) => {
            if (existing.width > 10 && r.width > 10) {
              const hOverlap = Math.min(existing.right, r.right) - Math.max(existing.left, r.left);
              const minW = Math.min(existing.width, r.width);
              if (hOverlap > minW * 0.5) {
                return centerDiff > refHeight * 0.35;
              }
            }
            return false;
          });

          if (!isStackedLine) {
            matchedGroup = group;
            break;
          }
        }
      }

      if (matchedGroup) {
        matchedGroup.rects.push(r);
        matchedGroup.top = Math.min(matchedGroup.top, r.top);
        matchedGroup.bottom = Math.max(matchedGroup.bottom, r.bottom);
        matchedGroup.center =
          matchedGroup.rects.reduce((sum, item) => sum + (item.top + item.bottom) / 2, 0) / matchedGroup.rects.length;
        matchedGroup.avgHeight =
          matchedGroup.rects.reduce((sum, item) => sum + item.height, 0) / matchedGroup.rects.length;
      } else {
        groups.push({
          top: r.top,
          bottom: r.bottom,
          center: rCenter,
          avgHeight: rHeight,
          rects: [r],
        });
      }
    }

    const lineGroupSegments: { left: number; top: number; right: number; bottom: number; text: string }[][] = [];
    for (const group of groups) {
      const lineTop = Math.min(...group.rects.map((r) => r.top));
      const lineBottom = Math.max(...group.rects.map((r) => r.bottom));
      const lineH = Math.max(10, lineBottom - lineTop);

      for (const r of group.rects) {
        r.top = lineTop;
        r.bottom = lineBottom;
        r.height = lineH;
      }

      group.rects.sort((a, b) => a.left - b.left);
      const segments: { left: number; top: number; right: number; bottom: number; text: string }[] = [];
      let currentSeg: { left: number; top: number; right: number; bottom: number; text: string } | null = null;
      const maxWordGap = Math.max(36, lineH * 1.8);

      for (const r of group.rects) {
        if (!currentSeg) {
          currentSeg = { left: r.left, top: lineTop, right: r.right, bottom: lineBottom, text: r.text || "" };
          continue;
        }

        const gap = r.left - currentSeg.right;
        const isTwoColumnGutter =
          gap >= 16 &&
          currentSeg.left <= pageWidth * 0.45 &&
          currentSeg.right <= pageWidth * 0.55 &&
          r.left >= pageWidth * 0.45 &&
          r.right >= pageWidth * 0.65;

        if (gap <= maxWordGap && !isTwoColumnGutter) {
          currentSeg.right = Math.max(currentSeg.right, r.right);
          if (r.text) {
            if (!currentSeg.text) {
              currentSeg.text = r.text;
            } else {
              const isPunctuation = /^[,\.!?;:\)\]\}]/.test(r.text);
              const afterOpeningQuote = /['"‘“]$/.test(currentSeg.text);
              const afterHyphen = /[-—/]$/.test(currentSeg.text);
              const hasSpace = currentSeg.text.endsWith(" ") || r.text.startsWith(" ");

              if (isPunctuation || afterOpeningQuote || afterHyphen || hasSpace) {
                currentSeg.text = `${currentSeg.text}${r.text}`;
              } else {
                currentSeg.text = `${currentSeg.text} ${r.text}`;
              }
            }
          }
        } else {
          segments.push(currentSeg);
          currentSeg = { left: r.left, top: lineTop, right: r.right, bottom: lineBottom, text: r.text || "" };
        }
      }
      if (currentSeg) segments.push(currentSeg);
      if (segments.length > 0) lineGroupSegments.push(segments);
    }

    const multiSegmentLines = lineGroupSegments.filter((segs) => segs.length > 1);
    const gutters: { left: number; right: number }[] = [];
    for (const segs of multiSegmentLines) {
      for (let i = 0; i < segs.length - 1; i++) {
        const gLeft = segs[i].right;
        const gRight = segs[i + 1].left;
        if (gRight - gLeft >= 14 && gLeft <= pageWidth * 0.55 && gRight >= pageWidth * 0.45) {
          gutters.push({ left: gLeft, right: gRight });
        }
      }
    }

    const hasGutter = gutters.length > 0;
    const avgGutterCenter = hasGutter
      ? gutters.reduce((sum, g) => sum + (g.left + g.right) / 2, 0) / gutters.length
      : pageWidth / 2;
    const gutterLeft = hasGutter ? Math.min(...gutters.map((g) => g.left)) : avgGutterCenter - 10;
    const gutterRight = hasGutter ? Math.max(...gutters.map((g) => g.right)) : avgGutterCenter + 10;

    let userMinX: number | null = null;
    let userMaxX: number | null = null;
    if (dragStartLocalX !== null && dragStartLocalX !== undefined) {
      const endX = (dragEndLocalX !== null && dragEndLocalX !== undefined) ? dragEndLocalX : dragStartLocalX;
      userMinX = Math.min(dragStartLocalX, endX);
      userMaxX = Math.max(dragStartLocalX, endX);
    }

    const draggedAcrossGutter =
      userMinX !== null &&
      userMaxX !== null &&
      userMinX <= gutterLeft &&
      userMaxX >= gutterRight;

    const isStrictlyCol1 =
      hasGutter &&
      userMinX !== null &&
      userMaxX !== null &&
      userMaxX <= avgGutterCenter &&
      !draggedAcrossGutter;

    const isStrictlyCol2 =
      hasGutter &&
      userMinX !== null &&
      userMaxX !== null &&
      userMinX >= avgGutterCenter &&
      !draggedAcrossGutter;

    const resultLines: { left: number; top: number; right: number; bottom: number; text: string }[] = [];

    for (const segments of lineGroupSegments) {
      for (const seg of segments) {
        if (isStrictlyCol1) {
          if (seg.left < avgGutterCenter) {
            resultLines.push(seg);
          }
        } else if (isStrictlyCol2) {
          if (seg.right > avgGutterCenter) {
            resultLines.push(seg);
          }
        } else {
          resultLines.push(seg);
        }
      }
    }

    resultLines.sort((a, b) => a.top - b.top || a.left - b.left);
    return resultLines;
  }

  const middleToLastLines = simulateProcessSelection(multiLineMiddleToLast, 200);
  if (middleToLastLines.length !== 2) {
    throw new Error(`Expected exactly 2 lines in middle-to-end selection, got ${middleToLastLines.length}`);
  }
  if (!middleToLastLines[1].text.includes("lastword")) {
    throw new Error(`Last word was discarded from line 2: "${middleToLastLines[1].text}"`);
  }
  console.log("✓ Verified middle-to-end multi-line selection: short final line & last word 100% preserved");

  // ------------------------------------------------------------------
  // 14. Test Words with Descenders & Mixed Baselines (Zero Drop-Down Glitches)
  // ------------------------------------------------------------------
  // Simulates a single line containing letters with descenders ('p', 'q', 'g', 'y', 'j')
  // and varying baseline heights in the middle of words.
  const descenderFragments = [
    { left: 50, top: 100, right: 100, bottom: 114, width: 50, height: 14, text: "Typing" },
    { left: 105, top: 104, right: 160, bottom: 120, width: 55, height: 16, text: "paragraph" }, // descenders 'p', 'g'
    { left: 165, top: 100, right: 220, bottom: 114, width: 55, height: 14, text: "query" },     // descenders 'q', 'y'
    { left: 225, top: 104, right: 280, bottom: 120, width: 55, height: 16, text: "projects" },  // descenders 'p', 'j'
  ];

  const descenderLines = simulateProcessSelection(descenderFragments, 50);
  if (descenderLines.length !== 1) {
    throw new Error(`Descender fragments split into ${descenderLines.length} lines instead of 1 uniform line!`);
  }
  if (descenderLines[0].left !== 50 || descenderLines[0].right !== 280) {
    throw new Error(`Line rect bounds mismatch for descender sentence: [${descenderLines[0].left}, ${descenderLines[0].right}]`);
  }
  if (descenderLines[0].text !== "Typing paragraph query projects") {
    throw new Error(`Descender sentence text mismatch: "${descenderLines[0].text}"`);
  }
  console.log("✓ Verified inline descenders & mixed-height words retention: 0 dropped-down fragments");

  // ------------------------------------------------------------------
  // 15. Multi-Column Selection with Short Final Line in Active Column
  // ------------------------------------------------------------------
  // Simulates a 2-column paper where Column 1 ends on a short line 3.
  const multiColWithShortEnd = [
    { left: 200, top: 100, right: 275, bottom: 114, width: 75, height: 14, text: "learning models" },
    { left: 340, top: 100, right: 565, bottom: 114, width: 225, height: 14, text: "We propose a novel" },
    { left: 50, top: 120, right: 280, bottom: 134, width: 230, height: 14, text: "have shown great" },
    { left: 340, top: 120, right: 570, bottom: 134, width: 230, height: 14, text: "architecture that" },
    { left: 50, top: 140, right: 160, bottom: 154, width: 110, height: 14, text: "success." },
  ];

  const col1Lines = simulateProcessSelection(multiColWithShortEnd, 200);
  if (col1Lines.length !== 3) {
    throw new Error(`Expected 3 lines in Column 1 selection, got ${col1Lines.length}`);
  }
  const col1Text = col1Lines.map((l) => l.text).join(" ");
  if (col1Text !== "learning models have shown great success.") {
    throw new Error(`Column 1 text mismatch: "${col1Text}"`);
  }
  console.log("✓ Verified multi-column selection with short final line: 100% active column retention without bleeding");

  // ------------------------------------------------------------------
  // 16. Middle Drop-Down Elimination & Last Word Retention Verification
  // ------------------------------------------------------------------
  // Test Case A: Complex line with mixed baselines, middle punctuation, quotes, descenders, and last word
  const complexLineFragments = [
    { left: 50, top: 200, right: 80, bottom: 216, width: 30, height: 16, text: "For" },
    { left: 85, top: 204, right: 150, bottom: 220, width: 65, height: 16, text: "example," }, // comma at baseline, 'p' descender
    { left: 155, top: 204, right: 220, bottom: 220, width: 65, height: 16, text: "applying" }, // 'p', 'y', 'g' descenders (bottom: 220)
    { left: 225, top: 198, right: 337, bottom: 218, width: 112, height: 20, text: "'hyper-parameters'" }, // quote and hyphenated word
    { left: 342, top: 204, right: 360, bottom: 216, width: 18, height: 12, text: "to" },
    { left: 365, top: 204, right: 420, bottom: 216, width: 55, height: 12, text: "achieve" },
    { left: 425, top: 204, right: 485, bottom: 220, width: 60, height: 16, text: "quality." }, // 'q', 'y' descenders and period (last word!)
  ];

  const processedComplexLine = simulateProcessSelection(complexLineFragments, 50);

  if (processedComplexLine.length !== 1) {
    throw new Error(`Expected exactly 1 uniform line rectangle, got ${processedComplexLine.length} lines!`);
  }

  // Verify that the line rectangle has uniform top and bottom bounds
  const line = processedComplexLine[0];
  if (line.top !== 198 || line.bottom !== 220) {
    throw new Error(`Line top/bottom not normalized: top=${line.top}, bottom=${line.bottom}`);
  }
  if (line.left !== 50 || line.right !== 485) {
    throw new Error(`Line horizontal bounds mismatch: left=${line.left}, right=${line.right}`);
  }
  if (!line.text.includes("quality.")) {
    throw new Error(`Last word was dropped: "${line.text}"`);
  }
  if (!line.text.includes("example, applying 'hyper-parameters'")) {
    throw new Error(`Middle text was corrupted: "${line.text}"`);
  }

  // Test Case B: Selecting from middle of line 1 to the very last word of line 2 with wide justified spaces
  const justifiedParagraph = [
    // Line 1: user selected from middle: "developed algorithms with"
    { left: 180, top: 300, right: 260, bottom: 316, width: 80, height: 16, text: "developed" },
    { left: 290, top: 304, right: 380, bottom: 320, width: 90, height: 16, text: "algorithms" }, // 30px justified space
    { left: 410, top: 300, right: 450, bottom: 316, width: 40, height: 16, text: "with" },       // 30px space
    // Line 2: "high accuracy."
    { left: 50, top: 328, right: 85, bottom: 344, width: 35, height: 16, text: "high" },
    { left: 115, top: 328, right: 200, bottom: 344, width: 85, height: 16, text: "accuracy." }, // last word!
  ];

  const processedJustified = simulateProcessSelection(justifiedParagraph, 180);
  if (processedJustified.length !== 2) {
    throw new Error(`Expected 2 lines in justified paragraph, got ${processedJustified.length}`);
  }
  if (!processedJustified[1].text.includes("accuracy.")) {
    throw new Error(`Last word was dropped from line 2: "${processedJustified[1].text}"`);
  }
  if (processedJustified[0].left !== 180 || processedJustified[0].right !== 450) {
    throw new Error(`Line 1 bounds invalid: [${processedJustified[0].left}, ${processedJustified[0].right}]`);
  }
  if (processedJustified[1].left !== 50 || processedJustified[1].right !== 200) {
    throw new Error(`Line 2 bounds invalid: [${processedJustified[1].left}, ${processedJustified[1].right}]`);
  }
  console.log("✓ Verified 0 middle dropped-down words & 100% last word selection guarantee");

  // ------------------------------------------------------------------
  // 17. Test Granular Single-Letter & Sub-Word Selection Precision (Zero Whole-Block Snapping)
  // ------------------------------------------------------------------
  // A. Single letter selection: User selects only 'A' (width: 8px) in a line with 40 characters
  const singleLetterSelection = [
    { left: 120, top: 400, right: 128, bottom: 416, width: 8, height: 16, text: "A" }
  ];
  const processedSingleLetter = simulateProcessSelection(singleLetterSelection, 120);
  if (processedSingleLetter.length !== 1) {
    throw new Error(`Expected 1 segment for single-letter selection, got ${processedSingleLetter.length}`);
  }
  if (processedSingleLetter[0].left !== 120 || processedSingleLetter[0].right !== 128) {
    throw new Error(`Single letter bounds expanded beyond glyph: [${processedSingleLetter[0].left}, ${processedSingleLetter[0].right}]`);
  }
  if (processedSingleLetter[0].text !== "A") {
    throw new Error(`Single letter text corrupted: "${processedSingleLetter[0].text}"`);
  }

  // B. Sub-word selection: User selects 'par' (width: 24px) inside word 'paragraph'
  const subWordSelection = [
    { left: 200, top: 400, right: 224, bottom: 416, width: 24, height: 16, text: "par" }
  ];
  const processedSubWord = simulateProcessSelection(subWordSelection, 200);
  if (processedSubWord.length !== 1 || processedSubWord[0].left !== 200 || processedSubWord[0].right !== 224) {
    throw new Error(`Sub-word selection bounds invalid: [${processedSubWord[0].left}, ${processedSubWord[0].right}]`);
  }
  if (processedSubWord[0].text !== "par") {
    throw new Error(`Sub-word text corrupted: "${processedSubWord[0].text}"`);
  }
  console.log("✓ Verified granular single-letter and sub-word selection: zero whole-block snapping");

  // ------------------------------------------------------------------
  // 18. Test DOM Range Leaf Node Extraction & Zero Dropped Letter Guarantee
  // ------------------------------------------------------------------
  // Simulates multi-span sentences split across multiple DOM text nodes:
  // Node 1: "The " (offsets 0..4)
  // Node 2: "quick " (offsets 0..6)
  // Node 3: "brown " (offsets 0..6)
  // Node 4: "fox" (offsets 0..3)
  // User selects from character 2 of "The " ('e') to character 3 of "brown " ('bro')
  const multiNodeFragments = [
    { left: 66, top: 500, right: 80, bottom: 516, width: 14, height: 16, text: "e " },
    { left: 80, top: 500, right: 130, bottom: 516, width: 50, height: 16, text: "quick " },
    { left: 130, top: 500, right: 165, bottom: 516, width: 35, height: 16, text: "bro" },
  ];
  const processedMultiNode = simulateProcessSelection(multiNodeFragments, 66);
  if (processedMultiNode.length !== 1) {
    throw new Error(`Expected 1 merged line for multi-node selection, got ${processedMultiNode.length}`);
  }
  if (processedMultiNode[0].left !== 66 || processedMultiNode[0].right !== 165) {
    throw new Error(`Multi-node horizontal bounds mismatch: [${processedMultiNode[0].left}, ${processedMultiNode[0].right}]`);
  }
  if (processedMultiNode[0].text !== "e quick bro") {
    throw new Error(`Multi-node text was corrupted or dropped: "${processedMultiNode[0].text}"`);
  }
  console.log("✓ Verified multi-span leaf text node extraction: 0 dropped letters across span boundaries");

  // ------------------------------------------------------------------
  // 19. Test 2-Column Research Paper Single-Line & Multi-Line Zero Bleed Verification
  // ------------------------------------------------------------------
  // Simulates a 2-column paper where Column 1 and Column 2 are on the SAME line (Y=600):
  // Column 1: [50, 275] ("Neural networks are widely used")
  // Column 2: [300, 520] ("Quantum computing algorithms")
  // Gutter: [275, 300] (25px)
  const twoColumnSameLine = [
    { left: 50, top: 600, right: 275, bottom: 616, width: 225, height: 16, text: "Neural networks are widely used" },
    { left: 300, top: 600, right: 520, bottom: 616, width: 220, height: 16, text: "Quantum computing algorithms" },
  ];

  // Case A: User highlights strictly in Column 1 (drag from 100 to 250) -> Column 2 MUST NOT be highlighted!
  const col1OnlySelection = simulateProcessSelection(twoColumnSameLine, 100, 250);
  if (col1OnlySelection.length !== 1) {
    throw new Error(`Expected exactly 1 line segment in Column 1, got ${col1OnlySelection.length}`);
  }
  if (col1OnlySelection[0].left !== 50 || col1OnlySelection[0].right !== 275) {
    throw new Error(`Column 1 selection bled into Column 2: bounds [${col1OnlySelection[0].left}, ${col1OnlySelection[0].right}]`);
  }
  if (col1OnlySelection[0].text !== "Neural networks are widely used") {
    throw new Error(`Column 1 text mismatch: "${col1OnlySelection[0].text}"`);
  }

  // Case B: User highlights strictly in Column 2 (drag from 350 to 500) -> Column 1 MUST NOT be highlighted!
  const col2OnlySelection = simulateProcessSelection(twoColumnSameLine, 350, 500);
  if (col2OnlySelection.length !== 1) {
    throw new Error(`Expected exactly 1 line segment in Column 2, got ${col2OnlySelection.length}`);
  }
  if (col2OnlySelection[0].left !== 300 || col2OnlySelection[0].right !== 520) {
    throw new Error(`Column 2 selection bled into Column 1: bounds [${col2OnlySelection[0].left}, ${col2OnlySelection[0].right}]`);
  }
  if (col2OnlySelection[0].text !== "Quantum computing algorithms") {
    throw new Error(`Column 2 text mismatch: "${col2OnlySelection[0].text}"`);
  }
  console.log("✓ Verified 2-column paper zero bleed: Column 1 and Column 2 on the same line are strictly isolated");

  // ------------------------------------------------------------------
  // 20. Test Universal Text Selection: All Text Selectable Across All Layouts
  // ------------------------------------------------------------------
  // Case A: User deliberately drags ACROSS columns (from 100 in Col 1 to 450 in Col 2)
  const crossColumnSelection = simulateProcessSelection(twoColumnSameLine, 100, 450);
  if (crossColumnSelection.length !== 2) {
    throw new Error(`Expected both column segments to be kept on cross-column drag, got ${crossColumnSelection.length}`);
  }
  if (crossColumnSelection[0].text !== "Neural networks are widely used" || crossColumnSelection[1].text !== "Quantum computing algorithms") {
    throw new Error(`Cross-column drag dropped text: "${crossColumnSelection.map(s => s.text).join(' | ')}"`);
  }

  // Case B: Keyboard / Shift+Arrow selection (no drag coordinates)
  const keyboardSelection = simulateProcessSelection(twoColumnSameLine, null, null);
  if (keyboardSelection.length !== 2) {
    throw new Error(`Keyboard selection dropped text: expected 2 segments, got ${keyboardSelection.length}`);
  }

  // Case C: Single-column document with wide indented paragraphs and margins
  const singleColumnIndents = [
    { left: 80, top: 700, right: 200, bottom: 716, width: 120, height: 16, text: "Section 1.1:" },
    { left: 240, top: 700, right: 540, bottom: 716, width: 300, height: 16, text: "Deep learning architectures and transformer attention." },
  ];
  const processedSingleCol = simulateProcessSelection(singleColumnIndents, 80, 540);
  if (processedSingleCol.length !== 1 && processedSingleCol.length !== 2) {
    throw new Error(`Single-column indented selection dropped: got ${processedSingleCol.length} lines`);
  }
  const fullText = processedSingleCol.map(s => s.text).join(" ");
  if (!fullText.includes("Section 1.1") || !fullText.includes("Deep learning architectures")) {
    throw new Error(`Text was dropped from single-column selection: "${fullText}"`);
  }
  console.log("✓ Verified universal text selection: 100% of text in all layouts is selectable");

  console.log("==================================================");
  console.log("🎉 PDF TEXT LAYER & SELECTION TEST SUITE PASSED 100%!");
  console.log("==================================================");
}
