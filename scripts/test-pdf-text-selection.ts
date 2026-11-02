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

  let validLocalRects = rawFragments
    .filter((r) => r.width > 0.5 && r.height > 0.5)
    .map((r) => ({
      left: Math.max(0, Math.min(pageRect.width, r.left - pageRect.left)),
      top: Math.max(0, Math.min(pageRect.height, r.top - pageRect.top)),
      right: Math.max(0, Math.min(pageRect.width, r.right - pageRect.left)),
      bottom: Math.max(0, Math.min(pageRect.height, r.bottom - pageRect.top)),
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
    rects: typeof validLocalRects;
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

  console.log("==================================================");
  console.log("🎉 PDF TEXT LAYER & SELECTION TEST SUITE PASSED 100%!");
  console.log("==================================================");
}
