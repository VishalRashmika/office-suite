import JSZip from "jszip";
if (typeof DOMParser === "undefined") {
  const xmldom = require("@xmldom/xmldom");
  (globalThis as any).DOMParser = xmldom.DOMParser;
  (globalThis as any).XMLSerializer = xmldom.XMLSerializer;
}

import { XlsxDocument } from "../src/infrastructure/spreadsheet/xlsx-document";
import { cellKey } from "../src/domain/spreadsheet";

export async function runXlsxTestSuite() {
  console.log("=== RUNNING XLSX SPREADSHEET STYLE PERSISTENCE TEST SUITE ===");

  // 1. Create a rich mockup .xlsx file with styles, fonts, fills, borders, number formats, and sheetViews
  const zip = new JSZip();

  // [Content_Types].xml
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
  );

  // _rels/.rels
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  );

  // xl/_rels/workbook.xml.rels
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  // xl/workbook.xml
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Financials" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`
  );

  // xl/sharedStrings.xml
  zip.file(
    "xl/sharedStrings.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Quarter</t></si>
  <si><t>Revenue</t></si>
  <si><t>Growth</t></si>
  <si><t>Q1 Total</t></si>
</sst>`
  );

  // xl/styles.xml with rich fonts, fills, borders, cellXfs
  zip.file(
    "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><color rgb="FF007ACC"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F497D"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right style="thin"><color rgb="FF000000"/></right>
      <top style="thin"><color rgb="FF000000"/></top>
      <bottom style="double"><color rgb="FF000000"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="5">
    <!-- 0: Normal -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <!-- 1: Header (Bold white font, Dark Blue fill, Centered) -->
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 2: Currency with bottom double border -->
    <xf numFmtId="7" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <!-- 3: Percent with yellow highlight -->
    <xf numFmtId="10" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
    <!-- 4: Italic blue font, right aligned -->
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="right"/>
    </xf>
  </cellXfs>
</styleSheet>`
  );

  // xl/worksheets/sheet1.xml with sheetViews, freeze panes, and styled cells
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="16" customWidth="1"/>
    <col min="2" max="2" width="20" customWidth="1"/>
    <col min="3" max="3" width="15" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="28" customHeight="1">
      <c r="A1" s="1" t="s"><v>0</v></c>
      <c r="B1" s="1" t="s"><v>1</v></c>
      <c r="C1" s="1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" s="4" t="s"><v>3</v></c>
      <c r="B2" s="2"><v>1250000</v></c>
      <c r="C2" s="3"><v>0.155</v></c>
    </row>
    <row r="3">
      <!-- Empty styled cell with yellow fill -->
      <c r="C3" s="3"/>
    </row>
  </sheetData>
</worksheet>`
  );

  const initialBytes = await zip.generateAsync({ type: "arraybuffer" });
  console.log(`✓ Generated rich mock .xlsx with styles and frozen pane (${initialBytes.byteLength} bytes)`);

  // --- 2. Test Loading and Style Parsing ---
  const doc = await XlsxDocument.load(initialBytes);
  const data = doc.getWorkbookData();

  if (data.sheets.length !== 1) {
    throw new Error(`Expected 1 sheet, found ${data.sheets.length}`);
  }
  const sheet = data.sheets[0];
  if (sheet.name !== "Financials") {
    throw new Error(`Expected sheet name "Financials", got "${sheet.name}"`);
  }

  // Verify Header Cell A1 (style 1: bold, white text, dark blue background, centered)
  const cellA1 = sheet.cells[cellKey(0, 0)];
  if (!cellA1 || cellA1.value !== "Quarter") {
    throw new Error(`Cell A1 value mismatch. Expected "Quarter", got ${cellA1?.value}`);
  }
  if (!cellA1.style) {
    throw new Error("Cell A1 style was not loaded!");
  }
  if (!cellA1.style.bold) {
    throw new Error("Cell A1 should be bold");
  }
  if (cellA1.style.align !== "center") {
    throw new Error(`Cell A1 align expected "center", got "${cellA1.style.align}"`);
  }
  if (cellA1.style.background?.toLowerCase() !== "#1f497d") {
    throw new Error(`Cell A1 background expected "#1f497d", got "${cellA1.style.background}"`);
  }
  if (cellA1.rawStyleIndex !== 1) {
    throw new Error(`Cell A1 rawStyleIndex expected 1, got ${cellA1.rawStyleIndex}`);
  }
  console.log("✓ Verified Header cell A1 styles parsed (bold, #1f497d fill, center aligned, rawStyleIndex: 1)");

  // Verify Cell B2 (style 2: currency number format, double bottom border)
  const cellB2 = sheet.cells[cellKey(1, 1)];
  if (!cellB2 || cellB2.value !== 1250000) {
    throw new Error(`Cell B2 value expected 1250000, got ${cellB2?.value}`);
  }
  if (cellB2.style?.numberFormat !== "currency") {
    throw new Error(`Cell B2 numberFormat expected "currency", got "${cellB2.style?.numberFormat}"`);
  }
  if (!cellB2.style?.borderBottom || !cellB2.style.borderBottom.includes("double")) {
    throw new Error(`Cell B2 borderBottom expected double, got "${cellB2.style?.borderBottom}"`);
  }
  if (!cellB2.formatted?.includes("$")) {
    throw new Error(`Cell B2 formatted should include currency symbol, got "${cellB2.formatted}"`);
  }
  console.log(`✓ Verified Currency cell B2 formatted correctly: "${cellB2.formatted}" with double border`);

  // Verify Cell C2 (style 3: percent number format, yellow fill)
  const cellC2 = sheet.cells[cellKey(1, 2)];
  if (!cellC2 || cellC2.value !== 0.155) {
    throw new Error(`Cell C2 value expected 0.155, got ${cellC2?.value}`);
  }
  if (cellC2.style?.numberFormat !== "percent") {
    throw new Error(`Cell C2 numberFormat expected "percent", got "${cellC2.style?.numberFormat}"`);
  }
  if (cellC2.style?.background?.toLowerCase() !== "#ffff00") {
    throw new Error(`Cell C2 background expected "#ffff00", got "${cellC2.style?.background}"`);
  }
  console.log(`✓ Verified Percent cell C2 formatted correctly: "${cellC2.formatted}" with yellow fill`);

  // Verify Cell C3 (empty cell with yellow fill)
  const cellC3 = sheet.cells[cellKey(2, 2)];
  if (!cellC3) {
    throw new Error("Empty styled cell C3 should be preserved on load!");
  }
  if (cellC3.style?.background?.toLowerCase() !== "#ffff00") {
    throw new Error(`Cell C3 background expected "#ffff00", got "${cellC3.style?.background}"`);
  }
  console.log("✓ Verified empty styled cell C3 loaded with yellow background");

  // --- 3. Test Modifying Data and Adding New Styles ---
  // A. Edit text in A1 without changing style
  cellA1.value = "Fiscal Quarter";

  // B. Add a brand new cell with custom formatting (e.g. green fill, 14pt, bold)
  const cellB3Key = cellKey(2, 1);
  sheet.cells[cellB3Key] = {
    value: "Audited",
    style: {
      bold: true,
      fontSize: 14,
      color: "#ffffff",
      background: "#27ae60",
      align: "center",
    },
  };

  // --- 4. Test Saving and XML Inspection ---
  const savedBytes = await doc.save(data);
  const reloadedZip = await JSZip.loadAsync(savedBytes);

  const reloadedStylesXml = await reloadedZip.file("xl/styles.xml")?.async("text");
  if (!reloadedStylesXml) {
    throw new Error("xl/styles.xml was missing in saved zip!");
  }

  // Original styles MUST still be present
  if (!reloadedStylesXml.includes("1F497D") || !reloadedStylesXml.includes("FFFF00")) {
    throw new Error("Original fill colors (1F497D, FFFF00) were deleted from styles.xml!");
  }
  // Newly added green fill MUST be present
  if (!reloadedStylesXml.includes("27AE60")) {
    throw new Error("Newly added green fill (27AE60) was not serialized into styles.xml!");
  }
  console.log("✓ Verified xl/styles.xml preserves 100% of original styles and appends new styles");

  const reloadedSheetXml = await reloadedZip.file("xl/worksheets/sheet1.xml")?.async("text");
  if (!reloadedSheetXml) {
    throw new Error("xl/worksheets/sheet1.xml was missing in saved zip!");
  }

  // Frozen pane MUST NOT be deleted
  if (!reloadedSheetXml.includes("sheetViews") || !reloadedSheetXml.includes("state=\"frozen\"")) {
    throw new Error("Workshet structure lost! <sheetViews> or frozen pane was deleted!");
  }
  console.log("✓ Verified <sheetViews> and frozen pane are 100% preserved in worksheet XML");

  // Verify cell style indices in sheetData
  // A1 must have s="1"
  if (!reloadedSheetXml.includes('<c r="A1" s="1"')) {
    throw new Error('Cell A1 should have preserved s="1"');
  }
  // B2 must have s="2"
  if (!reloadedSheetXml.includes('<c r="B2" s="2"')) {
    throw new Error('Cell B2 should have preserved s="2"');
  }
  // Empty styled cell C3 must have s="3"
  if (!reloadedSheetXml.includes('<c r="C3" s="3"/>')) {
    throw new Error('Empty styled cell C3 should be present with s="3"');
  }
  // Newly created cell B3 must have an 's' attribute > 0
  const b3Match = reloadedSheetXml.match(/<c r="B3" s="(\d+)"/);
  if (!b3Match) {
    throw new Error("Newly created cell B3 does not have a style index attribute 's'!");
  }
  const b3StyleIndex = parseInt(b3Match[1], 10);
  if (b3StyleIndex < 5) {
    throw new Error(`Newly created cell B3 should have a new style index >= 5, got ${b3StyleIndex}`);
  }
  console.log(`✓ Verified cell B3 assigned new style index ${b3StyleIndex} with green fill`);

  // --- 5. Full Binary Reload Test ---
  const roundtripDoc = await XlsxDocument.load(savedBytes);
  const roundtripData = roundtripDoc.getWorkbookData();
  const roundtripSheet = roundtripData.sheets[0];

  const roundtripA1 = roundtripSheet.cells[cellKey(0, 0)];
  if (roundtripA1.value !== "Fiscal Quarter" || !roundtripA1.style?.bold) {
    throw new Error("Roundtrip cell A1 value or style lost!");
  }

  const roundtripB3 = roundtripSheet.cells[cellKey(2, 1)];
  if (roundtripB3.value !== "Audited" || roundtripB3.style?.background?.toLowerCase() !== "#27ae60") {
    throw new Error("Roundtrip cell B3 green background style lost!");
  }
  if (roundtripB3.style?.fontSize !== 14 || !roundtripB3.style?.bold) {
    throw new Error("Roundtrip cell B3 font properties lost!");
  }

  console.log("==================================================");
  console.log("🎉 ALL XLSX STYLE PERSISTENCE TESTS PASSED 100%! 🎉");
  console.log("==================================================");
}

// Allow direct CLI execution
if (require.main === module) {
  runXlsxTestSuite().catch((err) => {
    console.error("XLSX Test Failed:", err);
    process.exit(1);
  });
}
