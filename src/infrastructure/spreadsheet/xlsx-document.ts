import JSZip from "jszip";
import {
  WorkbookData,
  SheetData,
  cellKey,
  a1ToCoord,
  evaluateSheetFormulas,
  defaultXlsxService,
} from "../../domain/spreadsheet";

export class XlsxDocument {
  private zip: JSZip;
  private workbookData: WorkbookData;

  private constructor(zip: JSZip, workbookData: WorkbookData) {
    this.zip = zip;
    this.workbookData = workbookData;
  }

  static async load(bytes: ArrayBuffer): Promise<XlsxDocument> {
    const zip = await JSZip.loadAsync(bytes);
    const parser = new DOMParser();

    // 1. Shared Strings Table
    const sharedStrings: string[] = [];
    const sstFile = zip.file("xl/sharedStrings.xml");
    if (sstFile) {
      const sstXml = await sstFile.async("text");
      const sstDoc = parser.parseFromString(sstXml, "application/xml");
      const siNodes = sstDoc.getElementsByTagName("si");
      for (let i = 0; i < siNodes.length; i++) {
        const tNodes = siNodes[i].getElementsByTagName("t");
        let text = "";
        for (let j = 0; j < tNodes.length; j++) {
          text += tNodes[j].textContent || "";
        }
        sharedStrings.push(text);
      }
    }

    // 2. Workbook.xml: Sheets & Relationships
    const wbFile = zip.file("xl/workbook.xml");
    if (!wbFile) {
      throw new Error("Invalid .xlsx: xl/workbook.xml not found.");
    }
    const wbXml = await wbFile.async("text");
    const wbDoc = parser.parseFromString(wbXml, "application/xml");

    // Rel map
    const relsFile = zip.file("xl/_rels/workbook.xml.rels");
    const relMap: Record<string, string> = {};
    if (relsFile) {
      const relsXml = await relsFile.async("text");
      const relsDoc = parser.parseFromString(relsXml, "application/xml");
      const relNodes = relsDoc.getElementsByTagName("Relationship");
      for (let i = 0; i < relNodes.length; i++) {
        const id = relNodes[i].getAttribute("Id");
        const target = relNodes[i].getAttribute("Target");
        if (id && target) relMap[id] = target.replace(/^xl\//, "").replace(/^\//, "");
      }
    }

    const sheetNodes = wbDoc.getElementsByTagName("sheet");
    const sheets: SheetData[] = [];

    for (let i = 0; i < sheetNodes.length; i++) {
      const sNode = sheetNodes[i];
      const name = sNode.getAttribute("name") || `Sheet${i + 1}`;
      const rId = sNode.getAttribute("r:id") || sNode.getAttribute("id") || `rId${i + 1}`;
      let targetPath = relMap[rId] || `worksheets/sheet${i + 1}.xml`;
      if (!targetPath.startsWith("xl/")) targetPath = `xl/${targetPath}`;

      const wsFile = zip.file(targetPath);
      if (!wsFile) continue;

      const wsXml = await wsFile.async("text");
      const wsDoc = parser.parseFromString(wsXml, "application/xml");

      const sheetData: SheetData = {
        id: `sheet_${i + 1}`,
        name,
        cells: {},
        rowCount: 40,
        colCount: 26,
        colWidths: {},
        rowHeights: {},
        merges: [],
      };

      // Col widths
      const colNodes = wsDoc.getElementsByTagName("col");
      for (let c = 0; c < colNodes.length; c++) {
        const col = colNodes[c];
        const min = parseInt(col.getAttribute("min") || "1", 10) - 1;
        const max = parseInt(col.getAttribute("max") || "1", 10) - 1;
        const width = parseFloat(col.getAttribute("width") || "10");
        const pxWidth = Math.round(width * 8.5);
        for (let idx = min; idx <= max; idx++) {
          sheetData.colWidths[idx] = pxWidth;
        }
      }

      // Rows & Cells
      const rowNodes = wsDoc.getElementsByTagName("row");
      let maxRow = 0;
      let maxCol = 0;

      for (let r = 0; r < rowNodes.length; r++) {
        const rowEl = rowNodes[r];
        const rowIdx = parseInt(rowEl.getAttribute("r") || String(r + 1), 10) - 1;
        if (rowIdx > maxRow) maxRow = rowIdx;

        const ht = rowEl.getAttribute("ht");
        if (ht) {
          sheetData.rowHeights[rowIdx] = Math.round(parseFloat(ht) * 1.33);
        }

        const cNodes = rowEl.getElementsByTagName("c");
        for (let j = 0; j < cNodes.length; j++) {
          const cEl = cNodes[j];
          const a1Ref = cEl.getAttribute("r");
          if (!a1Ref) continue;
          const coord = a1ToCoord(a1Ref);
          if (!coord) continue;

          if (coord.col > maxCol) maxCol = coord.col;

          const cellType = cEl.getAttribute("t"); // "s" (shared string), "b" (boolean), "str" (formula string)
          const fEl = cEl.getElementsByTagName("f")[0];
          const vEl = cEl.getElementsByTagName("v")[0];

          let formula: string | undefined;
          let value: any = null;

          if (fEl && fEl.textContent) {
            formula = `=${fEl.textContent.trim()}`;
          }

          if (vEl && vEl.textContent !== null) {
            const rawVal = vEl.textContent;
            if (cellType === "s") {
              const strIdx = parseInt(rawVal, 10);
              value = sharedStrings[strIdx] ?? "";
            } else if (cellType === "b") {
              value = rawVal === "1";
            } else {
              const num = Number(rawVal);
              value = isNaN(num) ? rawVal : num;
            }
          } else {
            const tEl = cEl.getElementsByTagName("t")[0];
            if (tEl && tEl.textContent !== null) {
              value = tEl.textContent;
            }
          }

          sheetData.cells[cellKey(coord.row, coord.col)] = {
            value,
            formula,
            formatted: value !== null && value !== undefined && value !== "" ? String(value) : undefined,
          };
        }
      }

      // Merges
      const mergeNodes = wsDoc.getElementsByTagName("mergeCell");
      for (let m = 0; m < mergeNodes.length; m++) {
        const ref = mergeNodes[m].getAttribute("ref");
        if (ref) {
          const [startA1, endA1] = ref.split(":");
          const start = a1ToCoord(startA1);
          const end = endA1 ? a1ToCoord(endA1) : start;
          if (start && end) {
            sheetData.merges?.push({
              startRow: start.row,
              startCol: start.col,
              endRow: end.row,
              endCol: end.col,
            });
          }
        }
      }

      sheetData.rowCount = Math.max(40, maxRow + 10);
      sheetData.colCount = Math.max(26, maxCol + 5);

      evaluateSheetFormulas(sheetData);
      sheets.push(sheetData);
    }

    if (!sheets.length) {
      sheets.push({
        id: "sheet1",
        name: "Sheet1",
        cells: {},
        rowCount: 40,
        colCount: 26,
        colWidths: {},
        rowHeights: {},
      });
    }

    return new XlsxDocument(zip, {
      sheets,
      activeSheetIndex: 0,
    });
  }

  getWorkbookData(): WorkbookData {
    return this.workbookData;
  }

  static createEmpty(): XlsxDocument {
    const zip = new JSZip();
    const sheet: SheetData = {
      id: "sheet1",
      name: "Sheet1",
      cells: {},
      rowCount: 50,
      colCount: 26,
      colWidths: {},
      rowHeights: {},
    };
    return new XlsxDocument(zip, {
      sheets: [sheet],
      activeSheetIndex: 0,
    });
  }

  async save(data: WorkbookData): Promise<ArrayBuffer> {
    this.workbookData = data;
    const zip = this.zip;

    const bundle = defaultXlsxService.buildXmlBundle(data);

    // 1. Worksheets
    for (const ws of bundle.worksheets) {
      zip.file(ws.path, ws.xml);
    }

    // 2. Shared Strings
    zip.file("xl/sharedStrings.xml", bundle.sharedStringsXml);

    // 3. Workbook
    zip.file("xl/workbook.xml", bundle.workbookXml);

    // 4. Relationships
    zip.file("xl/_rels/workbook.xml.rels", bundle.relationshipsXml);

    // 5. Base Content_Types and styles if missing
    if (!zip.file("[Content_Types].xml")) {
      zip.file(
        "[Content_Types].xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
      );
    }

    if (!zip.file("_rels/.rels")) {
      zip.file(
        "_rels/.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
      );
    }

    if (!zip.file("xl/styles.xml")) {
      zip.file(
        "xl/styles.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`
      );
    }

    return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  }
}
