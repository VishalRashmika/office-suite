import { WorkbookData } from "../models/workbook";
import { coordToA1 } from "../models/coordinates";
import { evaluateSheetFormulas } from "../formula/formula-engine";
import { escapeXml } from "../../../core/utils/xml";
import { CellValue } from "../models/cell";
import { XlsxStyleService } from "./xlsx-style-service";

export interface SerializedXlsxXmls {
  worksheets: { path: string; xml: string }[];
  sharedStringsXml: string;
  workbookXml: string;
  relationshipsXml: string;
  stylesXml: string;
}

export class XlsxService {
  buildXmlBundle(
    data: WorkbookData,
    styleService: XlsxStyleService = new XlsxStyleService(),
    originalDocs?: Map<string, Document>
  ): SerializedXlsxXmls {
    // 1. Collect all shared strings
    const stringMap = new Map<string, number>();
    const sharedStrings: string[] = [];

    const getStringIndex = (text: string): number => {
      if (stringMap.has(text)) return stringMap.get(text)!;
      const idx = sharedStrings.length;
      stringMap.set(text, idx);
      sharedStrings.push(text);
      return idx;
    };

    // 2. Build Worksheets XML
    const worksheets: { path: string; xml: string }[] = [];

    const parser =
      typeof DOMParser !== "undefined"
        ? new DOMParser()
        : new (globalThis as any).DOMParser();
    const serializer =
      typeof XMLSerializer !== "undefined"
        ? new XMLSerializer()
        : new (globalThis as any).XMLSerializer();

    for (let sIdx = 0; sIdx < data.sheets.length; sIdx++) {
      const sheet = data.sheets[sIdx];
      evaluateSheetFormulas(sheet);

      // Group cells by row
      const rowMap = new Map<
        number,
        { col: number; val: CellValue; formula?: string; styleIdx?: number }[]
      >();
      for (const key in sheet.cells) {
        const [r, c] = key.split(",").map((n) => parseInt(n, 10));
        const cell = sheet.cells[key];
        if (
          cell &&
          (cell.value !== undefined ||
            cell.formula ||
            (cell.style && Object.keys(cell.style).length > 0) ||
            cell.rawStyleIndex !== undefined)
        ) {
          const styleIdx = styleService.getOrCreateStyleIndex(cell.style, cell.rawStyleIndex);
          if (!rowMap.has(r)) rowMap.set(r, []);
          rowMap.get(r)!.push({
            col: c,
            val: cell.value,
            formula: cell.formula,
            styleIdx,
          });
        }
      }

      const sortedRows = Array.from(rowMap.keys()).sort((a, b) => a - b);
      let sheetDataXml = `<sheetData>`;

      for (const r of sortedRows) {
        const cols = rowMap.get(r)!.sort((a, b) => a.col - b.col);
        const rowHt = sheet.rowHeights[r];
        const htAttr = rowHt ? ` ht="${(rowHt / 1.33).toFixed(1)}" customHeight="1"` : "";
        sheetDataXml += `<row r="${r + 1}"${htAttr}>`;

        for (const cellItem of cols) {
          const a1 = coordToA1(r, cellItem.col);
          const fXml = cellItem.formula
            ? `<f>${escapeXml(cellItem.formula.replace(/^=/, ""))}</f>`
            : "";
          const sAttr =
            cellItem.styleIdx !== undefined && cellItem.styleIdx !== 0
              ? ` s="${cellItem.styleIdx}"`
              : "";

          if (typeof cellItem.val === "string" && !cellItem.formula) {
            const strIdx = getStringIndex(cellItem.val);
            sheetDataXml += `<c r="${a1}"${sAttr} t="s">${fXml}<v>${strIdx}</v></c>`;
          } else if (typeof cellItem.val === "boolean") {
            sheetDataXml += `<c r="${a1}"${sAttr} t="b">${fXml}<v>${cellItem.val ? 1 : 0}</v></c>`;
          } else if (typeof cellItem.val === "number") {
            sheetDataXml += `<c r="${a1}"${sAttr}>${fXml}<v>${cellItem.val}</v></c>`;
          } else if (cellItem.val !== null && cellItem.val !== undefined) {
            sheetDataXml += `<c r="${a1}"${sAttr}>${fXml}<v>${escapeXml(String(cellItem.val))}</v></c>`;
          } else if (cellItem.formula) {
            sheetDataXml += `<c r="${a1}"${sAttr}>${fXml}</c>`;
          } else if (cellItem.styleIdx !== undefined && cellItem.styleIdx !== 0) {
            sheetDataXml += `<c r="${a1}"${sAttr}/>`;
          }
        }
        sheetDataXml += `</row>`;
      }
      sheetDataXml += `</sheetData>`;

      // Column widths XML
      let colsXml = "";
      const colIndices = Object.keys(sheet.colWidths).map((k) => parseInt(k, 10));
      if (colIndices.length > 0) {
        colsXml = `<cols>`;
        for (const c of colIndices) {
          const w = (sheet.colWidths[c] / 8.5).toFixed(2);
          colsXml += `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`;
        }
        colsXml += `</cols>`;
      }

      // Merges XML
      let mergesXml = "";
      if (sheet.merges && sheet.merges.length > 0) {
        mergesXml = `<mergeCells count="${sheet.merges.length}">`;
        for (const m of sheet.merges) {
          const ref = `${coordToA1(m.startRow, m.startCol)}:${coordToA1(m.endRow, m.endCol)}`;
          mergesXml += `<mergeCell ref="${ref}"/>`;
        }
        mergesXml += `</mergeCells>`;
      }

      const targetPath = `xl/worksheets/sheet${sIdx + 1}.xml`;
      const origDoc = originalDocs?.get(targetPath);

      let finalWsXml: string;
      if (origDoc) {
        // Update sheetData in existing DOM
        const parsedDataDoc = parser.parseFromString(sheetDataXml, "application/xml");
        const newSheetData = parsedDataDoc.documentElement;
        const oldSheetData = origDoc.getElementsByTagName("sheetData")[0];
        if (oldSheetData && oldSheetData.parentNode) {
          const imported = origDoc.importNode(newSheetData, true);
          oldSheetData.parentNode.replaceChild(imported, oldSheetData);
        }

        // Update cols
        if (colsXml) {
          const parsedCols = parser.parseFromString(colsXml, "application/xml").documentElement;
          const oldCols = origDoc.getElementsByTagName("cols")[0];
          if (oldCols && oldCols.parentNode) {
            oldCols.parentNode.replaceChild(origDoc.importNode(parsedCols, true), oldCols);
          } else {
            const curSheetData = origDoc.getElementsByTagName("sheetData")[0];
            if (curSheetData && curSheetData.parentNode) {
              curSheetData.parentNode.insertBefore(origDoc.importNode(parsedCols, true), curSheetData);
            }
          }
        }

        // Update mergeCells
        if (mergesXml) {
          const parsedMerges = parser.parseFromString(mergesXml, "application/xml").documentElement;
          const oldMerges = origDoc.getElementsByTagName("mergeCells")[0];
          if (oldMerges && oldMerges.parentNode) {
            oldMerges.parentNode.replaceChild(origDoc.importNode(parsedMerges, true), oldMerges);
          } else {
            const curSheetData = origDoc.getElementsByTagName("sheetData")[0];
            if (curSheetData && curSheetData.parentNode) {
              if (curSheetData.nextSibling) {
                curSheetData.parentNode.insertBefore(origDoc.importNode(parsedMerges, true), curSheetData.nextSibling);
              } else {
                curSheetData.parentNode.appendChild(origDoc.importNode(parsedMerges, true));
              }
            }
          }
        }

        finalWsXml = serializer.serializeToString(origDoc);
        if (!finalWsXml.startsWith("<?xml")) {
          finalWsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + finalWsXml;
        }
      } else {
        finalWsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${colsXml}
  ${sheetDataXml}
  ${mergesXml}
</worksheet>`;
      }

      worksheets.push({
        path: targetPath,
        xml: finalWsXml,
      });
    }

    // 3. Write Shared Strings XML
    let sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">`;
    for (const str of sharedStrings) {
      sstXml += `<si><t>${escapeXml(str)}</t></si>`;
    }
    sstXml += `</sst>`;

    // 4. Write Workbook XML
    let sheetsXml = `<sheets>`;
    for (let i = 0; i < data.sheets.length; i++) {
      sheetsXml += `<sheet name="${escapeXml(data.sheets[i].name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
    }
    sheetsXml += `</sheets>`;

    const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${sheetsXml}
</workbook>`;

    // 5. Write Relationships
    let relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (let i = 0; i < data.sheets.length; i++) {
      relsXml += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
    }
    relsXml += `<Relationship Id="rId${data.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
    relsXml += `<Relationship Id="rId${data.sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    relsXml += `</Relationships>`;

    // 6. Styles XML
    const stylesXml = styleService.toXml();

    return {
      worksheets,
      sharedStringsXml: sstXml,
      workbookXml: wbXml,
      relationshipsXml: relsXml,
      stylesXml,
    };
  }
}

export const defaultXlsxService = new XlsxService();
