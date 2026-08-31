import { WorkbookData, SheetData } from "../models/workbook";
import { cellKey } from "../models/coordinates";

export class CsvService {
  parse(text: string, delimiter?: string): WorkbookData {
    // Detect delimiter if not provided
    if (!delimiter) {
      const firstLine = text.split(/\r?\n/)[0] || "";
      const commaCount = (firstLine.match(/,/g) || []).length;
      const tabCount = (firstLine.match(/\t/g) || []).length;
      const semiCount = (firstLine.match(/;/g) || []).length;
      if (tabCount > commaCount && tabCount > semiCount) delimiter = "\t";
      else if (semiCount > commaCount && semiCount > tabCount) delimiter = ";";
      else delimiter = ",";
    }

    const rows: string[][] = [];
    let currentRow: string[] = [];
    let currentCell = "";
    let inQuotes = false;
    let i = 0;

    while (i < text.length) {
      const char = text[i];
      const nextChar = text[i + 1];

      if (inQuotes) {
        if (char === '"') {
          if (nextChar === '"') {
            // Escaped quote
            currentCell += '"';
            i += 2;
            continue;
          } else {
            // End quote
            inQuotes = false;
            i++;
            continue;
          }
        } else {
          currentCell += char;
          i++;
          continue;
        }
      } else {
        if (char === '"') {
          inQuotes = true;
          i++;
          continue;
        } else if (char === delimiter) {
          currentRow.push(currentCell);
          currentCell = "";
          i++;
          continue;
        } else if (char === "\r" && nextChar === "\n") {
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = "";
          i += 2;
          continue;
        } else if (char === "\n" || char === "\r") {
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = "";
          i++;
          continue;
        } else {
          currentCell += char;
          i++;
          continue;
        }
      }
    }

    if (currentCell.length > 0 || currentRow.length > 0) {
      currentRow.push(currentCell);
      rows.push(currentRow);
    }

    // Remove empty trailing rows
    while (rows.length > 1 && rows[rows.length - 1].every((c) => !c.trim())) {
      rows.pop();
    }

    const maxCols = Math.max(10, ...rows.map((r) => r.length));
    const maxRows = Math.max(25, rows.length + 5);

    const sheetCells: SheetData["cells"] = {};
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const raw = rows[r][c];
        if (raw !== "") {
          const num = Number(raw);
          const isNum = !isNaN(num) && raw.trim() !== "";
          sheetCells[cellKey(r, c)] = {
            value: isNum ? num : raw,
            formatted: raw,
          };
        }
      }
    }

    const sheet: SheetData = {
      id: "sheet1",
      name: "Sheet1",
      cells: sheetCells,
      rowCount: maxRows,
      colCount: maxCols,
      colWidths: {},
      rowHeights: {},
    };

    return {
      sheets: [sheet],
      activeSheetIndex: 0,
    };
  }

  serialize(workbook: WorkbookData, delimiter = ","): string {
    const sheet = workbook.sheets[workbook.activeSheetIndex || 0] || workbook.sheets[0];
    if (!sheet) return "";

    // Find max row and col with data
    let maxR = 0;
    let maxC = 0;
    for (const k in sheet.cells) {
      const [r, c] = k.split(",").map((n) => parseInt(n, 10));
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }

    const lines: string[] = [];
    for (let r = 0; r <= maxR; r++) {
      const rowCells: string[] = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = sheet.cells[cellKey(r, c)];
        let val = "";
        if (cell) {
          if (cell.formula) {
            val = cell.formula;
          } else if (cell.value !== undefined && cell.value !== null) {
            val = String(cell.value);
          }
        }

        // Escape quotes and delimiters
        if (val.includes('"') || val.includes(delimiter) || val.includes("\n") || val.includes("\r")) {
          val = `"${val.replace(/"/g, '""')}"`;
        }
        rowCells.push(val);
      }
      lines.push(rowCells.join(delimiter));
    }

    return lines.join("\r\n");
  }
}

export const defaultCsvService = new CsvService();

export function parseCSV(text: string, delimiter?: string): WorkbookData {
  return defaultCsvService.parse(text, delimiter);
}

export function serializeCSV(workbook: WorkbookData, delimiter = ","): string {
  return defaultCsvService.serialize(workbook, delimiter);
}
