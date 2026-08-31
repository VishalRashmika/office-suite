import { CellData } from "./cell";

export interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface SheetData {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key e.g. "0,0" (row,col)
  rowCount: number;
  colCount: number;
  colWidths: Record<number, number>; // colIndex -> px width
  rowHeights: Record<number, number>; // rowIndex -> px height
  merges?: MergeRange[];
}

export interface WorkbookData {
  sheets: SheetData[];
  activeSheetIndex: number;
}

export function createDefaultSheet(id = "sheet1", name = "Sheet1", rowCount = 40, colCount = 26): SheetData {
  return {
    id,
    name,
    cells: {},
    rowCount,
    colCount,
    colWidths: {},
    rowHeights: {},
    merges: [],
  };
}

export function createDefaultWorkbook(): WorkbookData {
  return {
    sheets: [createDefaultSheet()],
    activeSheetIndex: 0,
  };
}
