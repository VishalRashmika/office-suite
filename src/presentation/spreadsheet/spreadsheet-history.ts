import { WorkbookData } from "../../domain/spreadsheet";

export interface HistoryEntry {
  sheetsJson: string;
  activeSheetIndex: number;
}

export class SpreadsheetHistory {
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private maxDepth = 50;

  push(workbook: WorkbookData): void {
    this.undoStack.push({
      sheetsJson: JSON.stringify(workbook.sheets),
      activeSheetIndex: workbook.activeSheetIndex,
    });
    if (this.undoStack.length > this.maxDepth) {
      this.undoStack.shift();
    }
    this.redoStack = [];
  }

  undo(workbook: WorkbookData): boolean {
    if (this.undoStack.length <= 1) return false;
    const current = this.undoStack.pop()!;
    this.redoStack.push(current);
    const prev = this.undoStack[this.undoStack.length - 1];
    workbook.sheets = JSON.parse(prev.sheetsJson);
    workbook.activeSheetIndex = prev.activeSheetIndex;
    return true;
  }

  redo(workbook: WorkbookData): boolean {
    if (!this.redoStack.length) return false;
    const next = this.redoStack.pop()!;
    this.undoStack.push(next);
    workbook.sheets = JSON.parse(next.sheetsJson);
    workbook.activeSheetIndex = next.activeSheetIndex;
    return true;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
