import { FileView, TFile, WorkspaceLeaf, Notice } from "obsidian";
import {
  WorkbookData,
  SheetData,
  CellCoord,
  CellRange,
  CellStyle,
  CellData,
  cellKey,
  evaluateSheetFormulas,
  parseCSV,
  serializeCSV,
} from "../../domain/spreadsheet";
import { XlsxDocument } from "../../infrastructure/spreadsheet/xlsx-document";
import { SpreadsheetHistory } from "./spreadsheet-history";
import { FormulaBar, FormulaBarDelegate } from "./formula-bar";
import { SpreadsheetTabs, SpreadsheetTabsDelegate } from "./spreadsheet-tabs";
import { SpreadsheetGrid, SpreadsheetGridDelegate } from "./spreadsheet-grid";
import { SpreadsheetToolbar, SpreadsheetToolbarDelegate } from "./spreadsheet-toolbar";

export const VIEW_TYPE_SHEET = "spreadsheet-editor-view";

export class SpreadsheetView
  extends FileView
  implements FormulaBarDelegate, SpreadsheetTabsDelegate, SpreadsheetGridDelegate, SpreadsheetToolbarDelegate
{
  allowNoFile = false;
  workbook: WorkbookData | null = null;
  xlsxDoc: XlsxDocument | null = null;
  isCsv = false;
  dirty = false;
  paperTheme: "light" | "dark" = "light";

  private history = new SpreadsheetHistory();
  private toolbar = new SpreadsheetToolbar();
  private formulaBar = new FormulaBar();
  private grid = new SpreadsheetGrid();
  private tabs = new SpreadsheetTabs();

  private toolbarEl: HTMLElement | null = null;
  private formulaBarEl: HTMLElement | null = null;
  private gridContainerEl: HTMLElement | null = null;
  private bottomTabBarEl: HTMLElement | null = null;
  private saveTimer: number | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_SHEET;
  }

  getIcon(): string {
    return "table";
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Spreadsheet";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "xlsx" || extension === "csv" || extension === "tsv";
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("docx-editor-root", "sheet-editor-root");

    const status = this.contentEl.createDiv({ cls: "docx-editor-status", text: "Loading spreadsheet..." });

    this.isCsv = file.extension === "csv" || file.extension === "tsv";

    try {
      if (this.isCsv) {
        const text = await this.app.vault.read(file);
        this.workbook = parseCSV(text, file.extension === "tsv" ? "\t" : undefined);
      } else {
        const bytes = await this.app.vault.readBinary(file);
        this.xlsxDoc = await XlsxDocument.load(bytes);
        this.workbook = this.xlsxDoc.getWorkbookData();
      }
    } catch (err) {
      status.setText(`Could not open this spreadsheet: ${(err as Error).message}`);
      console.error(err);
      return;
    }

    status.remove();
    this.buildUI();
    this.setTheme(this.paperTheme);
    window.addEventListener("blur", this.onWindowBlur);
    this.history.push(this.workbook);
  }

  private onWindowBlur = () => {
    if (this.dirty) this.saveNow();
  };

  async onUnloadFile(file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    this.history.clear();
    this.workbook = null;
    this.xlsxDoc = null;
  }

  get currentSheet(): SheetData {
    if (!this.workbook || !this.workbook.sheets.length) {
      return {
        id: "sheet1",
        name: "Sheet1",
        cells: {},
        rowCount: 40,
        colCount: 26,
        colWidths: {},
        rowHeights: {},
      };
    }
    return this.workbook.sheets[this.workbook.activeSheetIndex || 0] || this.workbook.sheets[0];
  }

  private buildUI(): void {
    this.contentEl.empty();

    // 1. Toolbar
    this.toolbarEl = this.contentEl.createDiv({ cls: "docx-editor-toolbar sheet-toolbar" });
    this.toolbar.build(this.toolbarEl, this);

    // 2. Formula Bar (fx)
    this.formulaBarEl = this.contentEl.createDiv({ cls: "sheet-formula-bar" });
    this.formulaBar.build(this.formulaBarEl, this);

    // 3. Grid View Area
    this.gridContainerEl = this.contentEl.createDiv({ cls: "sheet-grid-container" });
    this.renderGrid();

    // 4. Bottom Tab Bar
    this.bottomTabBarEl = this.contentEl.createDiv({ cls: "sheet-tab-bar" });
    this.renderTabs();
  }

  private renderGrid(): void {
    if (!this.gridContainerEl) return;
    this.grid.render(this.gridContainerEl, this.currentSheet, this);
    this.updateFormulaBar();
  }

  private renderTabs(): void {
    if (!this.bottomTabBarEl || !this.workbook) return;
    this.tabs.render(this.bottomTabBarEl, this.workbook, this);
  }

  private updateFormulaBar(): void {
    const cell = this.currentSheet.cells[cellKey(this.grid.activeCell.row, this.grid.activeCell.col)];
    this.formulaBar.update(this.grid.activeCell, cell);
  }

  // --- FORMULA BAR DELEGATE ---

  onAddressEnter(coord: CellCoord): void {
    this.grid.selectCell(coord.row, coord.col, this.currentSheet);
    this.updateFormulaBar();
  }

  onFormulaInput(val: string): void {
    const sheet = this.currentSheet;
    const key = cellKey(this.grid.activeCell.row, this.grid.activeCell.col);
    if (!sheet.cells[key]) sheet.cells[key] = {};

    if (val.startsWith("=")) {
      sheet.cells[key].formula = val;
      sheet.cells[key].value = undefined;
      sheet.cells[key].formatted = undefined;
    } else {
      sheet.cells[key].formula = undefined;
      const num = Number(val);
      const isNum = !isNaN(num) && val.trim() !== "";
      sheet.cells[key].value = isNum ? num : val;
      sheet.cells[key].formatted = val;
    }
    evaluateSheetFormulas(sheet);
    this.grid.updateCellDOM(this.grid.activeCell.row, this.grid.activeCell.col, sheet);
    this.scheduleSave();
  }

  onFormulaEnter(): void {
    this.grid.selectCell(this.grid.activeCell.row + 1, this.grid.activeCell.col, this.currentSheet);
    this.updateFormulaBar();
    this.gridContainerEl?.focus();
  }

  // --- SPREADSHEET TABS DELEGATE ---

  onSelectSheet(index: number): void {
    if (!this.workbook) return;
    this.workbook.activeSheetIndex = index;
    this.renderGrid();
    this.renderTabs();
  }

  onRenameSheet(index: number, newName: string): void {
    if (!this.workbook || !this.workbook.sheets[index]) return;
    this.workbook.sheets[index].name = newName;
    this.renderTabs();
    this.scheduleSave();
  }

  onAddSheet(): void {
    if (!this.workbook) return;
    const newIdx = this.workbook.sheets.length + 1;
    this.workbook.sheets.push({
      id: `sheet_${Date.now()}`,
      name: `Sheet${newIdx}`,
      cells: {},
      rowCount: 40,
      colCount: 26,
      colWidths: {},
      rowHeights: {},
    });
    this.workbook.activeSheetIndex = this.workbook.sheets.length - 1;
    this.renderGrid();
    this.renderTabs();
    this.history.push(this.workbook);
    this.scheduleSave();
  }

  // --- SPREADSHEET GRID DELEGATE ---

  onSelectionChange(activeCoord: CellCoord, range: CellRange): void {
    this.updateFormulaBar();
  }

  onCellEditFinished(row: number, col: number, rawValue: string): void {
    const sheet = this.currentSheet;
    const key = cellKey(row, col);
    if (!sheet.cells[key]) sheet.cells[key] = {};

    if (rawValue.startsWith("=")) {
      sheet.cells[key].formula = rawValue;
      sheet.cells[key].value = undefined;
      sheet.cells[key].formatted = undefined;
    } else {
      sheet.cells[key].formula = undefined;
      const num = Number(rawValue);
      const isNum = !isNaN(num) && rawValue.trim() !== "";
      sheet.cells[key].value = isNum ? num : rawValue;
      sheet.cells[key].formatted = rawValue;
    }

    if (this.workbook) this.history.push(this.workbook);
    evaluateSheetFormulas(sheet);
    this.renderGrid();
    this.scheduleSave();
  }

  onColumnResized(colIdx: number, newWidth: number): void {
    if (this.workbook) this.history.push(this.workbook);
    this.scheduleSave();
  }

  onRowResized(rowIdx: number, newHeight: number): void {
    if (this.workbook) this.history.push(this.workbook);
    this.scheduleSave();
  }

  // --- ACTIONS & TOOLBAR DELEGATE ---

  onUndo(): void {
    if (!this.workbook) return;
    if (this.history.undo(this.workbook)) {
      this.renderGrid();
      this.renderTabs();
      this.scheduleSave();
    }
  }

  onRedo(): void {
    if (!this.workbook) return;
    if (this.history.redo(this.workbook)) {
      this.renderGrid();
      this.renderTabs();
      this.scheduleSave();
    }
  }

  onApplyFormat(format: string): void {
    this.onApplyStyle({ numberFormat: format });
  }

  onAdjustDecimals(delta: number): void {
    this.renderGrid();
    this.scheduleSave();
  }

  onApplyStyle(style: Partial<CellStyle>): void {
    const sheet = this.currentSheet;
    const minR = Math.min(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const maxR = Math.max(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const minC = Math.min(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);
    const maxC = Math.max(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const key = cellKey(r, c);
        if (!sheet.cells[key]) sheet.cells[key] = {};
        sheet.cells[key].style = { ...sheet.cells[key].style, ...style };
      }
    }

    if (this.workbook) this.history.push(this.workbook);
    this.renderGrid();
    this.scheduleSave();
  }

  onToggleStyle(prop: keyof CellStyle): void {
    const sheet = this.currentSheet;
    const key = cellKey(this.grid.activeCell.row, this.grid.activeCell.col);
    const curVal = sheet.cells[key]?.style?.[prop];
    this.onApplyStyle({ [prop]: !curVal });
  }

  onToggleBorders(): void {
    const sheet = this.currentSheet;
    const key = cellKey(this.grid.activeCell.row, this.grid.activeCell.col);
    const hasBorder = sheet.cells[key]?.style?.borderBottom;
    const borderStyle = hasBorder ? undefined : "1px solid #1a1a1a";
    this.onApplyStyle({
      borderTop: borderStyle,
      borderBottom: borderStyle,
      borderLeft: borderStyle,
      borderRight: borderStyle,
    });
  }

  onClear(): void {
    const sheet = this.currentSheet;
    const minR = Math.min(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const maxR = Math.max(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const minC = Math.min(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);
    const maxC = Math.max(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        delete sheet.cells[cellKey(r, c)];
      }
    }

    if (this.workbook) this.history.push(this.workbook);
    evaluateSheetFormulas(sheet);
    this.renderGrid();
    this.scheduleSave();
  }

  onCopy(): void {
    const sheet = this.currentSheet;
    const minR = Math.min(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const maxR = Math.max(this.grid.selectionRange.startRow, this.grid.selectionRange.endRow);
    const minC = Math.min(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);
    const maxC = Math.max(this.grid.selectionRange.startCol, this.grid.selectionRange.endCol);

    const rows: string[] = [];
    for (let r = minR; r <= maxR; r++) {
      const rowVals: string[] = [];
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.cells[cellKey(r, c)];
        rowVals.push(
          cell?.formula || (cell?.value !== undefined && cell?.value !== null ? String(cell.value) : "")
        );
      }
      rows.push(rowVals.join("\t"));
    }

    navigator.clipboard.writeText(rows.join("\n"));
  }

  async onPaste(): Promise<void> {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) return;

      const lines = text.split(/\r?\n/).map((l) => l.split("\t"));
      const sheet = this.currentSheet;
      const startR = this.grid.activeCell.row;
      const startC = this.grid.activeCell.col;

      for (let r = 0; r < lines.length; r++) {
        for (let c = 0; c < lines[r].length; c++) {
          const val = lines[r][c];
          const targetR = startR + r;
          const targetC = startC + c;
          const key = cellKey(targetR, targetC);

          if (!sheet.cells[key]) sheet.cells[key] = {};
          if (val.startsWith("=")) {
            sheet.cells[key].formula = val;
          } else {
            sheet.cells[key].formula = undefined;
            const num = Number(val);
            sheet.cells[key].value = !isNaN(num) && val.trim() !== "" ? num : val;
          }
        }
      }

      if (this.workbook) this.history.push(this.workbook);
      evaluateSheetFormulas(sheet);
      this.renderGrid();
      this.scheduleSave();
    } catch (err) {
      console.warn("Clipboard paste error:", err);
    }
  }

  onInsertRow(): void {
    const sheet = this.currentSheet;
    const targetR = this.grid.activeCell.row + 1;
    const newCells: Record<string, CellData> = {};

    for (const key in sheet.cells) {
      const [r, c] = key.split(",").map((n) => parseInt(n, 10));
      if (r >= targetR) {
        newCells[cellKey(r + 1, c)] = sheet.cells[key];
      } else {
        newCells[key] = sheet.cells[key];
      }
    }

    sheet.cells = newCells;
    sheet.rowCount++;
    if (this.workbook) this.history.push(this.workbook);
    this.renderGrid();
    this.scheduleSave();
  }

  onInsertColumn(): void {
    const sheet = this.currentSheet;
    const targetC = this.grid.activeCell.col + 1;
    const newCells: Record<string, CellData> = {};

    for (const key in sheet.cells) {
      const [r, c] = key.split(",").map((n) => parseInt(n, 10));
      if (c >= targetC) {
        newCells[cellKey(r, c + 1)] = sheet.cells[key];
      } else {
        newCells[key] = sheet.cells[key];
      }
    }

    sheet.cells = newCells;
    sheet.colCount++;
    if (this.workbook) this.history.push(this.workbook);
    this.renderGrid();
    this.scheduleSave();
  }

  onDeleteRow(): void {
    const sheet = this.currentSheet;
    const targetR = this.grid.activeCell.row;
    const newCells: Record<string, CellData> = {};

    for (const key in sheet.cells) {
      const [r, c] = key.split(",").map((n) => parseInt(n, 10));
      if (r === targetR) continue;
      if (r > targetR) {
        newCells[cellKey(r - 1, c)] = sheet.cells[key];
      } else {
        newCells[key] = sheet.cells[key];
      }
    }

    sheet.cells = newCells;
    sheet.rowCount = Math.max(10, sheet.rowCount - 1);
    if (this.workbook) this.history.push(this.workbook);
    this.renderGrid();
    this.scheduleSave();
  }

  onDeleteColumn(): void {
    const sheet = this.currentSheet;
    const targetC = this.grid.activeCell.col;
    const newCells: Record<string, CellData> = {};

    for (const key in sheet.cells) {
      const [r, c] = key.split(",").map((n) => parseInt(n, 10));
      if (c === targetC) continue;
      if (c > targetC) {
        newCells[cellKey(r, c - 1)] = sheet.cells[key];
      } else {
        newCells[key] = sheet.cells[key];
      }
    }

    sheet.cells = newCells;
    sheet.colCount = Math.max(5, sheet.colCount - 1);
    if (this.workbook) this.history.push(this.workbook);
    this.renderGrid();
    this.scheduleSave();
  }

  onInsertFunction(fn: string): void {
    const sheet = this.currentSheet;
    const key = cellKey(this.grid.activeCell.row, this.grid.activeCell.col);
    if (!sheet.cells[key]) sheet.cells[key] = {};
    sheet.cells[key].formula = `=${fn}()`;
    this.updateFormulaBar();
    this.formulaBar.focusFormulaInput();
  }

  setTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    if (theme === "light") {
      this.contentEl.removeClass("sheet-theme-dark");
      this.contentEl.addClass("sheet-theme-light");
      this.toolbar.lightThemeBtn?.addClass("is-active");
      this.toolbar.darkThemeBtn?.removeClass("is-active");
    } else {
      this.contentEl.removeClass("sheet-theme-light");
      this.contentEl.addClass("sheet-theme-dark");
      this.toolbar.darkThemeBtn?.addClass("is-active");
      this.toolbar.lightThemeBtn?.removeClass("is-active");
    }
  }

  // --- PERSISTENCE ---

  private scheduleSave(): void {
    this.dirty = true;
    this.toolbar.updateSaveStatus(this.dirty);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => this.saveNow(), 600);
  }

  private async saveNow(): Promise<void> {
    if (!this.workbook || !this.file) return;
    try {
      if (this.isCsv) {
        const csvText = serializeCSV(this.workbook, this.file.extension === "tsv" ? "\t" : ",");
        await this.app.vault.modify(this.file, csvText);
      } else {
        if (!this.xlsxDoc) this.xlsxDoc = XlsxDocument.createEmpty();
        const bytes = await this.xlsxDoc.save(this.workbook);
        await this.app.vault.modifyBinary(this.file, bytes);
      }
      this.dirty = false;
    } catch (err) {
      console.error("Failed to save spreadsheet:", err);
      new Notice(`Failed to save ${this.file.name}: ${(err as Error).message}`);
    } finally {
      this.toolbar.updateSaveStatus(this.dirty);
    }
  }

  async requestSave(): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
  }
}
