import {
  SheetData,
  CellCoord,
  CellRange,
  CellData,
  cellKey,
  coordToA1,
  evaluateSheetFormulas,
} from "../../domain/spreadsheet";

export interface SpreadsheetGridDelegate {
  onSelectionChange(activeCoord: CellCoord, range: CellRange): void;
  onCellEditFinished(row: number, col: number, rawValue: string): void;
  onColumnResized(colIdx: number, newWidth: number): void;
  onRowResized(rowIdx: number, newHeight: number): void;
  onUndo(): void;
  onRedo(): void;
  onCopy(): void;
  onPaste(): void | Promise<void>;
  onClear(): void;
  onToggleStyle(prop: "bold" | "italic" | "underline"): void;
}

export class SpreadsheetGrid {
  private container: HTMLElement | null = null;
  public tableEl: HTMLTableElement | null = null;

  public activeCell: CellCoord = { row: 0, col: 0 };
  public selectionRange: CellRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
  public isSelecting = false;
  public isEditing = false;

  render(container: HTMLElement, sheet: SheetData, delegate: SpreadsheetGridDelegate): void {
    this.container = container;
    container.empty();

    evaluateSheetFormulas(sheet);

    const table = container.createEl("table", { cls: "sheet-table" });
    this.tableEl = table;

    // Header Row (A, B, C...)
    const thead = table.createEl("thead");
    const headerRow = thead.createEl("tr");

    // Top-left corner cell
    const cornerTh = headerRow.createEl("th", { cls: "sheet-corner-header" });
    cornerTh.onclick = () => {
      this.selectAll(sheet);
      delegate.onSelectionChange(this.activeCell, this.selectionRange);
    };

    for (let c = 0; c < sheet.colCount; c++) {
      const th = headerRow.createEl("th", { cls: "sheet-col-header" });
      th.setText(coordToA1(0, c).replace(/\d+/, ""));
      const width = sheet.colWidths[c] || 85;
      th.style.width = `${width}px`;

      // Col resize handle
      const resizer = th.createDiv({ cls: "sheet-col-resizer" });
      this.bindColResizer(resizer, c, sheet, delegate);
    }

    // Body Rows
    const tbody = table.createEl("tbody");
    for (let r = 0; r < sheet.rowCount; r++) {
      const tr = tbody.createEl("tr");
      const rowHeight = sheet.rowHeights[r] || 24;
      tr.style.height = `${rowHeight}px`;

      // Row Number Header (1, 2, 3...)
      const rowTh = tr.createEl("th", { cls: "sheet-row-header" });
      rowTh.setText(String(r + 1));
      const rowResizer = rowTh.createDiv({ cls: "sheet-row-resizer" });
      this.bindRowResizer(rowResizer, r, sheet, delegate);

      for (let c = 0; c < sheet.colCount; c++) {
        const td = tr.createEl("td", { cls: "sheet-cell" });
        td.dataset.row = String(r);
        td.dataset.col = String(c);

        const width = sheet.colWidths[c] || 85;
        td.style.width = `${width}px`;

        this.applyCellStylesToDOM(td, sheet.cells[cellKey(r, c)]);
        this.renderCellContent(td, sheet.cells[cellKey(r, c)]);

        // Events
        td.onmousedown = (e) => {
          if (e.button !== 0) return;
          this.isSelecting = true;
          this.selectCell(r, c, sheet, e.shiftKey);
          delegate.onSelectionChange(this.activeCell, this.selectionRange);
        };

        td.onmouseenter = () => {
          if (this.isSelecting) {
            this.expandSelectionTo(r, c);
            delegate.onSelectionChange(this.activeCell, this.selectionRange);
          }
        };

        td.ondblclick = () => {
          this.startInCellEdit(r, c, td, sheet, delegate);
        };
      }
    }

    window.onmouseup = () => {
      this.isSelecting = false;
    };

    this.updateSelectionHighlights();
    this.bindKeyboardEvents(container, sheet, delegate);
  }

  renderCellContent(td: HTMLTableCellElement, cell?: CellData): void {
    if (!cell) {
      td.setText("");
      return;
    }
    if (cell.formatted !== undefined && cell.formatted !== null && cell.formatted !== "") {
      td.setText(String(cell.formatted));
    } else if (cell.value !== undefined && cell.value !== null && cell.value !== "") {
      td.setText(String(cell.value));
    } else if (cell.formula) {
      td.setText(String(cell.formula));
    } else {
      td.setText("");
    }
  }

  applyCellStylesToDOM(td: HTMLTableCellElement, cell?: CellData): void {
    if (!cell || !cell.style) return;
    const s = cell.style;
    const styles: Partial<CSSStyleDeclaration> = {};
    if (s.bold) styles.fontWeight = "bold";
    if (s.italic) styles.fontStyle = "italic";
    if (s.underline && s.strike) styles.textDecoration = "underline line-through";
    else if (s.underline) styles.textDecoration = "underline";
    else if (s.strike) styles.textDecoration = "line-through";
    if (s.color) styles.color = s.color;
    if (s.background) styles.backgroundColor = s.background;
    if (s.align) styles.textAlign = s.align;
    if (s.fontSize) styles.fontSize = `${s.fontSize}pt`;
    if (s.fontFamily) styles.fontFamily = s.fontFamily;
    if (s.borderTop) styles.borderTop = s.borderTop;
    if (s.borderBottom) styles.borderBottom = s.borderBottom;
    if (s.borderLeft) styles.borderLeft = s.borderLeft;
    if (s.borderRight) styles.borderRight = s.borderRight;
    td.setCssStyles(styles);
  }

  updateCellDOM(r: number, c: number, sheet: SheetData): void {
    if (!this.tableEl) return;
    const td = this.tableEl.querySelector(`td[data-row="${r}"][data-col="${c}"]`) as HTMLTableCellElement;
    if (td) {
      const cell = sheet.cells[cellKey(r, c)];
      this.renderCellContent(td, cell);
      this.applyCellStylesToDOM(td, cell);
    }
  }

  selectCell(row: number, col: number, sheet: SheetData, extend = false): void {
    row = Math.max(0, Math.min(sheet.rowCount - 1, row));
    col = Math.max(0, Math.min(sheet.colCount - 1, col));

    this.activeCell = { row, col };
    if (!extend) {
      this.selectionRange = { startRow: row, startCol: col, endRow: row, endCol: col };
    } else {
      this.selectionRange.endRow = row;
      this.selectionRange.endCol = col;
    }

    this.updateSelectionHighlights();
  }

  expandSelectionTo(row: number, col: number): void {
    this.selectionRange.endRow = row;
    this.selectionRange.endCol = col;
    this.updateSelectionHighlights();
  }

  selectAll(sheet: SheetData): void {
    this.selectionRange = {
      startRow: 0,
      startCol: 0,
      endRow: sheet.rowCount - 1,
      endCol: sheet.colCount - 1,
    };
    this.updateSelectionHighlights();
  }

  updateSelectionHighlights(): void {
    if (!this.tableEl) return;
    const minR = Math.min(this.selectionRange.startRow, this.selectionRange.endRow);
    const maxR = Math.max(this.selectionRange.startRow, this.selectionRange.endRow);
    const minC = Math.min(this.selectionRange.startCol, this.selectionRange.endCol);
    const maxC = Math.max(this.selectionRange.startCol, this.selectionRange.endCol);

    const cells = this.tableEl.querySelectorAll(".sheet-cell");
    cells.forEach((cell) => {
      const r = parseInt((cell as HTMLElement).dataset.row || "0", 10);
      const c = parseInt((cell as HTMLElement).dataset.col || "0", 10);

      const inRange = r >= minR && r <= maxR && c >= minC && c <= maxC;
      const isActive = r === this.activeCell.row && c === this.activeCell.col;

      cell.classList.toggle("sheet-cell-selected", inRange);
      cell.classList.toggle("sheet-cell-active", isActive);
    });
  }

  startInCellEdit(
    row: number,
    col: number,
    td: HTMLTableCellElement,
    sheet: SheetData,
    delegate: SpreadsheetGridDelegate,
    initialChar?: string
  ): void {
    this.isEditing = true;
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    const initialText =
      initialChar !== undefined
        ? initialChar
        : cell?.formula || (cell?.value !== undefined && cell?.value !== null ? String(cell.value) : "");

    td.empty();
    const input = td.createEl("input", { type: "text", cls: "sheet-cell-editor", value: initialText });
    input.focus();
    if (initialChar === undefined) {
      input.select();
    } else {
      input.setSelectionRange(initialText.length, initialText.length);
    }

    const finishEdit = () => {
      if (!this.isEditing) return;
      this.isEditing = false;
      delegate.onCellEditFinished(row, col, input.value);
    };

    input.onblur = finishEdit;
    input.onkeydown = (e) => {
      e.stopPropagation();
      if (e.key === "Enter") {
        finishEdit();
        this.selectCell(row + 1, col, sheet);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "Escape") {
        this.isEditing = false;
        this.render(this.container!, sheet, delegate);
      } else if (e.key === "Tab") {
        e.preventDefault();
        finishEdit();
        this.selectCell(row, col + 1, sheet);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      }
    };
  }

  private bindKeyboardEvents(
    container: HTMLElement,
    sheet: SheetData,
    delegate: SpreadsheetGridDelegate
  ): void {
    container.tabIndex = 0;
    container.onkeydown = (e) => {
      if (this.isEditing) return;

      if (e.key === "ArrowUp") {
        e.preventDefault();
        this.selectCell(this.activeCell.row - 1, this.activeCell.col, sheet, e.shiftKey);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        this.selectCell(this.activeCell.row + 1, this.activeCell.col, sheet, e.shiftKey);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "ArrowLeft") {
        e.preventDefault();
        this.selectCell(this.activeCell.row, this.activeCell.col - 1, sheet, e.shiftKey);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        this.selectCell(this.activeCell.row, this.activeCell.col + 1, sheet, e.shiftKey);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "Tab") {
        e.preventDefault();
        this.selectCell(this.activeCell.row, e.shiftKey ? this.activeCell.col - 1 : this.activeCell.col + 1, sheet);
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "Enter") {
        e.preventDefault();
        this.selectCell(
          e.shiftKey ? this.activeCell.row - 1 : this.activeCell.row + 1,
          this.activeCell.col,
          sheet
        );
        delegate.onSelectionChange(this.activeCell, this.selectionRange);
      } else if (e.key === "Delete" || e.key === "Backspace") {
        delegate.onClear();
      } else if (e.ctrlKey || e.metaKey) {
        if (e.key === "z" || e.key === "Z") {
          e.preventDefault();
          if (e.shiftKey) delegate.onRedo();
          else delegate.onUndo();
        } else if (e.key === "y" || e.key === "Y") {
          e.preventDefault();
          delegate.onRedo();
        } else if (e.key === "c" || e.key === "C") {
          delegate.onCopy();
        } else if (e.key === "v" || e.key === "V") {
          void delegate.onPaste();
        } else if (e.key === "b" || e.key === "B") {
          e.preventDefault();
          delegate.onToggleStyle("bold");
        } else if (e.key === "i" || e.key === "I") {
          e.preventDefault();
          delegate.onToggleStyle("italic");
        } else if (e.key === "u" || e.key === "U") {
          e.preventDefault();
          delegate.onToggleStyle("underline");
        }
      } else if (e.key === "F2") {
        e.preventDefault();
        const td = this.tableEl?.querySelector(
          `td[data-row="${this.activeCell.row}"][data-col="${this.activeCell.col}"]`
        ) as HTMLTableCellElement;
        if (td) {
          this.startInCellEdit(this.activeCell.row, this.activeCell.col, td, sheet, delegate);
        }
      } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        // Alphanumeric typing -> start edit with initial char
        e.preventDefault();
        const td = this.tableEl?.querySelector(
          `td[data-row="${this.activeCell.row}"][data-col="${this.activeCell.col}"]`
        ) as HTMLTableCellElement;
        if (td) {
          this.startInCellEdit(this.activeCell.row, this.activeCell.col, td, sheet, delegate, e.key);
        }
      }
    };
  }

  private bindColResizer(
    resizer: HTMLElement,
    colIdx: number,
    sheet: SheetData,
    delegate: SpreadsheetGridDelegate
  ): void {
    resizer.onmousedown = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startWidth = sheet.colWidths[colIdx] || 85;

      const onMouseMove = (moveEv: MouseEvent) => {
        const newWidth = Math.max(30, startWidth + (moveEv.clientX - startX));
        sheet.colWidths[colIdx] = newWidth;
        this.render(this.container!, sheet, delegate);
      };

      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
        delegate.onColumnResized(colIdx, sheet.colWidths[colIdx]);
      };

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    };
  }

  private bindRowResizer(
    resizer: HTMLElement,
    rowIdx: number,
    sheet: SheetData,
    delegate: SpreadsheetGridDelegate
  ): void {
    resizer.onmousedown = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const startY = e.clientY;
      const startHeight = sheet.rowHeights[rowIdx] || 24;

      const onMouseMove = (moveEv: MouseEvent) => {
        const newHeight = Math.max(18, startHeight + (moveEv.clientY - startY));
        sheet.rowHeights[rowIdx] = newHeight;
        this.render(this.container!, sheet, delegate);
      };

      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
        delegate.onRowResized(rowIdx, sheet.rowHeights[rowIdx]);
      };

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    };
  }
}
