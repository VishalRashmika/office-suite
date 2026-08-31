import { CellCoord, coordToA1, a1ToCoord, CellData } from "../../domain/spreadsheet";

export interface FormulaBarDelegate {
  onAddressEnter(coord: CellCoord): void;
  onFormulaInput(value: string): void;
  onFormulaEnter(): void;
}

export class FormulaBar {
  private container: HTMLElement | null = null;
  private cellAddressInput: HTMLInputElement | null = null;
  private formulaInput: HTMLInputElement | null = null;

  build(container: HTMLElement, delegate: FormulaBarDelegate): void {
    this.container = container;
    container.empty();

    // 1. Cell address box (e.g. A1)
    const addrWrap = container.createDiv({ cls: "sheet-address-box" });
    this.cellAddressInput = addrWrap.createEl("input", {
      type: "text",
      value: "A1",
      cls: "sheet-address-input",
    });
    this.cellAddressInput.onkeydown = (e) => {
      if (e.key === "Enter") {
        const coord = a1ToCoord(this.cellAddressInput!.value);
        if (coord) {
          delegate.onAddressEnter(coord);
        }
      }
    };

    // 2. fx symbol
    container.createSpan({ cls: "sheet-fx-symbol", text: "fx" });

    // 3. Formula input
    const formulaInputWrap = container.createDiv({ cls: "sheet-formula-input-wrap" });
    this.formulaInput = formulaInputWrap.createEl("input", {
      type: "text",
      cls: "sheet-formula-input",
      placeholder: "Enter value or formula (=SUM(A1:B5))",
    });

    this.formulaInput.oninput = () => {
      delegate.onFormulaInput(this.formulaInput!.value);
    };

    this.formulaInput.onkeydown = (e) => {
      if (e.key === "Enter") {
        delegate.onFormulaEnter();
      }
    };
  }

  update(activeCoord: CellCoord, cell?: CellData): void {
    if (this.cellAddressInput) {
      this.cellAddressInput.value = coordToA1(activeCoord.row, activeCoord.col);
    }
    if (this.formulaInput) {
      if (cell) {
        this.formulaInput.value =
          cell.formula ||
          (cell.value !== undefined && cell.value !== null ? String(cell.value) : "");
      } else {
        this.formulaInput.value = "";
      }
    }
  }

  focusFormulaInput(): void {
    this.formulaInput?.focus();
  }

  setFormulaValue(val: string): void {
    if (this.formulaInput) {
      this.formulaInput.value = val;
    }
  }
}
