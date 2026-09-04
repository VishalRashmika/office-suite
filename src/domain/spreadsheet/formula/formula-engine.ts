import { SheetData } from "../models/workbook";
import { formatCellValue, CellValue } from "../models/cell";
import { evaluateExpression, EvaluationContext } from "./expression-evaluator";

export class FormulaEngine {
  evaluateSheet(sheet: SheetData): void {
    const visited = new Set<string>();
    const evaluating = new Set<string>();

    for (const key in sheet.cells) {
      const cell = sheet.cells[key];
      if (cell && cell.formula && cell.formula.startsWith("=")) {
        this.evalCell(key, sheet, visited, evaluating);
      }
    }
  }

  evalCell(key: string, sheet: SheetData, visited: Set<string>, evaluating: Set<string>): CellValue {
    if (evaluating.has(key)) {
      return "#CIRCULAR!";
    }
    const cell = sheet.cells[key];
    if (!cell) return 0;
    if (!cell.formula || !cell.formula.startsWith("=")) {
      return cell.value;
    }
    if (visited.has(key)) {
      return cell.value;
    }

    evaluating.add(key);
    try {
      const formulaStr = cell.formula.substring(1).trim();
      const ctx: EvaluationContext = {
        sheet,
        visited,
        evaluating,
        evalCell: (k, s, v, e) => this.evalCell(k, s, v, e),
      };
      const result = evaluateExpression(formulaStr, ctx) as CellValue;
      cell.value = result;
      cell.formatted = formatCellValue(result, cell.style?.numberFormat);
      visited.add(key);
      return result;
    } catch {
      cell.value = "#ERROR!";
      cell.formatted = "#ERROR!";
      return "#ERROR!";
    } finally {
      evaluating.delete(key);
    }
  }

  evaluateSingleExpression(
    expr: string,
    sheet: SheetData,
    visited = new Set<string>(),
    evaluating = new Set<string>()
  ): CellValue {
    const ctx: EvaluationContext = {
      sheet,
      visited,
      evaluating,
      evalCell: (k, s, v, e) => this.evalCell(k, s, v, e),
    };
    return evaluateExpression(expr, ctx);
  }
}

export const defaultFormulaEngine = new FormulaEngine();

export function evaluateSheetFormulas(sheet: SheetData): void {
  defaultFormulaEngine.evaluateSheet(sheet);
}

export function evaluateExpressionStandalone(
  expr: string,
  sheet: SheetData,
  visited = new Set<string>(),
  evaluating = new Set<string>()
): CellValue {
  return defaultFormulaEngine.evaluateSingleExpression(expr, sheet, visited, evaluating);
}
