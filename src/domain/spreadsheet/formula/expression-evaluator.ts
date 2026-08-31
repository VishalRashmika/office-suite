import { SheetData } from "../models/workbook";
import { cellKey, a1ToCoord } from "../models/coordinates";
import { defaultFunctionRegistry, FunctionRegistry } from "./functions/registry";
import { FormulaFunctionContext } from "./functions/math";

export interface EvaluationContext {
  sheet: SheetData;
  visited: Set<string>;
  evaluating: Set<string>;
  evalCell: (key: string, sheet: SheetData, visited: Set<string>, evaluating: Set<string>) => any;
  registry?: FunctionRegistry;
}

export function evaluateExpression(expr: string, ctx: EvaluationContext): any {
  expr = expr.trim();
  if (!expr) return 0;

  // Check if string literal
  if (expr.startsWith('"') && expr.endsWith('"')) {
    return expr.slice(1, -1);
  }

  // Check if number
  if (/^-?\d+(\.\d+)?$/.test(expr)) {
    return parseFloat(expr);
  }

  // Function Call: e.g. SUM(A1:A5) or IF(A1>5, "Yes", "No")
  const funcMatch = expr.match(/^([A-Z_]+)\s*\((.*)\)$/is);
  if (funcMatch) {
    const funcName = funcMatch[1].toUpperCase();
    const argsStr = funcMatch[2];
    const args = splitArguments(argsStr);
    const registry = ctx.registry || defaultFunctionRegistry;
    const fn = registry.get(funcName);
    if (fn) {
      const fnContext: FormulaFunctionContext = {
        getRangeValues: (arg: string) => getRangeValues(arg, ctx),
        evaluateExpr: (innerExpr: string) => evaluateExpression(innerExpr, ctx),
      };
      return fn(args, fnContext);
    }
    return "#NAME?";
  }

  // Binary operators with precedence:
  // 1. Comparison: =, <>, <=, >=, <, >
  const compOp = findTopLevelOperator(expr, ["<=", ">=", "<>", "=", "<", ">"]);
  if (compOp) {
    const left = evaluateExpression(expr.slice(0, compOp.index), ctx);
    const right = evaluateExpression(expr.slice(compOp.index + compOp.op.length), ctx);
    switch (compOp.op) {
      case "=":
        return left == right;
      case "<>":
        return left != right;
      case "<":
        return Number(left) < Number(right);
      case ">":
        return Number(left) > Number(right);
      case "<=":
        return Number(left) <= Number(right);
      case ">=":
        return Number(left) >= Number(right);
    }
  }

  // 2. Concatenation: &
  const concatOp = findTopLevelOperator(expr, ["&"]);
  if (concatOp) {
    const left = evaluateExpression(expr.slice(0, concatOp.index), ctx);
    const right = evaluateExpression(expr.slice(concatOp.index + 1), ctx);
    return `${left ?? ""}${right ?? ""}`;
  }

  // 3. Addition / Subtraction: +, -
  const addSubOp = findTopLevelOperator(expr, ["+", "-"]);
  if (addSubOp && addSubOp.index > 0) {
    const left = evaluateExpression(expr.slice(0, addSubOp.index), ctx);
    const right = evaluateExpression(expr.slice(addSubOp.index + 1), ctx);
    return addSubOp.op === "+" ? Number(left) + Number(right) : Number(left) - Number(right);
  }

  // 4. Multiplication / Division: *, /, %
  const mulDivOp = findTopLevelOperator(expr, ["*", "/", "%"]);
  if (mulDivOp) {
    const left = evaluateExpression(expr.slice(0, mulDivOp.index), ctx);
    const right = evaluateExpression(expr.slice(mulDivOp.index + 1), ctx);
    if (mulDivOp.op === "*") return Number(left) * Number(right);
    if (mulDivOp.op === "/") return Number(right) === 0 ? "#DIV/0!" : Number(left) / Number(right);
    if (mulDivOp.op === "%") return Number(left) % Number(right);
  }

  // 5. Exponentiation: ^
  const expOp = findTopLevelOperator(expr, ["^"]);
  if (expOp) {
    const left = evaluateExpression(expr.slice(0, expOp.index), ctx);
    const right = evaluateExpression(expr.slice(expOp.index + 1), ctx);
    return Math.pow(Number(left), Number(right));
  }

  // Parentheses: (expr)
  if (expr.startsWith("(") && expr.endsWith(")")) {
    return evaluateExpression(expr.slice(1, -1), ctx);
  }

  // Single Cell Reference: A1, B4
  const coord = a1ToCoord(expr);
  if (coord) {
    return ctx.evalCell(cellKey(coord.row, coord.col), ctx.sheet, ctx.visited, ctx.evaluating);
  }

  return expr;
}

export function findTopLevelOperator(expr: string, ops: string[]): { op: string; index: number } | null {
  let depth = 0;
  let inQuotes = false;

  for (let i = expr.length - 1; i >= 0; i--) {
    const char = expr[i];
    if (char === '"') inQuotes = !inQuotes;
    if (inQuotes) continue;

    if (char === ")") depth++;
    else if (char === "(") depth--;

    if (depth === 0) {
      for (const op of ops) {
        if (expr.substring(i, i + op.length) === op) {
          // ensure not part of scientific notation like 1e-5
          if (op === "-" && i > 0 && /[eE]/.test(expr[i - 1])) continue;
          return { op, index: i };
        }
      }
    }
  }
  return null;
}

export function splitArguments(argsStr: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let inQuotes = false;
  let current = "";

  for (let i = 0; i < argsStr.length; i++) {
    const char = argsStr[i];
    if (char === '"') inQuotes = !inQuotes;
    if (!inQuotes) {
      if (char === "(") depth++;
      else if (char === ")") depth--;
      else if (char === "," && depth === 0) {
        args.push(current.trim());
        current = "";
        continue;
      }
    }
    current += char;
  }
  if (current.trim().length > 0) {
    args.push(current.trim());
  }
  return args;
}

export function getRangeValues(rangeStr: string, ctx: EvaluationContext): any[] {
  const rangeMatch = rangeStr.match(/^([A-Z]+[0-9]+)\s*:\s*([A-Z]+[0-9]+)$/i);
  if (rangeMatch) {
    const start = a1ToCoord(rangeMatch[1]);
    const end = a1ToCoord(rangeMatch[2]);
    if (!start || !end) return [];

    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    const values: any[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        values.push(ctx.evalCell(cellKey(r, c), ctx.sheet, ctx.visited, ctx.evaluating));
      }
    }
    return values;
  }

  const val = evaluateExpression(rangeStr, ctx);
  return Array.isArray(val) ? val : [val];
}
