import { CellValue } from "../../models/cell";

export interface FormulaFunctionContext {
  getRangeValues(arg: string): CellValue[];
  evaluateExpr(expr: string): CellValue;
}


export function sum(args: string[], ctx: FormulaFunctionContext): number {
  let total = 0;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      const num = Number(v);
      if (!isNaN(num)) total += num;
    }
  }
  return total;
}

export function average(args: string[], ctx: FormulaFunctionContext): number {
  let total = 0;
  let count = 0;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      const num = Number(v);
      if (!isNaN(num) && v !== null && v !== "") {
        total += num;
        count++;
      }
    }
  }
  return count === 0 ? 0 : total / count;
}

export function count(args: string[], ctx: FormulaFunctionContext): number {
  let count = 0;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      if (typeof v === "number" || (!isNaN(Number(v)) && v !== "" && v !== null)) count++;
    }
  }
  return count;
}

export function counta(args: string[], ctx: FormulaFunctionContext): number {
  let count = 0;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      if (v !== undefined && v !== null && v !== "") count++;
    }
  }
  return count;
}

export function min(args: string[], ctx: FormulaFunctionContext): number {
  let minVal = Infinity;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      const num = Number(v);
      if (!isNaN(num) && v !== "" && v !== null) {
        if (num < minVal) minVal = num;
      }
    }
  }
  return minVal === Infinity ? 0 : minVal;
}

export function max(args: string[], ctx: FormulaFunctionContext): number {
  let maxVal = -Infinity;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      const num = Number(v);
      if (!isNaN(num) && v !== "" && v !== null) {
        if (num > maxVal) maxVal = num;
      }
    }
  }
  return maxVal === -Infinity ? 0 : maxVal;
}

export function product(args: string[], ctx: FormulaFunctionContext): number {
  let prod = 1;
  let hasVals = false;
  for (const arg of args) {
    const vals = ctx.getRangeValues(arg);
    for (const v of vals) {
      const num = Number(v);
      if (!isNaN(num)) {
        prod *= num;
        hasVals = true;
      }
    }
  }
  return hasVals ? prod : 0;
}

export function round(args: string[], ctx: FormulaFunctionContext): number {
  const val = Number(ctx.evaluateExpr(args[0]));
  const digits = args[1] ? Number(ctx.evaluateExpr(args[1])) : 0;
  const factor = Math.pow(10, digits);
  return Math.round(val * factor) / factor;
}

export function sqrt(args: string[], ctx: FormulaFunctionContext): number {
  const val = Number(ctx.evaluateExpr(args[0]));
  return Math.sqrt(val);
}

export function abs(args: string[], ctx: FormulaFunctionContext): number {
  const val = Number(ctx.evaluateExpr(args[0]));
  return Math.abs(val);
}
