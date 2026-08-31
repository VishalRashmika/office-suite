import { FormulaFunctionContext } from "./math";

export function ifFunc(args: string[], ctx: FormulaFunctionContext): any {
  const condition = ctx.evaluateExpr(args[0]);
  if (condition && condition !== "0" && condition !== "#ERROR!") {
    return args[1] !== undefined ? ctx.evaluateExpr(args[1]) : true;
  } else {
    return args[2] !== undefined ? ctx.evaluateExpr(args[2]) : false;
  }
}

export function ifError(args: string[], ctx: FormulaFunctionContext): any {
  try {
    const res = ctx.evaluateExpr(args[0]);
    if (typeof res === "string" && res.startsWith("#")) {
      return args[1] !== undefined ? ctx.evaluateExpr(args[1]) : "";
    }
    return res;
  } catch {
    return args[1] !== undefined ? ctx.evaluateExpr(args[1]) : "";
  }
}

export function andFunc(args: string[], ctx: FormulaFunctionContext): boolean {
  for (const arg of args) {
    const res = ctx.evaluateExpr(arg);
    if (!res) return false;
  }
  return true;
}

export function orFunc(args: string[], ctx: FormulaFunctionContext): boolean {
  for (const arg of args) {
    const res = ctx.evaluateExpr(arg);
    if (res) return true;
  }
  return false;
}

export function notFunc(args: string[], ctx: FormulaFunctionContext): boolean {
  const res = ctx.evaluateExpr(args[0]);
  return !res;
}
