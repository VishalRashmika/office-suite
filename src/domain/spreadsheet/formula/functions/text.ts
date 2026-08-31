import { FormulaFunctionContext } from "./math";

export function concat(args: string[], ctx: FormulaFunctionContext): string {
  let str = "";
  for (const arg of args) {
    const res = ctx.evaluateExpr(arg);
    str += res !== null && res !== undefined ? String(res) : "";
  }
  return str;
}

export function textJoin(args: string[], ctx: FormulaFunctionContext): string {
  const delimiter = String(ctx.evaluateExpr(args[0]) ?? "");
  const ignoreEmpty = Boolean(ctx.evaluateExpr(args[1]));
  const items: string[] = [];
  for (let i = 2; i < args.length; i++) {
    const vals = ctx.getRangeValues(args[i]);
    for (const v of vals) {
      if (ignoreEmpty && (v === "" || v === null || v === undefined)) continue;
      items.push(v !== null && v !== undefined ? String(v) : "");
    }
  }
  return items.join(delimiter);
}

export function left(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  const count = args[1] ? Number(ctx.evaluateExpr(args[1])) : 1;
  return str.slice(0, count);
}

export function right(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  const count = args[1] ? Number(ctx.evaluateExpr(args[1])) : 1;
  return str.slice(-count);
}

export function mid(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  const start = Math.max(0, Number(ctx.evaluateExpr(args[1])) - 1);
  const length = Number(ctx.evaluateExpr(args[2]));
  return str.substring(start, start + length);
}

export function len(args: string[], ctx: FormulaFunctionContext): number {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  return str.length;
}

export function trim(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  return str.trim();
}

export function upper(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  return str.toUpperCase();
}

export function lower(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  return str.toLowerCase();
}

export function proper(args: string[], ctx: FormulaFunctionContext): string {
  const str = String(ctx.evaluateExpr(args[0]) ?? "");
  return str.replace(/\b\w/g, (c) => c.toUpperCase());
}
