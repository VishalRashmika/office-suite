import { FormulaFunctionContext, sum, average, count, counta, min, max, product, round, sqrt, abs } from "./math";
import { ifFunc, ifError, andFunc, orFunc, notFunc } from "./logical";
import { concat, textJoin, left, right, mid, len, trim, upper, lower, proper } from "./text";
import { today, now } from "./date";

export type FormulaFunction = (args: string[], ctx: FormulaFunctionContext) => unknown;

export class FunctionRegistry {
  private functions = new Map<string, FormulaFunction>();

  constructor() {
    this.registerDefaults();
  }

  register(name: string, fn: FormulaFunction): void {
    this.functions.set(name.toUpperCase(), fn);
  }

  get(name: string): FormulaFunction | undefined {
    return this.functions.get(name.toUpperCase());
  }

  has(name: string): boolean {
    return this.functions.has(name.toUpperCase());
  }

  private registerDefaults(): void {
    // Math
    this.register("SUM", sum);
    this.register("AVERAGE", average);
    this.register("COUNT", count);
    this.register("COUNTA", counta);
    this.register("MIN", min);
    this.register("MAX", max);
    this.register("PRODUCT", product);
    this.register("ROUND", round);
    this.register("SQRT", sqrt);
    this.register("ABS", abs);

    // Logical
    this.register("IF", ifFunc);
    this.register("IFERROR", ifError);
    this.register("AND", andFunc);
    this.register("OR", orFunc);
    this.register("NOT", notFunc);

    // Text
    this.register("CONCAT", concat);
    this.register("CONCATENATE", concat);
    this.register("TEXTJOIN", textJoin);
    this.register("LEFT", left);
    this.register("RIGHT", right);
    this.register("MID", mid);
    this.register("LEN", len);
    this.register("TRIM", trim);
    this.register("UPPER", upper);
    this.register("LOWER", lower);
    this.register("PROPER", proper);

    // Date
    this.register("TODAY", () => today());
    this.register("NOW", () => now());
  }
}

export const defaultFunctionRegistry = new FunctionRegistry();
