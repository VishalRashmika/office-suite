export type CellValue = string | number | boolean | null | undefined;

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  fontSize?: number; // pt
  fontFamily?: string;
  numberFormat?: string; // "text" | "number" | "currency" | "percent" | "date"
  wrap?: boolean;
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
}

export interface CellData {
  value?: CellValue;
  formula?: string; // e.g. "=SUM(A1:B5)"
  formatted?: string;
  style?: CellStyle;
}

export function formatCellValue(val: CellValue, numFormat?: string): string {
  if (val === undefined || val === null) return "";
  if (typeof val === "boolean") return val ? "TRUE" : "FALSE";

  if (typeof val === "number" && !isNaN(val)) {
    if (numFormat === "currency") {
      return `$${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    if (numFormat === "percent") {
      return `${(val * 100).toFixed(2)}%`;
    }
    return val.toLocaleString();
  }

  return String(val);
}
