export interface CellCoord {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** Convert (0, 0) -> "A1", (1, 2) -> "C2" */
export function coordToA1(row: number, col: number): string {
  let colStr = "";
  let c = col;
  while (c >= 0) {
    colStr = String.fromCharCode(65 + (c % 26)) + colStr;
    c = Math.floor(c / 26) - 1;
  }
  return `${colStr}${row + 1}`;
}

/** Convert "A1" -> { row: 0, col: 0 }, "C2" -> { row: 1, col: 2 } */
export function a1ToCoord(a1: string): CellCoord | null {
  const match = a1.trim().toUpperCase().match(/^([A-Z]+)([0-9]+)$/);
  if (!match) return null;
  const colLetters = match[1];
  const rowNum = parseInt(match[2], 10) - 1;
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row: rowNum, col: col - 1 };
}

export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

export function parseCellKey(key: string): CellCoord {
  const [r, c] = key.split(",").map((n) => parseInt(n, 10));
  return { row: r, col: c };
}
