// Small helpers for translating between spreadsheet A1 notation and 0-based
// {row, col} coordinates used by HyperFormula.

export interface CellCoord {
  row: number; // 0-based
  col: number; // 0-based
}

export interface CellRange {
  start: CellCoord;
  end: CellCoord; // inclusive
}

export function columnLetterToIndex(letters: string): number {
  const upper = letters.toUpperCase();
  let n = 0;
  for (const ch of upper) {
    const code = ch.charCodeAt(0) - 64; // 'A' -> 1
    if (code < 1 || code > 26) {
      throw new Error(`Invalid column letters: "${letters}"`);
    }
    n = n * 26 + code;
  }
  return n - 1;
}

export function indexToColumnLetter(index: number): string {
  if (index < 0 || !Number.isInteger(index)) {
    throw new Error(`Invalid column index: ${index}`);
  }
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

const A1_RE = /^([A-Za-z]+)(\d+)$/;

export function parseA1(addr: string): CellCoord {
  const m = A1_RE.exec(addr.trim());
  if (!m) {
    throw new Error(`Invalid A1 address: "${addr}"`);
  }
  const col = columnLetterToIndex(m[1]!);
  const row = Number.parseInt(m[2]!, 10) - 1;
  if (row < 0) {
    throw new Error(`Invalid row in A1 address: "${addr}"`);
  }
  return { row, col };
}

export function formatA1(coord: CellCoord): string {
  return `${indexToColumnLetter(coord.col)}${coord.row + 1}`;
}

export function parseRange(range: string): CellRange {
  const parts = range.split(':');
  if (parts.length === 1) {
    const c = parseA1(parts[0]!);
    return { start: c, end: c };
  }
  if (parts.length !== 2) {
    throw new Error(`Invalid range: "${range}"`);
  }
  const start = parseA1(parts[0]!);
  const end = parseA1(parts[1]!);
  // Normalize so start <= end on both axes.
  return {
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  };
}

export function formatRange(range: CellRange): string {
  return `${formatA1(range.start)}:${formatA1(range.end)}`;
}

export function rangeSize(range: CellRange): { rows: number; cols: number; cells: number } {
  const rows = range.end.row - range.start.row + 1;
  const cols = range.end.col - range.start.col + 1;
  return { rows, cols, cells: rows * cols };
}
