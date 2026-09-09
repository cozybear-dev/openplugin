import type { HostKind } from "../llm/types.js";
import type { Changeset } from "../tools/changeset.js";

export type ExcelSheet = {
  name: string;
  values: unknown[][];
  formulas?: string[][];
  tables?: Array<{ name: string; address: string }>;
  charts?: Array<{ name: string }>;
};

export type ExcelSelection = {
  sheet: string;
  address: string;
  values: unknown[][];
};

export type WordParagraph = { text: string; style: string };

export type WordSelection = { text: string; style: string; paragraphIndex: number };

export type PptShape = { name: string; text: string };

export type PptSlide = { title: string; shapes: PptShape[]; notes: string };

export type PptSelection = { slideIndex: number; shapeName?: string; text: string };

export type RawFacts =
  | {
      host: "excel";
      title: string;
      sheets: Array<{ name: string; rows: number; cols: number; tables: string[] }>;
      namedRanges?: string[];
      selection: ExcelSelection;
    }
  | {
      host: "word";
      title: string;
      headings: Array<{ text: string; style: string; index: number }>;
      paragraphCount: number;
      selection: WordSelection;
      surrounding: string[];
    }
  | {
      host: "powerpoint";
      title: string;
      slides: Array<{ index: number; title: string; shapeCount: number }>;
      selection: PptSelection;
      currentNotes?: string;
    };

export interface HostAdapter {
  kind: HostKind;
  getRawFacts(): Promise<RawFacts>;
  apply(changeset: Changeset): Promise<void>;
  readRange?(args: { sheet?: string; address: string }): Promise<{
    values: unknown[][];
    truncated: boolean;
  }>;
  readSelectionText?(): Promise<string>;
  readShapeText?(args: { slideIndex: number; shapeName?: string }): Promise<string>;
  readNotes?(slideIndex: number): Promise<string>;
  search?(args: { query: string; sheet?: string }): Promise<Array<{ sheet: string; address: string; value: unknown }>>;
  readParagraphs?(args?: { start?: number; count?: number }): Promise<Array<{ index: number; text: string; style: string }>>;
  findText?(args: { query: string; max?: number }): Promise<Array<{ paragraphIndex: number; text: string }>>;
  listComments?(): Promise<Array<{ index: number; text: string; author?: string; resolved?: boolean }>>;
  getRevisions?(): Promise<Array<{ type: string; text: string }>>;
  readSlide?(slideIndex: number): Promise<{
    title: string;
    shapes: Array<{ name: string; text: string }>;
    notes: string;
  }>;
  listLayouts?(): Promise<string[]>;
}

export const READ_RANGE_MAX_ROWS = 200;
export const READ_RANGE_MAX_COLS = 20;

export function truncateGrid(values: unknown[][], maxRows = READ_RANGE_MAX_ROWS, maxCols = READ_RANGE_MAX_COLS): {
  values: unknown[][];
  truncated: boolean;
} {
  const truncated = values.length > maxRows || values.some((row) => row.length > maxCols);
  return {
    truncated,
    values: values.slice(0, maxRows).map((row) => row.slice(0, maxCols))
  };
}

export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;
export const FORMAT_CELL_CAP = 50_000;

export type ExcelBounds = {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  entireColumn?: boolean;
  entireRow?: boolean;
};

export function canonicalizeExcelAddress(address: string): string {
  return boundsToAddress(parseLocalAddress(stripSheetQualifier(address)));
}

export function cellAddressInUsedRange(usedAddress: string, rowOffset: number, colOffset: number): string {
  const origin = addressToBounds(canonicalizeExcelAddress(usedAddress));
  return boundsToAddress({
    r1: origin.r1 + rowOffset,
    c1: origin.c1 + colOffset,
    r2: origin.r1 + rowOffset,
    c2: origin.c1 + colOffset
  });
}

export function addressToBounds(address: string): ExcelBounds {
  return parseLocalAddress(stripSheetQualifier(address));
}

export function boundsToAddress(bounds: ExcelBounds): string {
  if (bounds.entireColumn && !bounds.entireRow) {
    return `${colLetter(bounds.c1)}:${colLetter(bounds.c2)}`;
  }
  if (bounds.entireRow && !bounds.entireColumn) {
    return `${bounds.r1 + 1}:${bounds.r2 + 1}`;
  }
  const start = `${colLetter(bounds.c1)}${bounds.r1 + 1}`;
  if (bounds.r1 === bounds.r2 && bounds.c1 === bounds.c2) return start;
  return `${start}:${colLetter(bounds.c2)}${bounds.r2 + 1}`;
}

export function formatApplySpec(
  address: string,
  used?: { rows: number; cols: number }
): { address: string; rows: number; cols: number } {
  const bounds = addressToBounds(address);
  let { r1, c1, r2, c2 } = bounds;
  if (bounds.entireColumn) {
    r1 = 0;
    r2 = Math.max(0, (used?.rows || 1) - 1);
  }
  if (bounds.entireRow) {
    c1 = 0;
    c2 = Math.max(0, (used?.cols || 1) - 1);
  }
  let rows = r2 - r1 + 1;
  const cols = c2 - c1 + 1;
  if (rows * cols > FORMAT_CELL_CAP) {
    rows = Math.max(1, Math.floor(FORMAT_CELL_CAP / Math.max(cols, 1)));
    r2 = r1 + rows - 1;
  }
  return {
    address: boundsToAddress({ r1, c1, r2, c2 }),
    rows,
    cols
  };
}

export function numberFormatMatrix(rows: number, cols: number, format: string): string[][] {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => format));
}

export function sheetQualifiedAddress(sheet: string, address: string): string {
  const local = canonicalizeExcelAddress(address);
  const quoted = /^[A-Za-z0-9_]+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return `${quoted}!${local}`;
}

function stripSheetQualifier(address: string): string {
  const trimmed = address.trim().replace(/\$/g, "");
  const m = trimmed.match(/^(?:'((?:[^']|'')*)'|([^'!]+))!(.*)$/);
  return (m?.[3] ?? trimmed).trim();
}

function parseLocalAddress(local: string): ExcelBounds {
  const upper = local.toUpperCase();
  if (!upper) throw new Error(`Invalid Excel address: ${local}`);
  const colRange = upper.match(/^([A-Z]+):([A-Z]+)$/);
  if (colRange) {
    const c1 = colIndex(colRange[1]);
    const c2 = colIndex(colRange[2]);
    return {
      r1: 0,
      c1: Math.min(c1, c2),
      r2: EXCEL_MAX_ROWS - 1,
      c2: Math.max(c1, c2),
      entireColumn: true
    };
  }
  const rowRange = upper.match(/^(\d+):(\d+)$/);
  if (rowRange) {
    const r1 = Number(rowRange[1]) - 1;
    const r2 = Number(rowRange[2]) - 1;
    return {
      r1: Math.min(r1, r2),
      c1: 0,
      r2: Math.max(r1, r2),
      c2: EXCEL_MAX_COLS - 1,
      entireRow: true
    };
  }
  const parts = upper.split(":");
  if (parts.length > 2) throw new Error(`Invalid Excel address: ${local}`);
  const a = parseCell(parts[0] ?? "");
  const b = parts[1] ? parseCell(parts[1]) : a;
  if (!a || !b) throw new Error(`Invalid Excel address: ${local}`);
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c)
  };
}

function parseCell(ref: string): { r: number; c: number } | null {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  const r = Number(m[2]) - 1;
  if (r < 0) return null;
  return { r, c: colIndex(m[1]) };
}

function colIndex(letters: string): number {
  let c = 0;
  for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

export function colLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function normalizeGrid(raw: unknown): unknown[][] {
  if (raw == null || raw === "") return [];
  if (!Array.isArray(raw)) return [[raw]];
  if (raw.length === 0) return [];
  if (!raw.some((cell) => Array.isArray(cell))) return [raw];
  const rows = raw.map((row) => (Array.isArray(row) ? [...row] : [row]));
  const cols = rows.reduce((m, row) => Math.max(m, row.length), 0);
  return rows.map((row) => {
    while (row.length < cols) row.push("");
    return row;
  });
}

export function addressForGrid(originAddress: string, values: unknown[][]): string {
  const { r1, c1 } = addressToBounds(originAddress);
  const rows = values.length;
  const cols = values[0]?.length ?? 0;
  const origin = `${colLetter(c1)}${r1 + 1}`;
  if (rows === 0 || cols === 0) return origin;
  if (rows === 1 && cols === 1) return origin;
  const end = `${colLetter(c1 + cols - 1)}${r1 + rows}`;
  return `${origin}:${end}`;
}

export function excelMatrixAssign(address: string, raw: unknown): { address: string; matrix: unknown[][] } {
  const matrix = normalizeGrid(raw);
  return { address: addressForGrid(address, matrix), matrix };
}

export function writeIntoGrid(grid: unknown[][], address: string, values: unknown[][]): void {
  const { r1, c1 } = addressToBounds(address);
  for (let r = 0; r < values.length; r++) {
    const row = values[r] ?? [];
    while (grid.length <= r1 + r) grid.push([]);
    const target = grid[r1 + r];
    for (let c = 0; c < row.length; c++) {
      while (target.length <= c1 + c) target.push("");
      target[c1 + c] = row[c];
    }
  }
}

export function sliceGrid(grid: unknown[][], address: string): unknown[][] {
  let { r1, c1, r2, c2, entireColumn, entireRow } = addressToBounds(address);
  if (entireColumn) r2 = Math.min(r2, Math.max(grid.length - 1, r1));
  if (entireRow) {
    const maxCol = grid.reduce((m, row) => Math.max(m, row.length), 0) - 1;
    c2 = Math.min(c2, Math.max(maxCol, c1));
  }
  const out: unknown[][] = [];
  for (let r = r1; r <= r2; r++) {
    const row: unknown[] = [];
    for (let c = c1; c <= c2; c++) row.push(grid[r]?.[c] ?? "");
    out.push(row);
  }
  return out;
}
