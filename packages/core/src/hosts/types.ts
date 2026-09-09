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

export function addressToBounds(address: string): { r1: number; c1: number; r2: number; c2: number } {
  const cleaned = address.replace(/\$/g, "").split(":")[0] ?? "A1";
  const end = address.replace(/\$/g, "").split(":")[1] ?? cleaned;
  const a = cell(cleaned);
  const b = cell(end);
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c)
  };
}

function cell(ref: string): { r: number; c: number } {
  const m = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return { r: 0, c: 0 };
  const letters = m[1].toUpperCase();
  let c = 0;
  for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: c - 1 };
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
  const { r1, c1, r2, c2 } = addressToBounds(address);
  const out: unknown[][] = [];
  for (let r = r1; r <= r2; r++) {
    const row: unknown[] = [];
    for (let c = c1; c <= c2; c++) row.push(grid[r]?.[c] ?? "");
    out.push(row);
  }
  return out;
}
