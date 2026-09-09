import {
  addressToBounds,
  boundsToAddress,
  colLetter,
  type ExcelBounds
} from "../hosts/types.js";
import type { Change, ExcelChange } from "./changeset.js";

type Shift = { at: number; count: number; op: "insert" | "delete" };

type SheetState = {
  name: string;
  deleted: boolean;
  rows: Shift[];
  cols: Shift[];
};

export function rebaseExcelChanges(changes: Change[]): Change[] {
  const sheets = new Map<string, SheetState>();
  const aliases = new Map<string, string>();
  const out: Change[] = [];

  const state = (name: string): SheetState => {
    const current = resolveName(name, aliases);
    let s = sheets.get(current);
    if (!s) {
      s = { name: current, deleted: false, rows: [], cols: [] };
      sheets.set(current, s);
    }
    return s;
  };

  for (const change of changes) {
    if (change.host !== "excel") {
      out.push(change);
      continue;
    }
    const next = rewriteExcelChange(change, aliases, state);
    if (!next) continue;
    out.push(next);
    absorb(next, aliases, sheets, state);
  }
  return out;
}

function resolveName(name: string, aliases: Map<string, string>): string {
  let current = name;
  const seen = new Set<string>();
  while (aliases.has(current) && !seen.has(current)) {
    seen.add(current);
    current = aliases.get(current)!;
  }
  return current;
}

function rewriteExcelChange(
  change: ExcelChange,
  aliases: Map<string, string>,
  state: (name: string) => SheetState
): Change | null {
  const sheetName = "sheet" in change && change.sheet ? change.sheet : undefined;
  if (sheetName) {
    const resolved = resolveName(sheetName, aliases);
    const st = state(resolved);
    if (st.deleted) return null;
    return rewriteAddresses({ ...change, sheet: resolved } as ExcelChange, st);
  }
  if (change.op === "modifyWorkbook") {
    return change;
  }
  return change;
}

function rewriteAddresses(change: ExcelChange, st: SheetState): Change | null {
  const shift = (address: string): string | null => {
    try {
      const bounds = addressToBounds(address);
      const next = applyShifts(bounds, st);
      if (!next) return null;
      return boundsToAddress(next);
    } catch {
      return address;
    }
  };

  switch (change.op) {
    case "writeRange":
    case "setFormulas":
    case "createTable":
    case "formatRange":
    case "clearRange":
    case "resizeRange": {
      const address = shift(change.address);
      if (!address) return null;
      return { ...change, address };
    }
    case "createChart": {
      const source = shift(change.source);
      if (!source) return null;
      return { ...change, source };
    }
    case "copyRange": {
      const source = shift(change.source);
      const dest = shift(change.dest);
      if (!source || !dest) return null;
      return { ...change, source, dest };
    }
    case "createPivot": {
      const source = shift(change.source);
      const dest = shift(change.dest);
      if (!source || !dest) return null;
      return { ...change, source, dest };
    }
    case "modifySheet": {
      if (!change.reference || !change.dimension) return change;
      const shifted = shiftReference(change.reference, change.dimension, st);
      if (shifted == null) return null;
      return { ...change, reference: shifted };
    }
    default:
      return change;
  }
}

function applyShifts(bounds: ExcelBounds, st: SheetState): ExcelBounds | null {
  let b: ExcelBounds = { ...bounds };
  if (!b.entireColumn) {
    for (const s of st.rows) {
      const iv = shiftInterval(b.r1, b.r2, s.at, s.count, s.op);
      if (!iv) return null;
      b = { ...b, r1: iv.start, r2: iv.end };
    }
  }
  if (!b.entireRow) {
    for (const s of st.cols) {
      const iv = shiftInterval(b.c1, b.c2, s.at, s.count, s.op);
      if (!iv) return null;
      b = { ...b, c1: iv.start, c2: iv.end };
    }
  }
  return b;
}

function shiftInterval(
  start: number,
  end: number,
  at: number,
  count: number,
  op: "insert" | "delete"
): { start: number; end: number } | null {
  if (op === "insert") {
    if (end < at) return { start, end };
    if (start >= at) return { start: start + count, end: end + count };
    return { start, end: end + count };
  }
  const delEnd = at + count;
  if (end < at) return { start, end };
  if (start >= delEnd) return { start: start - count, end: end - count };
  if (start >= at && end < delEnd) return null;
  const nextStart = start < at ? start : at;
  const nextEnd = end >= delEnd ? end - count : at - 1;
  if (nextEnd < nextStart) return null;
  return { start: nextStart, end: nextEnd };
}

function shiftReference(reference: string, dimension: "rows" | "columns", st: SheetState): string | null {
  const at = parseReferenceIndex(reference, dimension);
  if (at == null) return reference;
  const dummy: ExcelBounds =
    dimension === "rows"
      ? { r1: at, r2: at, c1: 0, c2: 0 }
      : { r1: 0, r2: 0, c1: at, c2: at };
  const shifted = applyShifts(dummy, st);
  if (!shifted) return null;
  return dimension === "rows" ? String(shifted.r1 + 1) : colLetter(shifted.c1);
}

function parseReferenceIndex(reference: string, dimension: "rows" | "columns"): number | null {
  const local = reference.trim().replace(/\$/g, "").split("!").pop()?.toUpperCase() ?? "";
  if (dimension === "rows") {
    const m = local.match(/(\d+)/);
    if (!m) return null;
    return Number(m[1]) - 1;
  }
  const m = local.match(/^([A-Z]+)/);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

function absorb(
  change: Change,
  aliases: Map<string, string>,
  sheets: Map<string, SheetState>,
  state: (name: string) => SheetState
): void {
  if (change.host !== "excel") return;
  if (change.op === "modifyWorkbook") {
    if (change.operation === "rename" && change.sheet && change.newName) {
      const from = resolveName(change.sheet, aliases);
      let st = sheets.get(from);
      if (!st) {
        st = { name: from, deleted: false, rows: [], cols: [] };
        sheets.set(from, st);
      }
      aliases.set(change.sheet, change.newName);
      aliases.set(from, change.newName);
      st.name = change.newName;
      sheets.set(change.newName, st);
    } else if (change.operation === "delete" && change.sheet) {
      const st = state(change.sheet);
      st.deleted = true;
    }
    return;
  }
  if (change.op !== "modifySheet") return;
  if (change.operation !== "insert" && change.operation !== "delete") return;
  if (!change.dimension || !change.reference) return;
  const st = state(change.sheet);
  const at = parseReferenceIndex(change.reference, change.dimension);
  if (at == null) return;
  const shift: Shift = { at, count: change.count ?? 1, op: change.operation };
  if (change.dimension === "rows") st.rows.push(shift);
  else st.cols.push(shift);
}
