import type { Changeset } from "../tools/changeset.js";
import {
  canonicalizeExcelAddress,
  sliceGrid,
  truncateGrid,
  writeIntoGrid,
  type ExcelSheet,
  type ExcelSelection,
  type HostAdapter,
  type RawFacts
} from "./types.js";

export class FakeExcelHost implements HostAdapter {
  readonly kind = "excel" as const;
  title = "Workbook";
  sheets: Record<string, ExcelSheet> = {
    Sheet1: { name: "Sheet1", values: [["", ""], ["", ""]], tables: [], charts: [] }
  };
  selection: ExcelSelection = { sheet: "Sheet1", address: "A1", values: [[""]] };
  formats: Array<{ sheet: string; address: string; numberFormat?: string; bold?: boolean }> = [];

  async getRawFacts(): Promise<RawFacts> {
    return {
      host: "excel",
      title: this.title,
      sheets: Object.values(this.sheets).map((s) => ({
        name: s.name,
        rows: s.values.length,
        cols: s.values.reduce((m, r) => Math.max(m, r.length), 0),
        tables: (s.tables ?? []).map((t) => t.name)
      })),
      selection: this.selection
    };
  }

  async readSelectionText(): Promise<string> {
    return this.selection.values.flat().map((v) => String(v ?? "")).join(" ");
  }

  async readRange(args: { sheet?: string; address: string }) {
    const sheet = this.sheets[args.sheet ?? this.selection.sheet];
    if (!sheet) throw new Error(`Unknown sheet: ${args.sheet}`);
    return truncateGrid(sliceGrid(sheet.values, args.address));
  }

  async search(args: { query: string; sheet?: string }) {
    const needle = args.query.toLowerCase();
    const hits: Array<{ sheet: string; address: string; value: unknown }> = [];
    const names = args.sheet ? [args.sheet] : Object.keys(this.sheets);
    for (const name of names) {
      const sheet = this.sheets[name];
      if (!sheet) continue;
      sheet.values.forEach((row, r) => {
        row.forEach((value, c) => {
          if (String(value ?? "").toLowerCase().includes(needle)) {
            hits.push({ sheet: name, address: `${col(c)}${r + 1}`, value });
          }
        });
      });
    }
    return hits;
  }

  async apply(changeset: Changeset): Promise<void> {
    for (const change of changeset.forHost("excel")) {
      if (change.op === "writeRange") {
        const sheet = this.ensureSheet(change.sheet);
        writeIntoGrid(sheet.values, change.address, change.values);
      } else if (change.op === "setFormulas") {
        const sheet = this.ensureSheet(change.sheet);
        sheet.formulas ??= [];
        writeIntoGrid(sheet.formulas as unknown as unknown[][], change.address, change.formulas);
      } else if (change.op === "createTable") {
        const sheet = this.ensureSheet(change.sheet);
        sheet.tables ??= [];
        sheet.tables.push({ name: change.name ?? `Table${sheet.tables.length + 1}`, address: change.address });
      } else if (change.op === "createChart") {
        const sheet = this.ensureSheet(change.sheet);
        sheet.charts ??= [];
        sheet.charts.push({ name: `${change.chartType}-${sheet.charts.length + 1}` });
      } else if (change.op === "formatRange") {
        this.formats.push({
          sheet: change.sheet,
          address: canonicalizeExcelAddress(change.address),
          numberFormat: change.numberFormat,
          bold: change.bold
        });
      } else if (change.op === "modifySheet") {
        applyFakeSheetStructure(this.ensureSheet(change.sheet), change);
      } else if (change.op === "clearRange") {
        const sheet = this.ensureSheet(change.sheet);
        const grid = sliceGrid(sheet.values, change.address);
        writeIntoGrid(
          sheet.values,
          change.address,
          grid.map((row) => row.map(() => ""))
        );
      } else if (change.op === "copyRange") {
        const sheet = this.ensureSheet(change.sheet);
        writeIntoGrid(sheet.values, change.dest, sliceGrid(sheet.values, change.source));
      } else if (change.op === "modifyWorkbook") {
        if (change.operation === "create") this.ensureSheet(change.newName ?? change.sheet ?? "Sheet");
        if (change.operation === "delete" && change.sheet) delete this.sheets[change.sheet];
        if (change.operation === "rename" && change.sheet && change.newName) {
          const existing = this.sheets[change.sheet];
          if (existing) {
            existing.name = change.newName;
            this.sheets[change.newName] = existing;
            delete this.sheets[change.sheet];
          }
        }
        if (change.operation === "duplicate" && change.sheet) {
          const existing = this.sheets[change.sheet];
          if (existing) {
            const name = change.newName ?? `${existing.name} Copy`;
            this.sheets[name] = { ...existing, name, values: existing.values.map((r) => [...r]) };
          }
        }
      }
    }
  }

  private ensureSheet(name: string): ExcelSheet {
    this.sheets[name] ??= { name, values: [], tables: [], charts: [] };
    return this.sheets[name];
  }
}

function applyFakeSheetStructure(
  sheet: ExcelSheet,
  change: { operation: string; dimension?: string; reference?: string; count?: number }
): void {
  if (change.operation !== "insert" && change.operation !== "delete") return;
  const count = change.count ?? 1;
  if (change.dimension === "columns") {
    const at = colIndexFromRef(change.reference);
    for (const row of sheet.values) {
      if (change.operation === "insert") row.splice(at, 0, ...Array.from({ length: count }, () => ""));
      else row.splice(at, count);
    }
    return;
  }
  const at = rowIndexFromRef(change.reference);
  if (change.operation === "insert") {
    const width = sheet.values.reduce((m, row) => Math.max(m, row.length), 0);
    const inserted = Array.from({ length: count }, () => Array.from({ length: width }, () => ""));
    sheet.values.splice(at, 0, ...inserted);
  } else {
    sheet.values.splice(at, count);
  }
}

function rowIndexFromRef(reference?: string): number {
  const n = Number((reference ?? "1").match(/\d+/)?.[0] ?? "1");
  return Math.max(0, n - 1);
}

function colIndexFromRef(reference?: string): number {
  const letters = ((reference ?? "A").match(/[A-Za-z]+/)?.[0] ?? "A").toUpperCase();
  let c = 0;
  for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
  return Math.max(0, c - 1);
}

function col(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
