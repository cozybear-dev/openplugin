import type { Changeset } from "../tools/changeset.js";
import {
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

  async readRange(args: { sheet?: string; address: string }) {
    const sheet = this.sheets[args.sheet ?? this.selection.sheet];
    if (!sheet) throw new Error(`Unknown sheet: ${args.sheet}`);
    return truncateGrid(sliceGrid(sheet.values, args.address));
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
      }
    }
  }

  private ensureSheet(name: string): ExcelSheet {
    this.sheets[name] ??= { name, values: [], tables: [], charts: [] };
    return this.sheets[name];
  }
}
