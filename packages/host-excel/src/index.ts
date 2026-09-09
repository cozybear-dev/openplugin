import { Changeset, type HostAdapter, type RawFacts, truncateGrid } from "@openplugin/core";

async function withExcel<T>(fn: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  return Excel.run(fn);
}

export class ExcelHost implements HostAdapter {
  readonly kind = "excel" as const;

  async getRawFacts(): Promise<RawFacts> {
    return withExcel(async (context) => {
      const wb = context.workbook;
      const sheets = wb.worksheets;
      sheets.load("items/name");
      const selected = wb.getSelectedRange();
      selected.load(["address", "values", "worksheet/name"]);
      const tables = wb.tables;
      tables.load("items/name");
      await context.sync();

      const used: Array<{ name: string; rows: number; cols: number; tables: string[] }> = [];
      const usedRanges = sheets.items.map((sheet) => {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["rowCount", "columnCount", "isNullObject"]);
        used.push({ name: sheet.name, rows: 0, cols: 0, tables: [] });
        return usedRange;
      });
      await context.sync();

      usedRanges.forEach((usedRange, i) => {
        if (!usedRange.isNullObject) {
          used[i].rows = usedRange.rowCount;
          used[i].cols = usedRange.columnCount;
        }
      });

      const tableNames = tables.items.map((t) => t.name);
      if (used[0]) used[0].tables = tableNames;

      const values = (selected.values as unknown[][]) ?? [[""]];
      return {
        host: "excel",
        title: "Workbook",
        sheets: used,
        selection: {
          sheet: selected.worksheet.name,
          address: selected.address.split("!").pop() ?? selected.address,
          values: truncateGrid(values, 30, 12).values
        }
      };
    });
  }

  async readSelectionText(): Promise<string> {
    return withExcel(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load("values");
      await context.sync();
      return ((selected.values as unknown[][]) ?? []).flat().map((v) => String(v ?? "")).join(" ");
    });
  }

  async readRange(args: { sheet?: string; address: string }) {
    return withExcel(async (context) => {
      const sheet = args.sheet
        ? context.workbook.worksheets.getItem(args.sheet)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.address);
      range.load("values");
      await context.sync();
      return truncateGrid((range.values as unknown[][]) ?? []);
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await withExcel(async (context) => {
      for (const change of changeset.forHost("excel")) {
        if (change.op === "writeRange") {
          context.workbook.worksheets.getItem(change.sheet).getRange(change.address).values =
            change.values as string[][];
        } else if (change.op === "setFormulas") {
          context.workbook.worksheets.getItem(change.sheet).getRange(change.address).formulas =
            change.formulas;
        } else if (change.op === "createTable") {
          context.workbook.tables.add(`${change.sheet}!${change.address}`, true);
        } else if (change.op === "createChart") {
          const sheet = context.workbook.worksheets.getItem(change.sheet);
          const type = mapChart(change.chartType);
          sheet.charts.add(type, sheet.getRange(change.source));
        }
      }
      await context.sync();
    });
  }
}

function mapChart(type: string): Excel.ChartType {
  const t = type.toLowerCase();
  if (t.includes("line")) return Excel.ChartType.line;
  if (t.includes("pie")) return Excel.ChartType.pie;
  if (t.includes("bar")) return Excel.ChartType.barClustered;
  return Excel.ChartType.columnClustered;
}

export { sliceGrid, writeIntoGrid };
