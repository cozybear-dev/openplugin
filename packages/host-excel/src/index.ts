import {
  addressToBounds,
  Changeset,
  excelMatrixAssign,
  type HostAdapter,
  type RawFacts,
  truncateGrid
} from "@openplugin/core";

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

  async search(args: { query: string; sheet?: string }) {
    return withExcel(async (context) => {
      const hits: Array<{ sheet: string; address: string; value: unknown }> = [];
      const sheets = args.sheet
        ? [context.workbook.worksheets.getItem(args.sheet)]
        : context.workbook.worksheets.items;
      if (!args.sheet) {
        context.workbook.worksheets.load("items/name");
        await context.sync();
      }
      const needle = args.query.toLowerCase();
      for (const sheet of args.sheet ? sheets : context.workbook.worksheets.items) {
        const used = sheet.getUsedRangeOrNullObject();
        used.load(["values", "address", "isNullObject"]);
        await context.sync();
        if (used.isNullObject) continue;
        const values = (used.values as unknown[][]) ?? [];
        values.slice(0, 200).forEach((row, r) => {
          row.slice(0, 40).forEach((value, c) => {
            if (String(value ?? "").toLowerCase().includes(needle)) {
              hits.push({
                sheet: sheet.name,
                address: used.address.split("!")[1] ?? `${r + 1}:${c + 1}`,
                value
              });
            }
          });
        });
      }
      return hits.slice(0, 50);
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await withExcel(async (context) => {
      for (const change of changeset.forHost("excel")) {
        if (change.op === "writeRange") {
          const { address, matrix } = excelMatrixAssign(change.address, change.values);
          if (matrix.length && matrix[0].length) {
            context.workbook.worksheets.getItem(change.sheet).getRange(address).values = matrix;
          }
        } else if (change.op === "setFormulas") {
          const { address, matrix } = excelMatrixAssign(change.address, change.formulas);
          if (matrix.length && matrix[0].length) {
            context.workbook.worksheets.getItem(change.sheet).getRange(address).formulas =
              matrix as string[][];
          }
        } else if (change.op === "createTable") {
          context.workbook.tables.add(`${change.sheet}!${change.address}`, true);
        } else if (change.op === "createChart") {
          const sheet = context.workbook.worksheets.getItem(change.sheet);
          const type = mapChart(change.chartType);
          sheet.charts.add(type, sheet.getRange(change.source));
        } else if (change.op === "formatRange") {
          const range = context.workbook.worksheets.getItem(change.sheet).getRange(change.address);
          if (change.bold) range.format.font.bold = true;
          if (change.numberFormat) {
            const { r1, c1, r2, c2 } = addressToBounds(change.address);
            const rows = r2 - r1 + 1;
            const cols = c2 - c1 + 1;
            range.numberFormat = Array.from({ length: rows }, () =>
              Array.from({ length: cols }, () => change.numberFormat as string)
            );
          }
        } else if (change.op === "clearRange") {
          const range = context.workbook.worksheets.getItem(change.sheet).getRange(change.address);
          if (change.clearType === "formats") range.clear(Excel.ClearApplyTo.formats);
          else if (change.clearType === "all") range.clear(Excel.ClearApplyTo.all);
          else range.clear(Excel.ClearApplyTo.contents);
        } else if (change.op === "copyRange") {
          const sheet = context.workbook.worksheets.getItem(change.sheet);
          sheet.getRange(change.dest).copyFrom(sheet.getRange(change.source));
        } else if (change.op === "modifySheet") {
          applySheetStructure(context, change);
        } else if (change.op === "modifyWorkbook") {
          applyWorkbookStructure(context, change);
        } else if (change.op === "resizeRange") {
          const range = context.workbook.worksheets.getItem(change.sheet).getRange(change.address);
          if (change.columnWidth != null) range.format.columnWidth = change.columnWidth;
          if (change.rowHeight != null) range.format.rowHeight = change.rowHeight;
        } else if (change.op === "createPivot") {
          const sheet = context.workbook.worksheets.getItem(change.sheet);
          sheet.pivotTables.add("Pivot", sheet.getRange(change.source), sheet.getRange(change.dest));
        }
      }
      await context.sync();
    });
  }
}

function applySheetStructure(
  context: Excel.RequestContext,
  change: {
    sheet: string;
    operation: string;
    dimension?: string;
    reference?: string;
    count?: number;
  }
): void {
  const sheet = context.workbook.worksheets.getItem(change.sheet);
  const count = change.count ?? 1;
  if (change.operation === "freeze") {
    sheet.freezePanes.freezeRows(count);
    return;
  }
  if (change.operation === "unfreeze") {
    sheet.freezePanes.unfreeze();
    return;
  }
  if (!change.reference) return;
  if (change.dimension === "columns") {
    const range = sheet.getRange(`${change.reference}:${change.reference}`);
    if (change.operation === "insert") range.insert(Excel.InsertShiftDirection.right);
    else if (change.operation === "delete") range.delete(Excel.DeleteShiftDirection.left);
    else if (change.operation === "hide") range.columnHidden = true;
    else if (change.operation === "unhide") range.columnHidden = false;
  } else {
    const range = sheet.getRange(`${change.reference}:${change.reference}`);
    if (change.operation === "insert") range.insert(Excel.InsertShiftDirection.down);
    else if (change.operation === "delete") range.delete(Excel.DeleteShiftDirection.up);
    else if (change.operation === "hide") range.rowHidden = true;
    else if (change.operation === "unhide") range.rowHidden = false;
  }
}

function applyWorkbookStructure(
  context: Excel.RequestContext,
  change: { operation: string; sheet?: string; newName?: string }
): void {
  if (change.operation === "create") context.workbook.worksheets.add(change.newName ?? change.sheet);
  else if (change.operation === "delete" && change.sheet) context.workbook.worksheets.getItem(change.sheet).delete();
  else if (change.operation === "rename" && change.sheet && change.newName) {
    context.workbook.worksheets.getItem(change.sheet).name = change.newName;
  } else if (change.operation === "duplicate" && change.sheet) {
    context.workbook.worksheets.getItem(change.sheet).copy(Excel.WorksheetPositionType.end);
  }
}

function mapChart(type: string): Excel.ChartType {
  const t = type.toLowerCase();
  if (t.includes("line")) return Excel.ChartType.line;
  if (t.includes("pie")) return Excel.ChartType.pie;
  if (t.includes("bar")) return Excel.ChartType.barClustered;
  return Excel.ChartType.columnClustered;
}

export { sliceGrid, writeIntoGrid } from "@openplugin/core";
