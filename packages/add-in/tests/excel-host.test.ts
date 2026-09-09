import { Changeset } from "@openplugin/core";
import { ExcelHost } from "@openplugin/host-excel";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

describe("ExcelHost duplicate sheet", () => {
  const copied = { name: "Sheet1 (2)" };
  const source = { copy: vi.fn(() => copied) };

  beforeEach(() => {
    copied.name = "Sheet1 (2)";
    source.copy.mockClear();
    (globalThis as { Excel?: unknown }).Excel = {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          workbook: { worksheets: { getItem: () => source } },
          sync: async () => undefined
        }),
      WorksheetPositionType: { end: "end" }
    };
  });

  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("renames the copied sheet when newName is set", async () => {
    const host = new ExcelHost();
    const cs = new Changeset();
    cs.add({
      host: "excel",
      op: "modifyWorkbook",
      operation: "duplicate",
      sheet: "Sheet1",
      newName: "Budget"
    });
    await host.apply(cs);
    expect(source.copy).toHaveBeenCalled();
    expect(copied.name).toBe("Budget");
  });
});

describe("ExcelHost search", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("returns the cell address of a hit, not the used range", async () => {
    const used = {
      load() {},
      isNullObject: false,
      address: "Sheet1!A1:C3",
      values: [
        ["Item", "Qty"],
        ["Apples", 3],
        ["Pears", 2]
      ]
    };
    (globalThis as { Excel?: unknown }).Excel = {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          workbook: {
            worksheets: {
              load: () => undefined,
              items: [{ name: "Sheet1", getUsedRangeOrNullObject: () => used }],
              getItem: () => ({ name: "Sheet1", getUsedRangeOrNullObject: () => used })
            }
          },
          sync: async () => undefined
        })
    };
    const hits = await new ExcelHost().search({ query: "Pears" });
    expect(hits).toEqual([{ sheet: "Sheet1", address: "A3", value: "Pears" }]);
  });
});

describe("ExcelHost selection fallback", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("getRawFacts returns a fallback selection when getSelectedRange throws", async () => {
    const usedRange = { load() {}, isNullObject: true, rowCount: 0, columnCount: 0 };
    (globalThis as { Excel?: unknown }).Excel = {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          workbook: {
            worksheets: {
              load: () => undefined,
              items: [
                {
                  name: "Sheet1",
                  getUsedRangeOrNullObject: () => usedRange
                }
              ]
            },
            tables: { load() {}, items: [] },
            getSelectedRange: () => {
              throw new Error("The current selection is invalid for this operation.");
            }
          },
          sync: async () => undefined
        })
    };
    const facts = await new ExcelHost().getRawFacts();
    expect(facts.host).toBe("excel");
    expect(facts.selection.address).toBe("A1");
    expect(facts.selection.sheet).toBe("Sheet1");
  });

  it("attaches tables to the worksheet that owns them", async () => {
    const usedRange = { load() {}, isNullObject: true, rowCount: 0, columnCount: 0 };
    (globalThis as { Excel?: unknown }).Excel = {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          workbook: {
            worksheets: {
              load: () => undefined,
              items: [
                { name: "Sheet1", getUsedRangeOrNullObject: () => usedRange },
                { name: "Data", getUsedRangeOrNullObject: () => usedRange }
              ]
            },
            tables: {
              load() {},
              items: [{ name: "Sales", worksheet: { name: "Data" } }]
            },
            getSelectedRange: () => {
              throw new Error("The current selection is invalid for this operation.");
            }
          },
          sync: async () => undefined
        })
    };
    const facts = await new ExcelHost().getRawFacts();
    expect(facts.host).toBe("excel");
    expect(facts.sheets.find((s) => s.name === "Data")?.tables).toEqual(["Sales"]);
    expect(facts.sheets.find((s) => s.name === "Sheet1")?.tables).toEqual([]);
  });

  it("createChart apply selects a cell after adding the chart", async () => {
    const select = vi.fn();
    const getRange = vi.fn(() => ({ select, copyFrom() {}, format: {} }));
    const charts = { add: vi.fn() };
    const sheet = { getRange, charts };
    (globalThis as { Excel?: unknown }).Excel = {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          workbook: { worksheets: { getItem: () => sheet } },
          sync: async () => undefined
        }),
      ChartType: { columnClustered: "columnClustered", line: "line", pie: "pie", barClustered: "barClustered" }
    };
    const cs = new Changeset();
    cs.add({ host: "excel", op: "createChart", sheet: "Sheet1", source: "A1:B3", chartType: "column" });
    await new ExcelHost().apply(cs);
    expect(charts.add).toHaveBeenCalled();
    expect(getRange).toHaveBeenCalledWith("A1");
    expect(select).toHaveBeenCalled();
  });
});
