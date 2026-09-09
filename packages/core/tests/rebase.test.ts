import { describe, expect, it } from "vitest";
import { rebaseExcelChanges } from "../src/tools/rebase.js";
import type { Change } from "../src/tools/changeset.js";

const write = (sheet: string, address: string): Change => ({
  host: "excel",
  op: "writeRange",
  sheet,
  address,
  values: [["x"]]
});

const format = (sheet: string, address: string): Change => ({
  host: "excel",
  op: "formatRange",
  sheet,
  address,
  numberFormat: "0.00"
});

const insertRows = (sheet: string, reference: string, count = 1): Change => ({
  host: "excel",
  op: "modifySheet",
  sheet,
  operation: "insert",
  dimension: "rows",
  reference,
  count
});

const insertCols = (sheet: string, reference: string, count = 1): Change => ({
  host: "excel",
  op: "modifySheet",
  sheet,
  operation: "insert",
  dimension: "columns",
  reference,
  count
});

const deleteRows = (sheet: string, reference: string, count = 1): Change => ({
  host: "excel",
  op: "modifySheet",
  sheet,
  operation: "delete",
  dimension: "rows",
  reference,
  count
});

describe("rebaseExcelChanges", () => {
  it("shifts a later format down after inserting a row above it", () => {
    const out = rebaseExcelChanges([insertRows("Sheet1", "1"), format("Sheet1", "A1")]);
    expect(out[1]).toMatchObject({ op: "formatRange", address: "A2" });
  });

  it("shifts a later address right after inserting a column", () => {
    const out = rebaseExcelChanges([insertCols("Sheet1", "B"), format("Sheet1", "C1")]);
    expect(out[1]).toMatchObject({ op: "formatRange", address: "D1" });
  });

  it("does not shift a later format when an earlier write overlays cells", () => {
    const out = rebaseExcelChanges([write("Sheet1", "A1:C10"), format("Sheet1", "A1:C10")]);
    expect(out[1]).toMatchObject({ op: "formatRange", address: "A1:C10" });
  });

  it("rewrites later sheet names after a rename", () => {
    const out = rebaseExcelChanges([
      {
        host: "excel",
        op: "modifyWorkbook",
        operation: "rename",
        sheet: "Sheet1",
        newName: "Data"
      },
      write("Sheet1", "A1")
    ]);
    expect(out[1]).toMatchObject({ op: "writeRange", sheet: "Data", address: "A1" });
  });

  it("drops later ops that target a deleted sheet", () => {
    const out = rebaseExcelChanges([
      { host: "excel", op: "modifyWorkbook", operation: "delete", sheet: "Sheet1" },
      format("Sheet1", "A1"),
      write("Sheet2", "A1")
    ]);
    expect(out.map((c) => c.op)).toEqual(["modifyWorkbook", "writeRange"]);
    expect(out[1]).toMatchObject({ sheet: "Sheet2" });
  });

  it("expands a range when rows are inserted inside it", () => {
    const out = rebaseExcelChanges([insertRows("Sheet1", "5"), format("Sheet1", "A1:C10")]);
    expect(out[1]).toMatchObject({ address: "A1:C11" });
  });

  it("drops a later op whose cells were deleted", () => {
    const out = rebaseExcelChanges([deleteRows("Sheet1", "1"), format("Sheet1", "A1")]);
    expect(out).toHaveLength(1);
    expect(out[0]).toMatchObject({ op: "modifySheet" });
  });

  it("leaves non-excel ops in place", () => {
    const word: Change = { host: "word", op: "insertComment", text: "n" };
    const out = rebaseExcelChanges([word, insertRows("Sheet1", "1"), format("Sheet1", "A1")]);
    expect(out[0]).toEqual(word);
    expect(out[2]).toMatchObject({ address: "A2" });
  });

  it("shifts copy source and dest together", () => {
    const out = rebaseExcelChanges([
      insertRows("Sheet1", "1"),
      { host: "excel", op: "copyRange", sheet: "Sheet1", source: "A1", dest: "B1" }
    ]);
    expect(out[1]).toMatchObject({ source: "A2", dest: "B2" });
  });
});
