import { describe, expect, it } from "vitest";
import {
  addressForGrid,
  addressToBounds,
  boundsToAddress,
  canonicalizeExcelAddress,
  cellAddressInUsedRange,
  excelMatrixAssign,
  formatApplySpec,
  normalizeGrid,
  numberFormatMatrix,
  sheetQualifiedAddress
} from "../src/hosts/types.js";

describe("normalizeGrid", () => {
  it("wraps a 1D row", () => {
    expect(normalizeGrid(["a", "b", "c"])).toEqual([["a", "b", "c"]]);
  });
  it("pads a jagged 2D array", () => {
    expect(normalizeGrid([["a", "b"], ["c"]])).toEqual([["a", "b"], ["c", ""]]);
  });
  it("wraps a scalar", () => {
    expect(normalizeGrid("hello")).toEqual([["hello"]]);
  });
});

describe("cellAddressInUsedRange", () => {
  it("maps a hit inside A1:C3 to A3", () => {
    expect(cellAddressInUsedRange("Sheet1!A1:C3", 2, 0)).toBe("A3");
  });

  it("maps a hit inside B2:D5 to C4", () => {
    expect(cellAddressInUsedRange("B2:D5", 2, 1)).toBe("C4");
  });
});

describe("addressForGrid", () => {
  it("sizes from origin, ignoring a too-large claimed span", () => {
    expect(addressForGrid("A1:C10", [["h1", "h2"], [1, 2]])).toBe("A1:B2");
  });
  it("expands a single cell to the values shape", () => {
    expect(addressForGrid("A1", [["a", "b"], ["c", "d"]])).toBe("A1:B2");
  });
  it("keeps a 1x1 write as a cell", () => {
    expect(addressForGrid("B5:Z99", [["x"]])).toBe("B5");
  });
});

describe("excelMatrixAssign", () => {
  it("empty grid → origin cell", () => {
    expect(excelMatrixAssign("C2:Z99", [])).toEqual({ address: "C2", matrix: [] });
    expect(excelMatrixAssign("A1:C10", null)).toEqual({ address: "A1", matrix: [] });
  });
  it("2×3 at C2 → C2:E3", () => {
    expect(excelMatrixAssign("C2", [["a", "b", "c"], [1, 2, 3]])).toEqual({
      address: "C2:E3",
      matrix: [["a", "b", "c"], [1, 2, 3]]
    });
  });
  it("ignores a too-large claimed span and normalizes jagged input", () => {
    expect(excelMatrixAssign("A1:C10", [["h1", "h2"], ["c"]])).toEqual({
      address: "A1:B2",
      matrix: [["h1", "h2"], ["c", ""]]
    });
  });
  it("wraps a 1D row and sizes from origin", () => {
    expect(excelMatrixAssign("B5:Z99", ["x", "y"])).toEqual({
      address: "B5:C5",
      matrix: [["x", "y"]]
    });
  });
  it("keeps a sheet-qualified origin instead of collapsing to A1", () => {
    expect(excelMatrixAssign("Sheet1!C2", [["a", "b", "c"], [1, 2, 3]])).toEqual({
      address: "C2:E3",
      matrix: [["a", "b", "c"], [1, 2, 3]]
    });
  });
});

describe("canonicalizeExcelAddress", () => {
  it("strips a sheet qualifier, $, and quotes", () => {
    expect(canonicalizeExcelAddress("Sheet1!B2:B10")).toBe("B2:B10");
    expect(canonicalizeExcelAddress("'My Sheet'!A1:C10")).toBe("A1:C10");
    expect(canonicalizeExcelAddress("$B$2:$B$20")).toBe("B2:B20");
    expect(canonicalizeExcelAddress("b2")).toBe("B2");
  });
  it("keeps whole-column and whole-row addresses", () => {
    expect(canonicalizeExcelAddress("C:C")).toBe("C:C");
    expect(canonicalizeExcelAddress("Sheet1!C:C")).toBe("C:C");
    expect(canonicalizeExcelAddress("2:10")).toBe("2:10");
  });
  it("rejects structured references instead of mapping them to A1", () => {
    expect(() => canonicalizeExcelAddress("Table1[Amt]")).toThrow(/structured|named|invalid/i);
  });
});

describe("addressToBounds", () => {
  it("parses a sheet-qualified range as that range, not A1", () => {
    expect(addressToBounds("Sheet1!B2:B10")).toMatchObject({ r1: 1, c1: 1, r2: 9, c2: 1 });
  });
  it("marks whole columns and rows", () => {
    expect(addressToBounds("C:C")).toMatchObject({ c1: 2, c2: 2, entireColumn: true });
    expect(addressToBounds("2:10")).toMatchObject({ r1: 1, r2: 9, entireRow: true });
  });
  it("throws on garbage instead of returning A1", () => {
    expect(() => addressToBounds("Table1[Amt]")).toThrow();
  });
});

describe("boundsToAddress", () => {
  it("round-trips cells, ranges, columns, and rows", () => {
    expect(boundsToAddress(addressToBounds("B2"))).toBe("B2");
    expect(boundsToAddress(addressToBounds("A1:C10"))).toBe("A1:C10");
    expect(boundsToAddress(addressToBounds("C:C"))).toBe("C:C");
    expect(boundsToAddress(addressToBounds("2:10"))).toBe("2:10");
  });
});

describe("formatApplySpec", () => {
  it("keeps a finite A1 range and sizes the number-format matrix from it", () => {
    expect(formatApplySpec("A1:C10")).toEqual({ address: "A1:C10", rows: 10, cols: 3 });
  });
  it("clamps a whole column to the used range instead of a million-row matrix", () => {
    expect(formatApplySpec("C:C", { rows: 20, cols: 6 })).toEqual({
      address: "C1:C20",
      rows: 20,
      cols: 1
    });
  });
  it("falls back to the first cell when the sheet is empty", () => {
    expect(formatApplySpec("C:C")).toEqual({ address: "C1", rows: 1, cols: 1 });
  });
  it("caps a huge finite range so the format matrix stays bounded", () => {
    const spec = formatApplySpec("A1:C1048576");
    expect(spec.rows * spec.cols).toBeLessThanOrEqual(50_000);
    expect(spec.cols).toBe(3);
    expect(spec.address.startsWith("A1:")).toBe(true);
  });
});

describe("numberFormatMatrix", () => {
  it("builds a 2D array matching the range size", () => {
    expect(numberFormatMatrix(2, 3, "0.00")).toEqual([
      ["0.00", "0.00", "0.00"],
      ["0.00", "0.00", "0.00"]
    ]);
  });
});

describe("sheetQualifiedAddress", () => {
  it("quotes sheet names that need it and uses a local A1 address", () => {
    expect(sheetQualifiedAddress("Sheet1", "A1:C10")).toBe("Sheet1!A1:C10");
    expect(sheetQualifiedAddress("My Sheet", "Sheet1!A1")).toBe("'My Sheet'!A1");
  });
});
