import { describe, expect, it } from "vitest";
import { addressForGrid, excelMatrixAssign, normalizeGrid } from "../src/hosts/types.js";

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
});
