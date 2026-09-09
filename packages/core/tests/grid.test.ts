import { describe, expect, it } from "vitest";
import { addressForGrid, normalizeGrid } from "../src/hosts/types.js";

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
