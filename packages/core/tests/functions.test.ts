import { describe, expect, it } from "vitest";
import {
  buildMapCall,
  buildTranslateCall,
  cacheKey,
  chunkRows,
  flattenCells,
  parseExtract,
  parseMappedArray,
  reshapeCells,
  serializeRange
} from "../src/functions/excel-fn.js";

describe("excel custom function helpers", () => {
  it("chunks rows into groups of 25", () => {
    const rows = Array.from({ length: 60 }, (_, i) => [i]);
    const chunks = chunkRows(rows, 25);
    expect(chunks).toHaveLength(3);
    expect(chunks[0]).toHaveLength(25);
    expect(chunks[2]).toHaveLength(10);
  });

  it("serializes a range as TSV", () => {
    expect(serializeRange([["a", "b"], [1, 2]])).toBe("a\tb\n1\t2");
  });

  it("parses a mapped JSON array", () => {
    expect(parseMappedArray('["x","y"]', 2)).toEqual(["x", "y"]);
  });

  it("extracts object fields in schema order", () => {
    expect(parseExtract('{"name":"Ada","role":"CEO"}', "name, role")).toEqual(["Ada", "CEO"]);
  });

  it("builds a map prompt that includes the rows", () => {
    const call = buildMapCall("uppercase", [["a"], ["b"]]);
    expect(call.user).toContain("uppercase");
    expect(call.user).toContain('["a"]');
  });

  it("cache keys differ when values differ", () => {
    expect(cacheKey(["m", "p", [[1]]])).not.toBe(cacheKey(["m", "p", [[2]]]));
  });

  it("flattens a range row-major and treats null as empty", () => {
    expect(flattenCells([["a", null], [1, "b"]])).toEqual(["a", "", "1", "b"]);
  });

  it("reshapes a flat list back to a grid", () => {
    expect(reshapeCells(["a", "", "1", "b"], 2, 2)).toEqual([
      ["a", ""],
      ["1", "b"]
    ]);
  });

  it("builds a translate prompt with auto-detect and cell JSON", () => {
    const call = buildTranslateCall("French", ["Hello", ""], undefined);
    expect(call.system).toMatch(/JSON array/i);
    expect(call.user).toContain("Target language: French");
    expect(call.user).toMatch(/auto-detect/i);
    expect(call.user).toContain('["Hello",""]');
  });

  it("includes an explicit source language when provided", () => {
    const call = buildTranslateCall("ja", ["Hello"], "English");
    expect(call.user).toContain("Source language: English");
    expect(call.user).not.toMatch(/auto-detect/i);
  });
});
