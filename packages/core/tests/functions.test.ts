import { describe, expect, it } from "vitest";
import {
  buildMapCall,
  cacheKey,
  chunkRows,
  parseExtract,
  parseMappedArray,
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
});
