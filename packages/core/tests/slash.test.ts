import { describe, expect, it } from "vitest";
import { parseSlash, slashSuggestions } from "../src/skills/slash.js";

const known = ["selection-rewrite", "excel-range-cleanup", "financial-memo"];

describe("parseSlash", () => {
  it("returns none when the prompt is not a slash command", () => {
    expect(parseSlash("rewrite this", known)).toEqual({ kind: "none", rest: "rewrite this" });
  });

  it("lists skills for a bare slash", () => {
    expect(parseSlash("/", known)).toEqual({ kind: "list" });
    expect(parseSlash("/   ", known)).toEqual({ kind: "list" });
  });

  it("parses a known skill and the remainder", () => {
    expect(parseSlash("/selection-rewrite make it shorter", known)).toEqual({
      kind: "skill",
      name: "selection-rewrite",
      rest: "make it shorter"
    });
  });

  it("flags unknown skills", () => {
    expect(parseSlash("/nope", known)).toEqual({ kind: "unknown", name: "nope" });
  });

  it("filters suggestions by prefix", () => {
    expect(slashSuggestions("/sel", known)).toEqual(["selection-rewrite"]);
  });
});
