import { describe, expect, it } from "vitest";
import { splitCitations } from "../src/citations";

describe("splitCitations", () => {
  it("returns the original string when there are no citations", () => {
    expect(splitCitations("hello **world**")).toEqual([{ text: "hello **world**" }]);
  });

  it("wraps an Excel cell address", () => {
    const parts = splitCitations("See Sheet1!B2 for the total.");
    expect(parts).toEqual([
      { text: "See " },
      { text: "Sheet1!B2", citation: expect.objectContaining({ kind: "cell", sheet: "Sheet1", address: "B2" }) },
      { text: " for the total." }
    ]);
  });

  it("wraps a slide citation", () => {
    const parts = splitCitations("On slide 3 we tighten bullets.");
    expect(parts[1]).toEqual(
      expect.objectContaining({ text: "slide 3", citation: expect.objectContaining({ kind: "slide", slide: 3 }) })
    );
  });
});
