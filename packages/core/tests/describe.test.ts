import { describe, expect, it } from "vitest";
import { describeActivity, describeTool } from "../src/tools/describe.js";

describe("describeTool", () => {
  it("uses a human phrase instead of the raw tool id", () => {
    expect(describeTool("excel.readRange", { sheet: "Sheet1", address: "A1:B2" })).toBe(
      "Reading Sheet1!A1:B2"
    );
  });
});

describe("describeActivity", () => {
  it("classifies reads, writes, skills, and search", () => {
    expect(describeActivity("excel.readRange", { sheet: "Sheet1", address: "A1" })).toMatchObject({
      activity: "read",
      label: "Reading Sheet1!A1"
    });
    expect(describeActivity("word.replaceSelection", { text: "hi" })).toMatchObject({
      activity: "write",
      label: "Drafting a replacement"
    });
    expect(describeActivity("skills.load", { name: "selection-rewrite" })).toMatchObject({
      activity: "skill",
      label: "Reading skill selection-rewrite"
    });
    expect(describeActivity("web.search", { query: "gdp germany" })).toMatchObject({
      activity: "search",
      label: "Searching the web: gdp germany"
    });
    expect(describeActivity("web.fetch", { url: "https://example.com/a" })).toMatchObject({
      activity: "fetch",
      label: "Opening example.com/a"
    });
  });
});
