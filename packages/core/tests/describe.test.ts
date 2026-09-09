import { describe, expect, it } from "vitest";
import { describeTool } from "../src/tools/describe.js";

describe("describeTool", () => {
  it("uses a human phrase instead of the raw tool id", () => {
    expect(describeTool("excel.readRange", { sheet: "Sheet1", address: "A1:B2" })).toBe(
      "Reading Sheet1!A1:B2"
    );
  });
});
