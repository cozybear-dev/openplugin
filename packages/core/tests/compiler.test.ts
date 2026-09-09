import { describe, expect, it } from "vitest";
import { compileSnapshot, estimateTokens } from "../src/context/compiler.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { FakePowerPointHost } from "../src/hosts/fake-powerpoint.js";
import { FakeWordHost } from "../src/hosts/fake-word.js";

describe("estimateTokens", () => {
  it("uses ~4 characters per token", () => {
    expect(estimateTokens("abcd")).toBe(1);
    expect(estimateTokens("a".repeat(400))).toBe(100);
  });
});

describe("compileSnapshot", () => {
  it("keeps the Excel selection even when the sheet is huge", async () => {
    const host = new FakeExcelHost();
    host.sheets.Sheet1.values = Array.from({ length: 10000 }, (_, r) => [
      `row${r}`,
      String(r),
      "x".repeat(80)
    ]);
    host.selection = { sheet: "Sheet1", address: "A1:B2", values: [["h1", "h2"], ["a", "b"]] };

    const snap = await compileSnapshot(host, { tokenBudget: 4000 });
    expect(snap.host).toBe("excel");
    expect(JSON.stringify(snap.selection)).toContain("h1");
    expect(snap.tokenEstimate).toBeLessThanOrEqual(4000);
    expect(snap.outline).toEqual(
      expect.arrayContaining([expect.objectContaining({ name: "Sheet1", rows: 10000 })])
    );
  });

  it("captures a Word outline and selection window, not the full body", async () => {
    const host = new FakeWordHost();
    host.paragraphs = [
      { text: "Title", style: "Heading 1" },
      { text: "Intro " + "word ".repeat(2000), style: "Normal" },
      { text: "Section", style: "Heading 1" }
    ];
    host.selection = { text: "Intro word", style: "Normal", paragraphIndex: 1 };

    const snap = await compileSnapshot(host, { tokenBudget: 1500 });
    expect(snap.host).toBe("word");
    expect(JSON.stringify(snap.outline)).toContain("Title");
    expect(snap.tokenEstimate).toBeLessThanOrEqual(1500);
    expect(JSON.stringify(snap.selection)).toContain("Intro");
  });

  it("captures PowerPoint slide titles without dumping every shape", async () => {
    const host = new FakePowerPointHost();
    host.slides = [
      { title: "Agenda", shapes: [{ name: "t", text: "Agenda" }, { name: "body", text: "x".repeat(5000) }], notes: "" },
      { title: "Q3", shapes: [{ name: "t", text: "Q3" }], notes: "Speak slowly" }
    ];
    host.selection = { slideIndex: 0, shapeName: "t", text: "Agenda" };

    const snap = await compileSnapshot(host, { tokenBudget: 800 });
    expect(snap.host).toBe("powerpoint");
    expect(JSON.stringify(snap.outline)).toContain("Agenda");
    expect(JSON.stringify(snap.outline)).toContain("Q3");
    expect(snap.tokenEstimate).toBeLessThanOrEqual(800);
  });

  it("includes a compact diff against the previous snapshot", async () => {
    const host = new FakeExcelHost();
    host.selection = { sheet: "Sheet1", address: "A1", values: [["old"]] };
    const first = await compileSnapshot(host, { tokenBudget: 2000 });
    host.selection = { sheet: "Sheet1", address: "A1", values: [["new"]] };
    const second = await compileSnapshot(host, { tokenBudget: 2000, previous: first });
    expect(second.diffFromPrevious).toContain("new");
  });
});
