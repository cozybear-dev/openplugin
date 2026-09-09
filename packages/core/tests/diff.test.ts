import { describe, expect, it } from "vitest";
import { Changeset } from "../src/tools/changeset.js";
import { cellDiff, invertChange, textDiff } from "../src/tools/diff.js";
import { executeHostTool } from "../src/tools/registry.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { FakeWordHost } from "../src/hosts/fake-word.js";

describe("cellDiff", () => {
  it("pairs before and after by address", () => {
    const cells = cellDiff("Sheet1", "A1:B1", [["old", "keep"]], [["new", "keep"]]);
    expect(cells).toEqual([
      { address: "Sheet1!A1", before: "old", after: "new", changed: true },
      { address: "Sheet1!B1", before: "keep", after: "keep", changed: false }
    ]);
  });
});

describe("textDiff", () => {
  it("marks deleted and inserted words", () => {
    const parts = textDiff("hello gray world", "hello green world");
    expect(parts).toContainEqual({ type: "del", text: "gray" });
    expect(parts).toContainEqual({ type: "ins", text: "green" });
    expect(
      parts.filter((p) => p.type === "eq" && p.text.trim()).map((p) => p.text)
    ).toEqual(["hello", "world"]);
  });
});

describe("Changeset.diff and invert", () => {
  it("shows before and after for an excel write and inverts it", async () => {
    const host = new FakeExcelHost();
    host.sheets.Sheet1.values = [["old", "x"]];
    const cs = new Changeset();
    await executeHostTool(
      host,
      "excel.writeRange",
      { sheet: "Sheet1", address: "A1", values: [["new"]] },
      cs
    );
    const hunks = cs.diff();
    expect(hunks[0]?.kind).toBe("grid");
    if (hunks[0]?.kind !== "grid") throw new Error("expected grid");
    expect(hunks[0].cells[0]).toMatchObject({ before: "old", after: "new", changed: true });

    const inverse = cs.inverse();
    expect(inverse.changes[0]).toMatchObject({
      op: "writeRange",
      values: [["old"]]
    });
    await host.apply(cs);
    expect(host.sheets.Sheet1.values[0][0]).toBe("new");
    await host.apply(inverse);
    expect(host.sheets.Sheet1.values[0][0]).toBe("old");
  });

  it("diffs a word replacement", async () => {
    const host = new FakeWordHost();
    host.selection = { text: "draft copy", style: "Normal", paragraphIndex: 0 };
    host.paragraphs = [{ text: "draft copy", style: "Normal" }];
    const cs = new Changeset();
    await executeHostTool(host, "word.replaceSelection", { text: "final copy" }, cs);
    const hunks = cs.diff();
    expect(hunks[0]?.kind).toBe("text");
    if (hunks[0]?.kind !== "text") throw new Error("expected text");
    expect(hunks[0].parts.some((p) => p.type === "del" && p.text === "draft")).toBe(true);
    expect(hunks[0].parts.some((p) => p.type === "ins" && p.text === "final")).toBe(true);
  });
});

describe("invertChange", () => {
  it("returns null for irreversible ops", () => {
    expect(
      invertChange({ host: "excel", op: "createChart", sheet: "S", source: "A1", chartType: "bar" })
    ).toBeNull();
  });
});
