import { describe, expect, it } from "vitest";
import { applyChangeset, applyFailureMessage } from "../src/tools/apply.js";
import { Changeset } from "../src/tools/changeset.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { FakeWordHost } from "../src/hosts/fake-word.js";

class FlakyWordHost extends FakeWordHost {
  async apply(changeset: Changeset): Promise<void> {
    for (const change of changeset.forHost("word")) {
      if (change.op === "insertComment") throw new Error("no selection");
    }
    await super.apply(changeset);
  }
}

describe("applyChangeset", () => {
  it("applies later word ops after one op throws", async () => {
    const host = new FlakyWordHost();
    host.paragraphs = [{ text: "hello", style: "Normal" }];
    host.selection = { text: "hello", style: "Normal", paragraphIndex: 0 };
    const cs = new Changeset();
    cs.add({ host: "word", op: "replaceSelection", text: "hi" });
    cs.add({ host: "word", op: "insertComment", text: "note" });
    cs.add({
      host: "word",
      op: "insertParagraphs",
      paragraphs: ["tail"],
      location: "end"
    });

    const results = await applyChangeset(host, cs);
    expect(results.map((r) => r.ok)).toEqual([true, false, true]);
    expect(results[1]).toMatchObject({ ok: false, error: "no selection" });
    expect(host.selection.text).toBe("hi");
    expect(host.paragraphs.map((p) => p.text)).toContain("tail");
  });

  it("rebases a later excel format after an insert before applying", async () => {
    const host = new FakeExcelHost();
    const cs = new Changeset();
    cs.add({
      host: "excel",
      op: "modifySheet",
      sheet: "Sheet1",
      operation: "insert",
      dimension: "rows",
      reference: "1",
      count: 1
    });
    cs.add({
      host: "excel",
      op: "formatRange",
      sheet: "Sheet1",
      address: "A1",
      numberFormat: "0.00"
    });
    const results = await applyChangeset(host, cs);
    expect(results.every((r) => r.ok)).toBe(true);
    expect(host.formats).toEqual([{ sheet: "Sheet1", address: "A2", numberFormat: "0.00", bold: undefined }]);
  });

  it("applies a sheet-qualified format address without throwing", async () => {
    const host = new FakeExcelHost();
    const cs = new Changeset();
    cs.add({
      host: "excel",
      op: "formatRange",
      sheet: "Sheet1",
      address: "Sheet1!B2:B4",
      numberFormat: "0.00"
    });
    const results = await applyChangeset(host, cs);
    expect(results).toEqual([{ change: expect.objectContaining({ address: "B2:B4" }), ok: true }]);
  });
});

describe("applyFailureMessage", () => {
  it("warns when only formatting failed after other ops succeeded", () => {
    const write = {
      change: { host: "excel" as const, op: "writeRange" as const, sheet: "S", address: "A1", values: [["1"]] },
      ok: true as const
    };
    const format = {
      change: { host: "excel" as const, op: "formatRange" as const, sheet: "S", address: "A1", numberFormat: "0" },
      ok: false as const,
      error: "InvalidArgument"
    };
    const msg = applyFailureMessage([write], [format]);
    expect(msg.kind).toBe("warning");
    expect(msg.text).toMatch(/format/i);
    expect(msg.text).not.toMatch(/^Could not apply/);
  });

  it("errors when a non-cosmetic op failed", () => {
    const write = {
      change: { host: "excel" as const, op: "writeRange" as const, sheet: "S", address: "A1", values: [["1"]] },
      ok: false as const,
      error: "blocked"
    };
    const msg = applyFailureMessage([], [write]);
    expect(msg.kind).toBe("error");
    expect(msg.text).toMatch(/Could not apply 1 change/);
  });
});
