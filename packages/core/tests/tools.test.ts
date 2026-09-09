import { describe, expect, it } from "vitest";
import { Changeset } from "../src/tools/changeset.js";
import { executeHostTool, listToolDefinitions } from "../src/tools/registry.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { FakeWordHost } from "../src/hosts/fake-word.js";
import { FakePowerPointHost } from "../src/hosts/fake-powerpoint.js";

describe("listToolDefinitions", () => {
  it("returns only tools for the current host plus meta tools", () => {
    const names = listToolDefinitions("excel")
      .filter((t): t is Extract<typeof t, { type: "function" }> => t.type === "function")
      .map((t) => t.function.name);
    expect(names).toContain("excel.readRange");
    expect(names).toContain("skills.load");
    expect(names).not.toContain("word.insertParagraphs");
    expect(names).not.toContain("host.executeOfficeJs");
  });

  it("includes web.search only when function search is enabled", () => {
    const names = (opts?: Parameters<typeof listToolDefinitions>[1]) =>
      listToolDefinitions("excel", opts)
        .filter((t): t is Extract<typeof t, { type: "function" }> => t.type === "function")
        .map((t) => t.function.name);
    expect(names()).not.toContain("web.search");
    expect(names({ webSearch: "function" })).toContain("web.search");
  });

  it("includes executeOfficeJs only when enabled", () => {
    const names = listToolDefinitions("word", { executeJsEnabled: true })
      .filter((t): t is Extract<typeof t, { type: "function" }> => t.type === "function")
      .map((t) => t.function.name);
    expect(names).toContain("host.executeOfficeJs");
  });
});

describe("executeHostTool", () => {
  it("reads a truncated excel range", async () => {
    const host = new FakeExcelHost();
    host.sheets.Sheet1.values = Array.from({ length: 250 }, (_, r) => [r, "x"]);
    const cs = new Changeset();
    const result = (await executeHostTool(host, "excel.readRange", { sheet: "Sheet1", address: "A1:B250" }, cs)) as {
      truncated: boolean;
      values: unknown[][];
    };
    expect(result.truncated).toBe(true);
    expect(result.values.length).toBeLessThanOrEqual(200);
    expect(cs.isEmpty()).toBe(true);
  });

  it("queues writes instead of mutating the host", async () => {
    const host = new FakeExcelHost();
    const cs = new Changeset();
    await executeHostTool(
      host,
      "excel.writeRange",
      { sheet: "Sheet1", address: "A1", values: [["z"]] },
      cs
    );
    expect(host.sheets.Sheet1.values[0][0]).toBe("");
    expect(cs.changes[0]).toMatchObject({ op: "writeRange", address: "A1" });
  });

  it("rewrites writeRange address to the values shape before snapshotting", async () => {
    const host = new FakeExcelHost();
    host.sheets.Sheet1.values = [
      ["old", "keep"],
      ["x", "y"]
    ];
    const cs = new Changeset();
    await executeHostTool(
      host,
      "excel.writeRange",
      { sheet: "Sheet1", address: "A1", values: [["h1", "h2"], [1, 2]] },
      cs
    );
    expect(cs.changes[0]).toMatchObject({
      op: "writeRange",
      address: "A1:B2",
      values: [["h1", "h2"], [1, 2]]
    });
    expect((cs.changes[0] as { before?: unknown[][] }).before).toEqual([
      ["old", "keep"],
      ["x", "y"]
    ]);
  });

  it("applies an excel changeset in one shot", async () => {
    const host = new FakeExcelHost();
    const cs = new Changeset();
    await executeHostTool(
      host,
      "excel.writeRange",
      { sheet: "Sheet1", address: "A1:B1", values: [["h1", "h2"]] },
      cs
    );
    await host.apply(cs);
    expect(host.sheets.Sheet1.values[0]).toEqual(["h1", "h2"]);
  });

  it("queues a word insert and applies it", async () => {
    const host = new FakeWordHost();
    const cs = new Changeset();
    await executeHostTool(
      host,
      "word.insertParagraphs",
      { paragraphs: ["Hello memo"], location: "end" },
      cs
    );
    await host.apply(cs);
    expect(host.paragraphs.map((p) => p.text)).toContain("Hello memo");
  });

  it("queues a powerpoint slide add and applies it", async () => {
    const host = new FakePowerPointHost();
    const cs = new Changeset();
    await executeHostTool(
      host,
      "ppt.addSlide",
      { title: "Q3 plan", bullets: ["Hire", "Ship"] },
      cs
    );
    await host.apply(cs);
    expect(host.slides.at(-1)?.title).toBe("Q3 plan");
  });

  it("reads a range as csv and searches cells", async () => {
    const host = new FakeExcelHost();
    host.sheets.Sheet1.values = [
      ["Name", "Amt"],
      ["Gadgets", 12]
    ];
    const cs = new Changeset();
    const csv = (await executeHostTool(host, "excel.readCsv", { sheet: "Sheet1", address: "A1:B2" }, cs)) as {
      csv: string;
    };
    expect(csv.csv).toContain("Gadgets,12");
    const hits = (await executeHostTool(host, "excel.search", { query: "Gadget" }, cs)) as Array<{
      address: string;
    }>;
    expect(hits[0]?.address).toBe("A2");
  });

  it("queues a word paragraph replace and a ppt duplicate", async () => {
    const word = new FakeWordHost();
    word.paragraphs = [{ text: "old", style: "Normal" }];
    const wcs = new Changeset();
    await executeHostTool(word, "word.replaceParagraph", { index: 0, text: "new" }, wcs);
    await word.apply(wcs);
    expect(word.paragraphs[0]?.text).toBe("new");

    const ppt = new FakePowerPointHost();
    const pcs = new Changeset();
    await executeHostTool(ppt, "ppt.duplicateSlide", { slideIndex: 0 }, pcs);
    await ppt.apply(pcs);
    expect(ppt.slides).toHaveLength(2);
  });

  it("rejects unknown tools instead of returning an empty read result", async () => {
    const host = new FakeExcelHost();
    const cs = new Changeset();

    await expect(executeHostTool(host, "excel.missingTool", {}, cs)).rejects.toThrow(
      "Unknown host tool: excel.missingTool"
    );
    expect(cs.isEmpty()).toBe(true);
  });
});
