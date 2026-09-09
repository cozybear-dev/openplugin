import { describe, expect, it } from "vitest";
import { Changeset } from "../src/tools/changeset.js";
import { executeHostTool, listToolDefinitions } from "../src/tools/registry.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { FakeWordHost } from "../src/hosts/fake-word.js";
import { FakePowerPointHost } from "../src/hosts/fake-powerpoint.js";

describe("listToolDefinitions", () => {
  it("returns only tools for the current host plus meta tools", () => {
    const names = listToolDefinitions("excel").map((t) => t.function.name);
    expect(names).toContain("excel.readRange");
    expect(names).toContain("skills.load");
    expect(names).not.toContain("word.insertParagraphs");
    expect(names).not.toContain("host.executeOfficeJs");
  });

  it("includes executeOfficeJs only when enabled", () => {
    const names = listToolDefinitions("word", { executeJsEnabled: true }).map((t) => t.function.name);
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
});
