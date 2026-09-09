import { Changeset } from "@openplugin/core";
import { WordHost } from "@openplugin/host-word";
import { afterEach, describe, expect, it } from "vitest";

function installWord(ctx: unknown) {
  (globalThis as { Word?: unknown }).Word = {
    InsertLocation: { replace: "Replace", start: "Start", end: "End", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    run: async (fn: (c: unknown) => Promise<unknown>) => fn(ctx)
  };
}

describe("WordHost", () => {
  afterEach(() => {
    delete (globalThis as { Word?: unknown }).Word;
  });

  it("applyStyle with paragraphIndex sets that paragraph's style", async () => {
    const paras = [{ style: "Normal" }, { style: "Normal" }];
    installWord({
      document: {
        changeTrackingMode: undefined,
        body: { paragraphs: { load() {}, items: paras } },
        getSelection: () => ({ paragraphs: { load() {}, items: [paras[0]] } })
      },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({
      host: "word",
      op: "applyStyle",
      style: "Heading 1",
      target: "selection",
      paragraphIndex: 1
    });
    await new WordHost().apply(cs);
    expect(paras[1].style).toBe("Heading 1");
    expect(paras[0].style).toBe("Normal");
  });

  it("searchReplace applies from last match to first", async () => {
    const calls: string[] = [];
    const makeRange = (id: string) => ({
      insertText: (value: string, loc: string) => {
        calls.push(`${id}:${loc}:${value}`);
      }
    });
    installWord({
      document: {
        changeTrackingMode: undefined,
        body: {
          search: () => ({ load() {}, items: [makeRange("r0"), makeRange("r1")] }),
          paragraphs: { load() {}, items: [] }
        },
        getSelection: () => ({})
      },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({ host: "word", op: "searchReplace", search: "Q3", replace: "Q4", all: true });
    await new WordHost().apply(cs);
    expect(calls).toEqual(["r1:Replace:Q4", "r0:Replace:Q4"]);
  });

  it("findText maps hits to the paragraph that contains them", async () => {
    const hits = [{ text: "Weekly" }];
    installWord({
      document: {
        body: {
          search: () => ({ load() {}, items: hits }),
          paragraphs: {
            load() {},
            items: [{ text: "OpenPlugin live fixture" }, { text: "Q3 sales" }, { text: "Weekly Update" }]
          }
        },
        getSelection: () => ({ load() {}, text: "" })
      },
      sync: async () => undefined
    });
    expect(await new WordHost().findText({ query: "Weekly" })).toEqual([
      { paragraphIndex: 2, text: "Weekly" }
    ]);
  });

  it("insertTable writes cell values when provided", async () => {
    const written: Array<{ r: number; c: number; text: string }> = [];
    const table = {
      getCell: (r: number, c: number) => ({
        body: {
          insertText: (text: string) => {
            written.push({ r, c, text });
          }
        }
      })
    };
    installWord({
      document: {
        changeTrackingMode: undefined,
        body: {
          insertTable: () => table,
          paragraphs: { load() {}, items: [] }
        },
        getSelection: () => ({})
      },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({
      host: "word",
      op: "insertTable",
      rows: 2,
      cols: 2,
      cells: [
        ["A", "B"],
        ["1", "2"]
      ]
    });
    await new WordHost().apply(cs);
    expect(written).toEqual([
      { r: 0, c: 0, text: "A" },
      { r: 0, c: 1, text: "B" },
      { r: 1, c: 0, text: "1" },
      { r: 1, c: 1, text: "2" }
    ]);
  });
});
