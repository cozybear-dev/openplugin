import { describe, expect, it } from "vitest";
import { Changeset } from "../src/tools/changeset.js";

describe("Changeset", () => {
  it("accumulates mutations without applying them", () => {
    const cs = new Changeset();
    cs.add({
      host: "excel",
      op: "writeRange",
      sheet: "Sheet1",
      address: "A1:B2",
      values: [
        ["h1", "h2"],
        [1, 2]
      ]
    });
    expect(cs.changes).toHaveLength(1);
    expect(cs.isEmpty()).toBe(false);
  });

  it("builds inverses in reverse order so sequential edits undo correctly", () => {
    const cs = new Changeset();
    cs.add({ host: "word", op: "replaceSelection", text: "middle", beforeText: "first" });
    cs.add({ host: "word", op: "replaceSelection", text: "last", beforeText: "middle" });

    expect(cs.inverse().changes.map((change) => change.text)).toEqual(["middle", "first"]);
  });

  it("renders a preview that names the target ranges", () => {
    const cs = new Changeset();
    cs.add({
      host: "excel",
      op: "writeRange",
      sheet: "Sheet1",
      address: "A1",
      values: [["hello"]]
    });
    cs.add({
      host: "word",
      op: "insertParagraphs",
      paragraphs: ["A memo."],
      location: "end"
    });
    const preview = cs.preview();
    expect(preview).toContain("Sheet1!A1");
    expect(preview).toContain("Insert paragraphs");
    expect(preview).toContain("A memo.");
  });

  it("clears after discard", () => {
    const cs = new Changeset();
    cs.add({
      host: "powerpoint",
      op: "addSlide",
      title: "Q3",
      bullets: ["One"]
    });
    cs.clear();
    expect(cs.isEmpty()).toBe(true);
    expect(cs.preview()).toBe("No pending changes.");
  });

  it("filters selected hunks by the same ids the diff panel uses", () => {
    const cs = new Changeset();
    cs.add({ host: "word", op: "replaceSelection", text: "one" });
    cs.add({
      host: "word",
      op: "insertParagraphs",
      paragraphs: ["two"],
      location: "end"
    });
    cs.add({ host: "word", op: "searchReplace", search: "a", replace: "b", all: true });
    cs.add({ host: "word", op: "applyStyle", style: "Heading 1", target: "selection" });
    cs.add({ host: "word", op: "insertComment", text: "note" });
    const hunks = cs.diff();
    expect(hunks).toHaveLength(5);
    const selected = new Set(hunks.map((h) => h.id));
    const filtered = cs.filter(selected);
    expect(filtered.changes).toHaveLength(5);
    expect(filtered.changes.map((c) => c.op)).toEqual([
      "replaceSelection",
      "insertParagraphs",
      "searchReplace",
      "applyStyle",
      "insertComment"
    ]);
  });
});
