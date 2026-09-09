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
});
