import { describe, expect, it } from "vitest";
import { applyChangeset } from "../src/tools/apply.js";
import { Changeset } from "../src/tools/changeset.js";
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
});
