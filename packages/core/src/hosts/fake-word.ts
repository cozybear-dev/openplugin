import type { Changeset } from "../tools/changeset.js";
import type { HostAdapter, RawFacts, WordParagraph, WordSelection } from "./types.js";

export class FakeWordHost implements HostAdapter {
  readonly kind = "word" as const;
  title = "Document";
  paragraphs: WordParagraph[] = [{ text: "", style: "Normal" }];
  selection: WordSelection = { text: "", style: "Normal", paragraphIndex: 0 };

  async getRawFacts(): Promise<RawFacts> {
    const headings = this.paragraphs
      .map((p, index) => ({ ...p, index }))
      .filter((p) => /heading/i.test(p.style));
    const i = this.selection.paragraphIndex;
    const surrounding = this.paragraphs
      .slice(Math.max(0, i - 1), i + 2)
      .map((p) => p.text.slice(0, 400));
    return {
      host: "word",
      title: this.title,
      headings,
      paragraphCount: this.paragraphs.length,
      selection: this.selection,
      surrounding
    };
  }

  async apply(changeset: Changeset): Promise<void> {
    for (const change of changeset.forHost("word")) {
      if (change.op === "replaceSelection") {
        const p = this.paragraphs[this.selection.paragraphIndex];
        if (p) p.text = change.text;
        this.selection.text = change.text;
      } else if (change.op === "insertParagraphs") {
        const items = change.paragraphs.map((text) => ({ text, style: "Normal" }));
        if (change.location === "start") this.paragraphs.unshift(...items);
        else if (change.location === "afterSelection") {
          this.paragraphs.splice(this.selection.paragraphIndex + 1, 0, ...items);
        } else this.paragraphs.push(...items);
      } else if (change.op === "searchReplace") {
        for (const p of this.paragraphs) {
          p.text = change.all ? p.text.split(change.search).join(change.replace) : p.text.replace(change.search, change.replace);
        }
      } else if (change.op === "applyStyle") {
        const p = this.paragraphs[this.selection.paragraphIndex];
        if (p) p.style = change.style;
      } else if (change.op === "insertTable") {
        this.paragraphs.push({
          text: `[table ${change.rows}x${change.cols}]`,
          style: "Normal"
        });
      } else if (change.op === "insertComment") {
        const p = this.paragraphs[this.selection.paragraphIndex];
        if (p) p.text += ` /* ${change.text} */`;
      }
    }
  }
}
