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

  async readSelectionText(): Promise<string> {
    return this.selection.text;
  }

  comments: Array<{ text: string; author?: string; resolved?: boolean }> = [];

  async readParagraphs(args?: { start?: number; count?: number }) {
    const start = args?.start ?? 0;
    const count = args?.count ?? this.paragraphs.length;
    return this.paragraphs.slice(start, start + count).map((p, i) => ({
      index: start + i,
      text: p.text,
      style: p.style
    }));
  }

  async findText(args: { query: string; max?: number }) {
    const q = args.query.toLowerCase();
    const hits: Array<{ paragraphIndex: number; text: string }> = [];
    this.paragraphs.forEach((p, i) => {
      if (p.text.toLowerCase().includes(q)) hits.push({ paragraphIndex: i, text: p.text });
    });
    return hits.slice(0, args.max ?? 50);
  }

  async listComments() {
    return this.comments.map((c, index) => ({ index, ...c }));
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
        const index = change.paragraphIndex ?? this.selection.paragraphIndex;
        const p = this.paragraphs[index];
        if (p) p.style = change.style;
      } else if (change.op === "insertTable") {
        const cells = change.cells?.map((row) => row.join("|")).join(" / ");
        this.paragraphs.push({
          text: cells ? `[table ${change.rows}x${change.cols} ${cells}]` : `[table ${change.rows}x${change.cols}]`,
          style: "Normal"
        });
      } else if (change.op === "insertComment") {
        const p = this.paragraphs[this.selection.paragraphIndex];
        if (p) p.text += ` /* ${change.text} */`;
        this.comments.push({ text: change.text });
      } else if (change.op === "replaceParagraph") {
        const p = this.paragraphs[change.index];
        if (p) p.text = change.text;
      } else if (change.op === "replyComment") {
        const c = this.comments[change.commentIndex];
        if (c) c.text += ` / ${change.text}`;
      } else if (change.op === "resolveComment") {
        const c = this.comments[change.commentIndex];
        if (c) c.resolved = true;
      }
    }
  }
}
