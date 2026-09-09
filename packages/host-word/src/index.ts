import { Changeset, type HostAdapter, type RawFacts } from "@openplugin/core";

export class WordHost implements HostAdapter {
  readonly kind = "word" as const;

  async getRawFacts(): Promise<RawFacts> {
    return Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load(["text", "style"]);
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const items = paragraphs.items.map((p, index) => ({
        text: p.text,
        style: p.style,
        index
      }));
      const headings = items.filter((p) => /heading/i.test(p.style));
      const selText = selection.text ?? "";
      const paragraphIndex = Math.max(
        0,
        items.findIndex((p) => p.text.includes(selText.slice(0, 40)))
      );
      const surrounding = items
        .slice(Math.max(0, paragraphIndex - 1), paragraphIndex + 2)
        .map((p) => p.text.slice(0, 400));

      return {
        host: "word",
        title: "Document",
        headings: headings.map((h) => ({ text: h.text.slice(0, 200), style: h.style, index: h.index })),
        paragraphCount: items.length,
        selection: {
          text: selText.slice(0, 2000),
          style: items[paragraphIndex]?.style ?? "Normal",
          paragraphIndex
        },
        surrounding
      };
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await Word.run(async (context) => {
      const body = context.document.body;
      const selection = context.document.getSelection();
      for (const change of changeset.forHost("word")) {
        if (change.op === "replaceSelection") {
          selection.insertText(change.text, Word.InsertLocation.replace);
        } else if (change.op === "insertParagraphs") {
          const loc =
            change.location === "start"
              ? Word.InsertLocation.start
              : change.location === "afterSelection"
                ? Word.InsertLocation.after
                : Word.InsertLocation.end;
          const target = change.location === "afterSelection" ? selection : body;
          for (const paragraph of change.paragraphs) {
            target.insertParagraph(paragraph, loc);
          }
        } else if (change.op === "searchReplace") {
          const results = body.search(change.search);
          results.load("items");
          await context.sync();
          const items = change.all ? results.items : results.items.slice(0, 1);
          for (const item of items) item.insertText(change.replace, Word.InsertLocation.replace);
        } else if (change.op === "applyStyle") {
          selection.paragraphs.load("items");
          await context.sync();
          for (const p of selection.paragraphs.items) p.style = change.style;
        } else if (change.op === "insertTable") {
          body.insertTable(change.rows, change.cols, Word.InsertLocation.end);
        } else if (change.op === "insertComment") {
          selection.insertComment(change.text);
        }
      }
      await context.sync();
    });
  }
}
