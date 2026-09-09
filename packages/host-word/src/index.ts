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

  async readSelectionText(): Promise<string> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text ?? "";
    });
  }

  async readParagraphs(args?: { start?: number; count?: number }) {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load(["text", "style"]);
      await context.sync();
      const start = args?.start ?? 0;
      const count = args?.count ?? paragraphs.items.length;
      return paragraphs.items.slice(start, start + count).map((p, i) => ({
        index: start + i,
        text: p.text,
        style: p.style
      }));
    });
  }

  async findText(args: { query: string; max?: number }) {
    return Word.run(async (context) => {
      const results = context.document.body.search(args.query);
      results.load("items/text");
      await context.sync();
      return results.items.slice(0, args.max ?? 40).map((item, paragraphIndex) => ({
        paragraphIndex,
        text: item.text
      }));
    });
  }

  async listComments() {
    return Word.run(async (context) => {
      try {
        const comments = context.document.body.getComments();
        comments.load("items/content,items/resolved");
        await context.sync();
        return comments.items.map((c, index) => ({
          index,
          text: c.content,
          resolved: Boolean(c.resolved)
        }));
      } catch {
        return [];
      }
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await Word.run(async (context) => {
      try {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      } catch {
        /* requirement set missing */
      }
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
          if (change.location === "afterSelection") {
            for (const paragraph of [...change.paragraphs].reverse()) {
              selection.insertParagraph(paragraph, Word.InsertLocation.after);
            }
          } else {
            const paragraphs = change.location === "start" ? [...change.paragraphs].reverse() : change.paragraphs;
            for (const paragraph of paragraphs) {
              body.insertParagraph(paragraph, loc);
            }
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
        } else if (change.op === "replaceParagraph") {
          body.paragraphs.load("items");
          await context.sync();
          const p = body.paragraphs.items[change.index];
          if (p) p.insertText(change.text, Word.InsertLocation.replace);
        } else if (change.op === "replyComment") {
          try {
            const comments = body.getComments();
            comments.load("items");
            await context.sync();
            comments.items[change.commentIndex]?.reply(change.text);
          } catch {
            /* comments API */
          }
        } else if (change.op === "resolveComment") {
          try {
            const comments = body.getComments();
            comments.load("items");
            await context.sync();
            const c = comments.items[change.commentIndex];
            if (c) c.resolved = true;
          } catch {
            /* comments API */
          }
        }
      }
      await context.sync();
    });
  }
}
