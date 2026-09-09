import { Changeset, type HostAdapter, type RawFacts } from "@openplugin/core";

export class PowerPointHost implements HostAdapter {
  readonly kind = "powerpoint" as const;

  async getRawFacts(): Promise<RawFacts> {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      const outline: Array<{ index: number; title: string; shapeCount: number }> = [];
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i];
        const shapes = slide.shapes;
        shapes.load("items/name,items/textFrame/textRange/text");
        outline.push({ index: i, title: `Slide ${i + 1}`, shapeCount: 0 });
      }
      await context.sync();

      slides.items.forEach((slide, i) => {
        const shapes = slide.shapes.items;
        outline[i].shapeCount = shapes.length;
        const titleShape = shapes.find((s) => /title/i.test(s.name)) ?? shapes[0];
        try {
          outline[i].title = titleShape?.textFrame.textRange.text || outline[i].title;
        } catch {
          /* shape has no text frame */
        }
      });

      const selected = await getSelectedText();
      return {
        host: "powerpoint",
        title: "Presentation",
        slides: outline,
        selection: {
          slideIndex: 0,
          text: selected.slice(0, 1000)
        }
      };
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await PowerPoint.run(async (context) => {
      for (const change of changeset.forHost("powerpoint")) {
        if (change.op === "addSlide") {
          const slide = context.presentation.slides.add();
          const titleBox = slide.shapes.addTextBox(40, 30, 600, 50);
          titleBox.textFrame.textRange.text = change.title;
          if (change.bullets?.length) {
            const body = slide.shapes.addTextBox(40, 100, 600, 300);
            body.textFrame.textRange.text = change.bullets.map((b) => `• ${b}`).join("\n");
          }
        } else if (change.op === "setShapeText") {
          const slide = context.presentation.slides.getItemAt(change.slideIndex);
          const shapes = slide.shapes;
          shapes.load("items/name,items/textFrame/textRange/text");
          await context.sync();
          const shape = change.shapeName
            ? shapes.items.find((s) => s.name === change.shapeName)
            : shapes.items[0];
          if (shape) shape.textFrame.textRange.text = change.text;
        } else if (change.op === "deleteSlide") {
          context.presentation.slides.getItemAt(change.slideIndex).delete();
        } else if (change.op === "setNotes") {
          const slide = context.presentation.slides.getItemAt(change.slideIndex);
          try {
            slide.notesPage.body.textFrame.textRange.text = change.notes;
          } catch {
            /* notes API varies by requirement set */
          }
        }
      }
      await context.sync();
    });
  }
}

async function getSelectedText(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        resolve(typeof result.value === "string" ? result.value : "");
      });
    } catch {
      resolve("");
    }
  });
}
