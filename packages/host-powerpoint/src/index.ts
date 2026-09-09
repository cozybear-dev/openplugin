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

  async readSelectionText(): Promise<string> {
    return getSelectedText();
  }

  async readShapeText(args: { slideIndex: number; shapeName?: string }): Promise<string> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(args.slideIndex);
      const shapes = slide.shapes;
      shapes.load("items/name,items/textFrame/textRange/text");
      await context.sync();
      const shape = args.shapeName
        ? shapes.items.find((s) => s.name === args.shapeName)
        : shapes.items[0];
      try {
        return shape?.textFrame.textRange.text ?? "";
      } catch {
        return "";
      }
    });
  }

  async readNotes(slideIndex: number): Promise<string> {
    // PowerPoint's JavaScript API types do not expose slide notes.
    return "";
  }

  async readSlide(slideIndex: number) {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex);
      const shapes = slide.shapes;
      shapes.load("items/name,items/textFrame/textRange/text");
      await context.sync();
      const out = shapes.items.map((s) => {
        try {
          return { name: s.name, text: s.textFrame.textRange.text ?? "" };
        } catch {
          return { name: s.name, text: "" };
        }
      });
      return { title: out[0]?.text ?? `Slide ${slideIndex + 1}`, shapes: out, notes: "" };
    });
  }

  async listLayouts() {
    return PowerPoint.run(async (context) => {
      try {
        const layouts = context.presentation.slideMasters.getItemAt(0).layouts;
        layouts.load("items/name");
        await context.sync();
        return layouts.items.map((l) => l.name);
      } catch {
        return [];
      }
    });
  }

  async apply(changeset: Changeset): Promise<void> {
    await PowerPoint.run(async (context) => {
      for (const change of changeset.forHost("powerpoint")) {
        if (change.op === "addSlide") {
          const slides = context.presentation.slides;
          slides.add();
          slides.load("items");
          await context.sync();
          const slide = slides.items[slides.items.length - 1];
          const titleBox = slide.shapes.addTextBox(change.title, { left: 40, top: 30, width: 600, height: 50 });
          titleBox.textFrame.textRange.text = change.title;
          if (change.bullets?.length) {
            const body = slide.shapes.addTextBox(change.bullets.map((b) => `• ${b}`).join("\n"), {
              left: 40,
              top: 100,
              width: 600,
              height: 300
            });
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
          // PowerPoint's JavaScript API types do not expose slide notes.
        } else if (change.op === "duplicateSlide") {
          const src = context.presentation.slides.getItemAt(change.slideIndex);
          src.shapes.load("items/textFrame/textRange/text");
          await context.sync();
          const slides = context.presentation.slides;
          slides.add();
          slides.load("items");
          await context.sync();
          const copy = slides.items[slides.items.length - 1];
          const box = copy.shapes.addTextBox("", { left: 40, top: 30, width: 600, height: 300 });
          box.textFrame.textRange.text = src.shapes.items
            .map((s) => {
              try {
                return s.textFrame.textRange.text;
              } catch {
                return "";
              }
            })
            .filter(Boolean)
            .join("\n");
        } else if (change.op === "addChart") {
          const slide = context.presentation.slides.getItemAt(change.slideIndex);
          const box = slide.shapes.addTextBox(`${change.chartType}: ${change.categories.join(", ")}`, {
            left: 40,
            top: 180,
            width: 600,
            height: 200
          });
          box.textFrame.textRange.text = `${change.chartType}: ${change.categories.join(", ")}`;
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
