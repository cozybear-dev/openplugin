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
      let slideIndex = 0;
      try {
        const presentation = context.presentation as PowerPoint.Presentation & {
          getSelectedSlides?: () => { load: (p: string) => void; items: Array<{ id?: string }> };
        };
        if (typeof presentation.getSelectedSlides === "function") {
          const picked = presentation.getSelectedSlides();
          picked.load("items");
          slides.load("items/id");
          await context.sync();
          const id = picked.items[0]?.id;
          const idx = slides.items.findIndex((s) => (s as { id?: string }).id === id);
          if (idx >= 0) slideIndex = idx;
        }
      } catch {
        slideIndex = 0;
      }
      return {
        host: "powerpoint",
        title: "Presentation",
        slides: outline,
        selection: {
          slideIndex,
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
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex) as PowerPoint.Slide & {
        notesSlide?: { shapes: { load: (p: string) => void; items: Array<{ textFrame: { textRange: { text: string } } }> } };
      };
      const notes = slide.notesSlide;
      if (!notes) return "";
      notes.shapes.load("items/textFrame/textRange/text");
      await context.sync();
      return notes.shapes.items
        .map((s) => {
          try {
            return s.textFrame.textRange.text ?? "";
          } catch {
            return "";
          }
        })
        .filter(Boolean)
        .join("\n");
    });
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
          slide.shapes.load("items/name,items/textFrame/textRange/text");
          await context.sync();
          fillSlidePlaceholders(slide, change.title, change.bullets);
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
          await writeNotes(context, change.slideIndex, change.notes);
        } else if (change.op === "duplicateSlide") {
          const src = context.presentation.slides.getItemAt(change.slideIndex) as PowerPoint.Slide & {
            duplicate?: () => void;
          };
          if (typeof src.duplicate !== "function") {
            throw new Error("duplicateSlide is not supported in this PowerPoint build.");
          }
          src.duplicate();
        } else if (change.op === "reorderSlides") {
          const slide = context.presentation.slides.getItemAt(change.from) as PowerPoint.Slide & {
            moveTo?: (to: number) => void;
          };
          if (typeof slide.moveTo !== "function") {
            throw new Error("reorderSlides is not supported in this PowerPoint build.");
          }
          slide.moveTo(change.to);
        } else if (change.op === "addChart") {
          const slide = context.presentation.slides.getItemAt(change.slideIndex);
          const shapes = slide.shapes as PowerPoint.ShapeCollection & {
            addChart?: (...args: unknown[]) => unknown;
          };
          if (typeof shapes.addChart !== "function") {
            throw new Error("PowerPoint charts are not available in this Office.js build.");
          }
          shapes.addChart(change.chartType, change.categories, change.series);
        }
      }
      await context.sync();
    });
  }
}

function setShapeText(shape: { textFrame?: { textRange?: { text: string } } } | undefined, text: string): boolean {
  try {
    if (!shape?.textFrame?.textRange) return false;
    shape.textFrame.textRange.text = text;
    return true;
  } catch {
    return false;
  }
}

function fillSlidePlaceholders(
  slide: PowerPoint.Slide,
  title: string,
  bullets?: string[]
): void {
  const shapes = slide.shapes.items;
  const titleShape = shapes.find((s) => /^title\b/i.test(s.name));
  const bodyShape = shapes.find(
    (s) => /subtitle|content|^body\b/i.test(s.name) && !/^title\b/i.test(s.name)
  );
  if (!setShapeText(titleShape, title)) {
    const box = slide.shapes.addTextBox(title, { left: 40, top: 30, width: 600, height: 50 });
    box.textFrame.textRange.text = title;
  }
  if (bullets?.length) {
    const text = bullets.map((b) => `• ${b}`).join("\n");
    if (!setShapeText(bodyShape, text)) {
      const box = slide.shapes.addTextBox(text, { left: 40, top: 100, width: 600, height: 300 });
      box.textFrame.textRange.text = text;
    }
  }
}

async function writeNotes(context: PowerPoint.RequestContext, slideIndex: number, notes: string): Promise<void> {
  const slide = context.presentation.slides.getItemAt(slideIndex) as PowerPoint.Slide & {
    notesSlide?: { shapes: { load: (p: string) => void; items: Array<{ textFrame: { textRange: { text: string } } }> } };
  };
  const notesSlide = slide.notesSlide;
  if (!notesSlide) {
    throw new Error("Speaker notes are not available in this PowerPoint build.");
  }
  notesSlide.shapes.load("items/textFrame/textRange/text");
  await context.sync();
  const shape = notesSlide.shapes.items[0];
  if (!setShapeText(shape, notes)) {
    throw new Error("Speaker notes are not available in this PowerPoint build.");
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
