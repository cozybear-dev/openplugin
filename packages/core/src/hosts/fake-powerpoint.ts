import type { Changeset } from "../tools/changeset.js";
import type { HostAdapter, PptSelection, PptSlide, RawFacts } from "./types.js";

export class FakePowerPointHost implements HostAdapter {
  readonly kind = "powerpoint" as const;
  title = "Presentation";
  slides: PptSlide[] = [{ title: "Slide 1", shapes: [{ name: "Title", text: "Slide 1" }], notes: "" }];
  selection: PptSelection = { slideIndex: 0, shapeName: "Title", text: "Slide 1" };

  async getRawFacts(): Promise<RawFacts> {
    const current = this.slides[this.selection.slideIndex];
    return {
      host: "powerpoint",
      title: this.title,
      slides: this.slides.map((s, index) => ({
        index,
        title: s.title,
        shapeCount: s.shapes.length
      })),
      selection: this.selection,
      currentNotes: current?.notes
    };
  }

  async readSelectionText(): Promise<string> {
    return this.selection.text;
  }

  async readShapeText(args: { slideIndex: number; shapeName?: string }): Promise<string> {
    const slide = this.slides[args.slideIndex];
    if (!slide) return "";
    const shape = args.shapeName
      ? slide.shapes.find((s) => s.name === args.shapeName)
      : slide.shapes[0];
    return shape?.text ?? slide.title;
  }

  async readNotes(slideIndex: number): Promise<string> {
    return this.slides[slideIndex]?.notes ?? "";
  }

  async apply(changeset: Changeset): Promise<void> {
    for (const change of changeset.forHost("powerpoint")) {
      if (change.op === "addSlide") {
        this.slides.push({
          title: change.title,
          shapes: [
            { name: "Title", text: change.title },
            ...(change.bullets ?? []).map((b, i) => ({ name: `Bullet${i}`, text: b }))
          ],
          notes: ""
        });
      } else if (change.op === "setShapeText") {
        const slide = this.slides[change.slideIndex];
        if (!slide) continue;
        const shape =
          slide.shapes.find((s) => s.name === change.shapeName) ?? slide.shapes[0];
        if (shape) shape.text = change.text;
        if (!change.shapeName || change.shapeName === "Title") slide.title = change.text;
      } else if (change.op === "setNotes") {
        const slide = this.slides[change.slideIndex];
        if (slide) slide.notes = change.notes;
      } else if (change.op === "deleteSlide") {
        this.slides.splice(change.slideIndex, 1);
      }
    }
  }
}
