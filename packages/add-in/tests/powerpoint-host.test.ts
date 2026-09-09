import { Changeset } from "@openplugin/core";
import { PowerPointHost } from "@openplugin/host-powerpoint";
import { afterEach, describe, expect, it, vi } from "vitest";

function installPpt(ctx: unknown) {
  (globalThis as { PowerPoint?: unknown }).PowerPoint = {
    run: async (fn: (c: unknown) => Promise<unknown>) => fn(ctx)
  };
}

describe("PowerPointHost.getRawFacts", () => {
  afterEach(() => {
    delete (globalThis as { PowerPoint?: unknown; Office?: unknown }).PowerPoint;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("reports the selected slide index when getSelectedSlides is available", async () => {
    const slides = {
      load() {},
      items: [
        { id: "s0", shapes: { load() {}, items: [] } },
        { id: "s1", shapes: { load() {}, items: [] } },
        { id: "s2", shapes: { load() {}, items: [{ name: "Title 1", textFrame: { textRange: { text: "Agenda" } } }] } }
      ]
    };
    (globalThis as { Office?: unknown }).Office = {
      context: { document: { getSelectedDataAsync: (_c: unknown, cb: (r: { value: string }) => void) => cb({ value: "Agenda" }) } },
      CoercionType: { Text: "text" }
    };
    installPpt({
      presentation: {
        slides,
        getSelectedSlides: () => ({ load() {}, items: [{ id: "s2" }] })
      },
      sync: async () => undefined
    });
    const facts = await new PowerPointHost().getRawFacts();
    expect(facts.host).toBe("powerpoint");
    expect(facts.selection.slideIndex).toBe(2);
  });
});

describe("PowerPointHost.apply", () => {
  afterEach(() => {
    delete (globalThis as { PowerPoint?: unknown }).PowerPoint;
  });

  it("reorderSlides calls moveTo", async () => {
    const moveTo = vi.fn();
    const slide = { moveTo, shapes: { load() {}, items: [], addTextBox: vi.fn() } };
    installPpt({
      presentation: { slides: { getItemAt: () => slide, items: [slide], load() {}, add: vi.fn() } },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({ host: "powerpoint", op: "reorderSlides", from: 2, to: 0 });
    await new PowerPointHost().apply(cs);
    expect(moveTo).toHaveBeenCalledWith(0);
  });

  it("addSlide writes the title into the title placeholder", async () => {
    const title = { name: "Title 1", textFrame: { textRange: { text: "" } } };
    const subtitle = { name: "Subtitle 2", textFrame: { textRange: { text: "" } } };
    const addTextBox = vi.fn();
    const slide = { shapes: { load() {}, items: [title, subtitle], addTextBox } };
    const slides = {
      add: vi.fn(),
      load() {},
      items: [slide],
      getItemAt: () => slide
    };
    installPpt({
      presentation: { slides },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({ host: "powerpoint", op: "addSlide", title: "Q3 Review", bullets: ["Revenue up 12%"] });
    await new PowerPointHost().apply(cs);
    expect(title.textFrame.textRange.text).toBe("Q3 Review");
    expect(subtitle.textFrame.textRange.text).toContain("Revenue up 12%");
    expect(addTextBox).not.toHaveBeenCalled();
  });

  it("addChart throws when the shape API cannot add a chart", async () => {
    const addTextBox = vi.fn();
    const slide = { shapes: { load() {}, items: [], addTextBox } };
    installPpt({
      presentation: { slides: { getItemAt: () => slide, items: [slide], load() {}, add: vi.fn() } },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({
      host: "powerpoint",
      op: "addChart",
      slideIndex: 0,
      chartType: "bar",
      categories: ["Q1", "Q2", "Q3"],
      series: [{ name: "Rev", values: [10, 20, 30] }]
    });
    await expect(new PowerPointHost().apply(cs)).rejects.toThrow(/charts are not available/i);
    expect(addTextBox).not.toHaveBeenCalled();
  });

  it("duplicateSlide calls duplicate when the API exists", async () => {
    const duplicate = vi.fn();
    const src = {
      duplicate,
      shapes: {
        load() {},
        items: [
          { name: "Title 1", textFrame: { textRange: { text: "Agenda" } } },
          { name: "Content", textFrame: { textRange: { text: "• Wins" } } }
        ]
      }
    };
    installPpt({
      presentation: { slides: { getItemAt: () => src, items: [src], load() {}, add: vi.fn() } },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({ host: "powerpoint", op: "duplicateSlide", slideIndex: 0 });
    await new PowerPointHost().apply(cs);
    expect(duplicate).toHaveBeenCalled();
  });

  it("setNotes writes notes when notesSlide is present", async () => {
    const notesShape = { textFrame: { textRange: { text: "" } } };
    const slide = {
      notesSlide: { shapes: { load() {}, items: [notesShape] } },
      shapes: { load() {}, items: [], addTextBox: vi.fn() }
    };
    installPpt({
      presentation: { slides: { getItemAt: () => slide, items: [slide], load() {}, add: vi.fn() } },
      sync: async () => undefined
    });
    const cs = new Changeset();
    cs.add({ host: "powerpoint", op: "setNotes", slideIndex: 0, notes: "Welcome the team" });
    await new PowerPointHost().apply(cs);
    expect(notesShape.textFrame.textRange.text).toBe("Welcome the team");
  });
});
