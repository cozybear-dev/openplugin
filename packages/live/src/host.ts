import type { LiveCommand } from "./parse.js";
import { withTaskpane } from "./session.js";

export async function runHost(cmd: LiveCommand): Promise<void> {
  const sub = cmd.sub;
  if (sub === "facts") {
    const facts = await withTaskpane((page) => page.evaluate(readFacts));
    console.log(JSON.stringify(facts, null, 2));
    return;
  }
  if (sub === "range") {
    const sheet = cmd.rest.length > 1 ? cmd.rest[0] : undefined;
    const address = cmd.rest.length > 1 ? cmd.rest[1] : cmd.rest[0];
    if (!address) throw new Error("Usage: live excel host range [sheet] A1:C3");
    const values = await withTaskpane((page) => page.evaluate(readRange, { sheet, address }));
    console.log(JSON.stringify({ sheet: sheet ?? null, address, values }, null, 2));
    return;
  }
  if (sub === "eval") {
    const code = cmd.rest.join(" ");
    if (!code) throw new Error("Usage: live <host> host eval <javascript returning a value>");
    const result = await withTaskpane((page) =>
      page.evaluate(async (src) => {
        const fn = new Function(`return (async () => { ${src} })()`);
        return await fn();
      }, code)
    );
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  if (sub === "seed") {
    const result = await withTaskpane((page) => page.evaluate(seedHost, cmd.host));
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  throw new Error(`Unknown host command '${sub}'. Use facts|range|eval|seed`);
}

async function readFacts(): Promise<unknown> {
  const g = globalThis as {
    Excel?: { run: (fn: (ctx: ExcelCtx) => Promise<unknown>) => Promise<unknown> };
    Word?: { run: (fn: (ctx: WordCtx) => Promise<unknown>) => Promise<unknown> };
    PowerPoint?: { run: (fn: (ctx: PptCtx) => Promise<unknown>) => Promise<unknown> };
  };
  if (g.Excel) {
    return g.Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      const selected = ctx.workbook.getSelectedRange();
      selected.load(["address", "values", "worksheet/name"]);
      await ctx.sync();
      return {
        host: "excel",
        sheets: sheets.items.map((s) => s.name),
        selection: {
          sheet: selected.worksheet.name,
          address: selected.address,
          values: selected.values
        }
      };
    });
  }
  if (g.Word) {
    return g.Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text");
      await ctx.sync();
      return { host: "word", paragraphs: paras.items.map((p) => p.text).slice(0, 20) };
    });
  }
  if (g.PowerPoint) {
    return g.PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      return { host: "powerpoint", slideCount: slides.items.length };
    });
  }
  throw new Error("Office.js host APIs are not available in this page.");
}

async function readRange(args: { sheet?: string; address: string }): Promise<unknown[][]> {
  const Excel = (globalThis as { Excel?: { run: (fn: (ctx: ExcelCtx) => Promise<unknown[][]>) => Promise<unknown[][]> } }).Excel;
  if (!Excel) throw new Error("Excel.run is not available.");
  return Excel.run(async (ctx) => {
    const sheet = args.sheet
      ? ctx.workbook.worksheets.getItem(args.sheet)
      : ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(args.address);
    range.load("values");
    await ctx.sync();
    return range.values as unknown[][];
  });
}

async function seedHost(host: string): Promise<unknown> {
  if (host === "excel") {
    const Excel = (globalThis as { Excel?: { run: (fn: (ctx: ExcelCtx) => Promise<unknown>) => Promise<unknown> } }).Excel;
    if (!Excel) throw new Error("Excel.run is not available.");
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C3");
      range.values = [
        ["Item", "Qty", "Price"],
        ["Apples", 3, 1.2],
        ["Pears", 2, 1.8]
      ];
      await ctx.sync();
      return { seeded: "A1:C3" };
    });
  }
  if (host === "word") {
    const Word = (globalThis as { Word?: { run: (fn: (ctx: WordCtx) => Promise<unknown>) => Promise<unknown> } }).Word;
    if (!Word) throw new Error("Word.run is not available.");
    return Word.run(async (ctx) => {
      ctx.document.body.clear();
      ctx.document.body.insertParagraph("OpenPlugin live fixture", "Start");
      await ctx.sync();
      return { seeded: "paragraph" };
    });
  }
  const PowerPoint = (globalThis as { PowerPoint?: { run: (fn: (ctx: PptCtx) => Promise<unknown>) => Promise<unknown> } })
    .PowerPoint;
  if (!PowerPoint) throw new Error("PowerPoint.run is not available.");
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    return { seeded: "presentation", slideCount: slides.items.length };
  });
}

type ExcelCtx = {
  workbook: {
    worksheets: {
      load: (p: string) => void;
      items: Array<{ name: string }>;
      getItem: (name: string) => ExcelSheet;
      getActiveWorksheet: () => ExcelSheet;
    };
    getSelectedRange: () => {
      load: (p: string[]) => void;
      address: string;
      values: unknown[][];
      worksheet: { name: string };
    };
  };
  sync: () => Promise<void>;
};

type ExcelSheet = {
  getRange: (address: string) => { load: (p: string) => void; values: unknown[][]; };
};

type WordCtx = {
  document: {
    body: {
      paragraphs: { load: (p: string) => void; items: Array<{ text: string }> };
      clear: () => void;
      insertParagraph: (text: string, loc: string) => void;
    };
  };
  sync: () => Promise<void>;
};

type PptCtx = {
  presentation: { slides: { load: (p: string) => void; items: unknown[] } };
  sync: () => Promise<void>;
};
