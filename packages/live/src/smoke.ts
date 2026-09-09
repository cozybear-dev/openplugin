import { startLive, stopLive } from "./launch.js";
import type { LiveHost } from "./parse.js";
import { screenshotPath, withTaskpane } from "./session.js";
import { MOCK_LLM_ORIGIN } from "./mock-llm.js";

export async function runSmoke(host: LiveHost, flags: Record<string, string | boolean>): Promise<void> {
  const keep = Boolean(flags.keep);
  await startLive(host, { ...flags, force: flags.force ?? true });
  try {
    const shot = await screenshotPath(`smoke-${host}-pane`);
    const result = await withTaskpane(async (page) => {
      await page.waitForFunction(() => Boolean((window as { __openpluginLive?: unknown }).__openpluginLive), {
        timeout: 30_000
      });
      await page.screenshot({ path: shot, fullPage: true });
      const inOffice = await page.getAttribute("[data-testid='openplugin-app']", "data-in-office");
      if (inOffice !== "1") throw new Error(`Task pane is not in Office (data-in-office=${inOffice}).`);

      await page.evaluate(async ({ origin, kind }) => {
        const api = (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive;
        await api.configure({ baseUrl: origin, model: "openplugin-live", apiKey: "live" });
        if (kind === "excel") {
          const Excel = (globalThis as { Excel?: { run: (fn: (ctx: SeedCtx) => Promise<void>) => Promise<void> } }).Excel;
          if (!Excel) throw new Error("Excel.run missing");
          await Excel.run(async (ctx) => {
            ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C3").values = [
              ["Item", "Qty", "Price"],
              ["Apples", 3, 1.2],
              ["Pears", 2, 1.8]
            ];
            await ctx.sync();
          });
        }
        const prompt =
          kind === "word" ? "insert a hello paragraph" : kind === "powerpoint" ? "add a slide titled Outline" : "add a total";
        await api.send(prompt);
      }, { origin: MOCK_LLM_ORIGIN, kind: host });

      const start = Date.now();
      while (Date.now() - start < 90_000) {
        const busy = await page.evaluate(() => (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.getState().busy);
        if (!busy) break;
        await page.waitForTimeout(400);
      }
      const pending = await page.evaluate(() => (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.getState().pending);
      if (!pending.length) throw new Error("Mock LLM did not queue a changeset to apply.");
      await page.evaluate(async () => {
        await (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.apply();
      });
      await page.waitForTimeout(800);

      if (host === "excel") {
        const values = await page.evaluate(async () => {
          const Excel = (globalThis as { Excel?: { run: (fn: (ctx: SeedCtx) => Promise<unknown[][]>) => Promise<unknown[][]> } }).Excel;
          if (!Excel) throw new Error("Excel.run missing");
          return Excel.run(async (ctx) => {
            const range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A4");
            range.load("values");
            await ctx.sync();
            return range.values as unknown[][];
          });
        });
        if (String(values?.[0]?.[0] ?? "") !== "Total") {
          throw new Error(`Expected A4 to be Total after apply, got ${JSON.stringify(values)}`);
        }
        await page.evaluate(async () => {
          const Excel = (globalThis as { Excel?: { run: (fn: (ctx: ChartSeedCtx) => Promise<void>) => Promise<void> } }).Excel;
          if (!Excel) throw new Error("Excel.run missing");
          await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            sheet.charts.add("ColumnClustered", sheet.getRange("A1:B3"));
            await ctx.sync();
          });
        });
        await page.evaluate(async () => {
          await (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.send("what sheets exist");
        });
        const startFacts = Date.now();
        while (Date.now() - startFacts < 90_000) {
          const busy = await page.evaluate(() => (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.getState().busy);
          if (!busy) break;
          await page.waitForTimeout(400);
        }
        const afterChart = await page.evaluate(() => (window as unknown as { __openpluginLive: LiveApi }).__openpluginLive.getState().lastLine ?? "");
        if (/current selection is invalid/i.test(afterChart)) {
          throw new Error(`Excel turn failed after a chart was selected: ${afterChart}`);
        }
      }

      if (host === "word") {
        const paras = await page.evaluate(async () => {
          const Word = (globalThis as { Word?: { run: (fn: (ctx: WordSeedCtx) => Promise<string[]>) => Promise<string[]> } }).Word;
          if (!Word) throw new Error("Word.run missing");
          return Word.run(async (ctx) => {
            const items = ctx.document.body.paragraphs;
            items.load("items/text");
            await ctx.sync();
            return items.items.map((p) => p.text);
          });
        });
        if (!paras.some((p) => /Hello from OpenPlugin live/i.test(p))) {
          throw new Error(`Expected Word paragraph after apply, got ${JSON.stringify(paras)}`);
        }
      }

      if (host === "powerpoint") {
        const texts = await page.evaluate(async () => {
          const PowerPoint = (globalThis as { PowerPoint?: { run: (fn: (ctx: PptSeedCtx) => Promise<string[]>) => Promise<string[]> } }).PowerPoint;
          if (!PowerPoint) throw new Error("PowerPoint.run missing");
          return PowerPoint.run(async (ctx) => {
            const slides = ctx.presentation.slides;
            slides.load("items");
            await ctx.sync();
            const out: string[] = [];
            for (const slide of slides.items) {
              const shapes = slide.shapes;
              shapes.load("items/textFrame/textRange/text");
              await ctx.sync();
              for (const shape of shapes.items) {
                try {
                  out.push(shape.textFrame.textRange.text ?? "");
                } catch {
                  out.push("");
                }
              }
            }
            return out;
          });
        });
        if (!texts.some((t) => /Outline/i.test(t))) {
          throw new Error(`Expected PowerPoint shape text Outline after apply, got ${JSON.stringify(texts)}`);
        }
      }
      return { screenshot: shot, pending: pending.length };
    });
    console.log(JSON.stringify({ ok: true, host, ...result }, null, 2));
  } finally {
    if (!keep) await stopLive(host, { force: true });
  }
}

type LiveApi = {
  getState: () => { busy: boolean; pending: Array<{ id: string }>; lastLine?: string };
  configure: (p: { baseUrl: string; model: string; apiKey?: string }) => Promise<void>;
  send: (text: string) => Promise<void>;
  apply: () => Promise<void>;
};

type SeedCtx = {
  workbook: { worksheets: { getActiveWorksheet: () => { getRange: (a: string) => { values: unknown; load: (p: string) => void } } } };
  sync: () => Promise<void>;
};

type ChartSeedCtx = {
  workbook: {
    worksheets: {
      getActiveWorksheet: () => {
        getRange: (a: string) => unknown;
        charts: { add: (type: string, source: unknown) => void };
      };
    };
  };
  sync: () => Promise<void>;
};

type WordSeedCtx = {
  document: { body: { paragraphs: { load: (p: string) => void; items: Array<{ text: string }> } } };
  sync: () => Promise<void>;
};

type PptSeedCtx = {
  presentation: {
    slides: {
      load: (p: string) => void;
      items: Array<{
        shapes: {
          load: (p: string) => void;
          items: Array<{ textFrame: { textRange: { text: string } } }>;
        };
      }>;
    };
  };
  sync: () => Promise<void>;
};
