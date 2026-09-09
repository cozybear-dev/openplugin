import type { HostKind } from "../llm/types.js";

export type ExcelChange =
  | { host: "excel"; op: "writeRange"; sheet: string; address: string; values: unknown[][] }
  | { host: "excel"; op: "setFormulas"; sheet: string; address: string; formulas: string[][] }
  | { host: "excel"; op: "createTable"; sheet: string; address: string; name?: string }
  | { host: "excel"; op: "createChart"; sheet: string; source: string; chartType: string };

export type WordChange =
  | { host: "word"; op: "replaceSelection"; text: string }
  | {
      host: "word";
      op: "insertParagraphs";
      paragraphs: string[];
      location: "start" | "end" | "afterSelection";
    }
  | { host: "word"; op: "searchReplace"; search: string; replace: string; all: boolean }
  | { host: "word"; op: "applyStyle"; style: string; target: "selection" | "heading" }
  | { host: "word"; op: "insertTable"; rows: number; cols: number; cells?: string[][] }
  | { host: "word"; op: "insertComment"; text: string };

export type PptChange =
  | { host: "powerpoint"; op: "setShapeText"; slideIndex: number; shapeName?: string; text: string }
  | { host: "powerpoint"; op: "addSlide"; title: string; bullets?: string[] }
  | { host: "powerpoint"; op: "setNotes"; slideIndex: number; notes: string }
  | { host: "powerpoint"; op: "deleteSlide"; slideIndex: number };

export type Change = ExcelChange | WordChange | PptChange;

export class Changeset {
  readonly changes: Change[] = [];

  add(change: Change): void {
    this.changes.push(change);
  }

  isEmpty(): boolean {
    return this.changes.length === 0;
  }

  clear(): void {
    this.changes.length = 0;
  }

  preview(): string {
    if (this.changes.length === 0) return "No pending changes.";
    return this.changes.map((c, i) => `${i + 1}. ${describe(c)}`).join("\n");
  }

  forHost(host: HostKind): Change[] {
    return this.changes.filter((c) => c.host === host);
  }
}

function describe(change: Change): string {
  switch (change.op) {
    case "writeRange":
      return `excel writeRange ${change.sheet}!${change.address} ← ${summarizeGrid(change.values)}`;
    case "setFormulas":
      return `excel setFormulas ${change.sheet}!${change.address}`;
    case "createTable":
      return `excel createTable ${change.sheet}!${change.address}`;
    case "createChart":
      return `excel createChart ${change.chartType} from ${change.source}`;
    case "replaceSelection":
      return `word replaceSelection ← ${clip(change.text)}`;
    case "insertParagraphs":
      return `word insertParagraphs (${change.location}) ← ${clip(change.paragraphs.join(" / "))}`;
    case "searchReplace":
      return `word searchReplace "${change.search}" → "${change.replace}"`;
    case "applyStyle":
      return `word applyStyle ${change.style}`;
    case "insertTable":
      return `word insertTable ${change.rows}x${change.cols}`;
    case "insertComment":
      return `word insertComment ← ${clip(change.text)}`;
    case "setShapeText":
      return `ppt setShapeText slide ${change.slideIndex} ← ${clip(change.text)}`;
    case "addSlide":
      return `ppt addSlide "${change.title}"`;
    case "setNotes":
      return `ppt setNotes slide ${change.slideIndex}`;
    case "deleteSlide":
      return `ppt deleteSlide ${change.slideIndex}`;
  }
}

function summarizeGrid(values: unknown[][]): string {
  const first = values[0]?.[0];
  return clip(first == null ? `${values.length} rows` : String(first));
}

function clip(text: string, n = 80): string {
  const t = text.replace(/\s+/g, " ").trim();
  return t.length > n ? `${t.slice(0, n)}…` : t;
}
