import type { HostKind } from "../llm/types.js";
import { hunksForChange, invertChange } from "./diff.js";

export type ExcelChange =
  | {
      host: "excel";
      op: "writeRange";
      sheet: string;
      address: string;
      values: unknown[][];
      before?: unknown[][];
    }
  | {
      host: "excel";
      op: "setFormulas";
      sheet: string;
      address: string;
      formulas: string[][];
      before?: unknown[][];
    }
  | { host: "excel"; op: "createTable"; sheet: string; address: string; name?: string }
  | { host: "excel"; op: "createChart"; sheet: string; source: string; chartType: string }
  | {
      host: "excel";
      op: "formatRange";
      sheet: string;
      address: string;
      bold?: boolean;
      numberFormat?: string;
    }
  | {
      host: "excel";
      op: "clearRange";
      sheet: string;
      address: string;
      clearType: "contents" | "formats" | "all";
      before?: unknown[][];
    }
  | {
      host: "excel";
      op: "copyRange";
      sheet: string;
      source: string;
      dest: string;
      before?: unknown[][];
    }
  | {
      host: "excel";
      op: "modifySheet";
      sheet: string;
      operation: "insert" | "delete" | "hide" | "unhide" | "freeze" | "unfreeze";
      dimension?: "rows" | "columns";
      reference?: string;
      count?: number;
    }
  | {
      host: "excel";
      op: "modifyWorkbook";
      operation: "create" | "delete" | "rename" | "duplicate";
      sheet?: string;
      newName?: string;
    }
  | { host: "excel"; op: "resizeRange"; sheet: string; address: string; columnWidth?: number; rowHeight?: number }
  | {
      host: "excel";
      op: "createPivot";
      sheet: string;
      source: string;
      dest: string;
      rows: string[];
      columns?: string[];
      values: Array<{ field: string; summarizeBy?: string }>;
    };

export type WordChange =
  | { host: "word"; op: "replaceSelection"; text: string; beforeText?: string }
  | {
      host: "word";
      op: "insertParagraphs";
      paragraphs: string[];
      location: "start" | "end" | "afterSelection";
      beforeText?: string;
    }
  | { host: "word"; op: "searchReplace"; search: string; replace: string; all: boolean }
  | { host: "word"; op: "applyStyle"; style: string; target: "selection" | "heading" }
  | { host: "word"; op: "insertTable"; rows: number; cols: number; cells?: string[][] }
  | { host: "word"; op: "insertComment"; text: string }
  | { host: "word"; op: "replaceParagraph"; index: number; text: string; beforeText?: string }
  | { host: "word"; op: "replyComment"; commentIndex: number; text: string }
  | { host: "word"; op: "resolveComment"; commentIndex: number };

export type PptChange =
  | {
      host: "powerpoint";
      op: "setShapeText";
      slideIndex: number;
      shapeName?: string;
      text: string;
      beforeText?: string;
    }
  | { host: "powerpoint"; op: "addSlide"; title: string; bullets?: string[] }
  | { host: "powerpoint"; op: "setNotes"; slideIndex: number; notes: string; beforeText?: string }
  | { host: "powerpoint"; op: "deleteSlide"; slideIndex: number }
  | { host: "powerpoint"; op: "duplicateSlide"; slideIndex: number }
  | { host: "powerpoint"; op: "reorderSlides"; from: number; to: number }
  | {
      host: "powerpoint";
      op: "addChart";
      slideIndex: number;
      chartType: string;
      categories: string[];
      series: Array<{ name: string; values: number[] }>;
    };

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

  previewItems(): Array<{ title: string; detail: string }> {
    return this.changes.map((c) => item(c));
  }

  preview(): string {
    if (this.changes.length === 0) return "No pending changes.";
    return this.previewItems()
      .map((row, i) => `${i + 1}. ${row.title}${row.detail ? ` — ${row.detail}` : ""}`)
      .join("\n");
  }

  forHost(host: HostKind): Change[] {
    return this.changes.filter((c) => c.host === host);
  }

  diff() {
    return this.changes.map((c, i) => hunksForChange(c, i));
  }

  inverse(): Changeset {
    const next = new Changeset();
    for (const change of [...this.changes].reverse()) {
      const inv = invertChange(change);
      if (inv) next.add(inv);
    }
    return next;
  }

  filter(ids: Set<string>): Changeset {
    const next = new Changeset();
    this.changes.forEach((c, i) => {
      if (ids.has(`${c.op}-${i}`)) next.add(c);
    });
    return next;
  }

  reversible(): boolean {
    return this.changes.length > 0 && this.changes.every((c) => invertChange(c) != null);
  }
}

function item(change: Change): { title: string; detail: string } {
  switch (change.op) {
    case "writeRange":
      return { title: `Write ${change.sheet}!${change.address}`, detail: summarizeGrid(change.values) };
    case "setFormulas":
      return { title: `Formulas ${change.sheet}!${change.address}`, detail: "" };
    case "createTable":
      return { title: `Create table ${change.sheet}!${change.address}`, detail: change.name ?? "" };
    case "createChart":
      return { title: `Chart (${change.chartType})`, detail: change.source };
    case "formatRange":
      return { title: `Format ${change.sheet}!${change.address}`, detail: change.numberFormat ?? "" };
    case "clearRange":
      return { title: `Clear ${change.sheet}!${change.address}`, detail: change.clearType };
    case "copyRange":
      return { title: `Copy ${change.sheet}!${change.source}`, detail: change.dest };
    case "modifySheet":
      return { title: `${change.operation} ${change.dimension ?? ""}`.trim(), detail: change.sheet };
    case "modifyWorkbook":
      return { title: `${change.operation} sheet`, detail: change.newName ?? change.sheet ?? "" };
    case "resizeRange":
      return { title: `Resize ${change.sheet}!${change.address}`, detail: "" };
    case "createPivot":
      return { title: "Create pivot table", detail: change.source };
    case "replaceSelection":
      return { title: "Replace selection", detail: clip(change.text) };
    case "insertParagraphs":
      return { title: `Insert paragraphs (${change.location})`, detail: clip(change.paragraphs.join(" / ")) };
    case "searchReplace":
      return { title: "Find and replace", detail: `"${change.search}" → "${change.replace}"` };
    case "applyStyle":
      return { title: "Apply style", detail: change.style };
    case "insertTable":
      return { title: "Insert table", detail: `${change.rows}×${change.cols}` };
    case "insertComment":
      return { title: "Insert comment", detail: clip(change.text) };
    case "replaceParagraph":
      return { title: `Replace paragraph ${change.index + 1}`, detail: clip(change.text) };
    case "replyComment":
      return { title: "Reply to comment", detail: clip(change.text) };
    case "resolveComment":
      return { title: "Resolve comment", detail: String(change.commentIndex) };
    case "setShapeText":
      return { title: `Edit slide ${change.slideIndex + 1}`, detail: clip(change.text) };
    case "addSlide":
      return { title: "Add slide", detail: change.title };
    case "setNotes":
      return { title: `Speaker notes · slide ${change.slideIndex + 1}`, detail: clip(change.notes) };
    case "deleteSlide":
      return { title: `Delete slide ${change.slideIndex + 1}`, detail: "" };
    case "duplicateSlide":
      return { title: `Duplicate slide ${change.slideIndex + 1}`, detail: "" };
    case "reorderSlides":
      return { title: "Reorder slides", detail: `${change.from + 1} → ${change.to + 1}` };
    case "addChart":
      return { title: "Add chart", detail: change.chartType };
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
