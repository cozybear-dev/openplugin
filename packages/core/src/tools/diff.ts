import { addressToBounds } from "../hosts/types.js";
import type { HostKind } from "../llm/types.js";
import type { Change } from "./changeset.js";

export type CellDelta = {
  address: string;
  before: unknown;
  after: unknown;
  changed: boolean;
};

export type TextPart = { type: "eq" | "del" | "ins"; text: string };

export type DiffHunk =
  | {
      kind: "grid";
      id: string;
      title: string;
      reversible: true;
      cells: CellDelta[];
      truncated: boolean;
    }
  | {
      kind: "text";
      id: string;
      title: string;
      reversible: boolean;
      parts: TextPart[];
      before: string;
      after: string;
    }
  | {
      kind: "note";
      id: string;
      title: string;
      reversible: false;
      detail: string;
    };

const CELL_CAP = 80;

export function colLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function cellDiff(
  sheet: string,
  address: string,
  before: unknown[][] | undefined,
  after: unknown[][]
): CellDelta[] {
  const { r1, c1 } = addressToBounds(address);
  const rows = Math.max(after.length, before?.length ?? 0);
  const cols = Math.max(
    after.reduce((m, row) => Math.max(m, row.length), 0),
    (before ?? []).reduce((m, row) => Math.max(m, row.length), 0)
  );
  const cells: CellDelta[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const b = before?.[r]?.[c] ?? "";
      const a = after[r]?.[c] ?? "";
      const label = `${sheet}!${colLetter(c1 + c)}${r1 + r + 1}`;
      cells.push({
        address: label,
        before: b,
        after: a,
        changed: String(b) !== String(a)
      });
    }
  }
  return cells;
}

export function textDiff(before: string, after: string): TextPart[] {
  const a = tokenize(before);
  const b = tokenize(after);
  const m = a.length;
  const n = b.length;
  const dp: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const parts: TextPart[] = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (a[i] === b[j]) {
      parts.push({ type: "eq", text: a[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      parts.push({ type: "del", text: a[i] });
      i++;
    } else {
      parts.push({ type: "ins", text: b[j] });
      j++;
    }
  }
  while (i < m) parts.push({ type: "del", text: a[i++] });
  while (j < n) parts.push({ type: "ins", text: b[j++] });
  return parts;
}

function tokenize(text: string): string[] {
  if (!text) return [];
  if (text.includes("\n") && text.length > 80) return text.split(/(\n)/).filter((t) => t.length > 0);
  return text.split(/(\s+)/).filter((t) => t.length > 0);
}

export function hunksForChange(change: Change, index: number): DiffHunk {
  const id = `${change.op}-${index}`;
  switch (change.op) {
    case "writeRange":
    case "setFormulas": {
      const after = change.op === "writeRange" ? change.values : change.formulas;
      const cells = cellDiff(change.sheet, change.address, change.before, after);
      return {
        kind: "grid",
        id,
        title: `${change.op === "setFormulas" ? "Formulas" : "Write"} ${change.sheet}!${change.address}`,
        reversible: true,
        cells: cells.slice(0, CELL_CAP),
        truncated: cells.length > CELL_CAP
      };
    }
    case "replaceSelection":
    case "setShapeText":
    case "setNotes": {
      const after =
        change.op === "replaceSelection" ? change.text : change.op === "setShapeText" ? change.text : change.notes;
      const before = change.beforeText ?? "";
      return {
        kind: "text",
        id,
        title:
          change.op === "replaceSelection"
            ? "Replace selection"
            : change.op === "setNotes"
              ? `Speaker notes · slide ${change.slideIndex + 1}`
              : `Edit slide ${change.slideIndex + 1}`,
        reversible: change.beforeText != null,
        parts: textDiff(before, after),
        before,
        after
      };
    }
    case "insertParagraphs":
      return {
        kind: "text",
        id,
        title: `Insert paragraphs (${change.location})`,
        reversible: false,
        parts: textDiff("", change.paragraphs.join("\n")),
        before: "",
        after: change.paragraphs.join("\n")
      };
    case "searchReplace":
      return {
        kind: "text",
        id,
        title: "Find and replace",
        reversible: false,
        parts: textDiff(change.search, change.replace),
        before: change.search,
        after: change.replace
      };
    default:
      return {
        kind: "note",
        id,
        title: describeNote(change),
        reversible: false,
        detail: noteDetail(change)
      };
  }
}

function describeNote(change: Change): string {
  switch (change.op) {
    case "createTable":
      return `Create table ${change.sheet}!${change.address}`;
    case "createChart":
      return `Create ${change.chartType} chart`;
    case "insertTable":
      return `Insert table ${change.rows}×${change.cols}`;
    case "insertComment":
      return "Insert comment";
    case "applyStyle":
      return `Apply style ${change.style}`;
    case "addSlide":
      return `Add slide “${change.title}”`;
    case "deleteSlide":
      return `Delete slide ${change.slideIndex + 1}`;
    default:
      return change.op;
  }
}

function noteDetail(change: Change): string {
  if (change.op === "addSlide") return (change.bullets ?? []).join(" · ");
  if (change.op === "insertComment") return change.text;
  if (change.op === "createChart") return change.source;
  return "This change cannot be reverted automatically.";
}

export function invertChange(change: Change): Change | null {
  switch (change.op) {
    case "writeRange":
      if (!change.before) return null;
      return {
        ...change,
        values: change.before,
        before: change.values
      };
    case "setFormulas":
      if (!change.before) return null;
      return {
        ...change,
        formulas: change.before.map((row) => row.map((c) => String(c ?? ""))),
        before: change.formulas
      };
    case "replaceSelection":
      if (change.beforeText == null) return null;
      return { ...change, text: change.beforeText, beforeText: change.text };
    case "setShapeText":
      if (change.beforeText == null) return null;
      return { ...change, text: change.beforeText, beforeText: change.text };
    case "setNotes":
      if (change.beforeText == null) return null;
      return { ...change, notes: change.beforeText, beforeText: change.notes };
    default:
      return null;
  }
}

export function hostOf(change: Change): HostKind {
  return change.host;
}
