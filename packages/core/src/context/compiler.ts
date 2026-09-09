import type { HostAdapter, RawFacts } from "../hosts/types.js";
import type { HostKind } from "../llm/types.js";

export type DocumentSnapshot = {
  host: HostKind;
  title: string;
  selection: unknown;
  outline: unknown;
  schema?: unknown;
  diffFromPrevious?: string;
  tokenEstimate: number;
};

export function estimateTokens(text: string): number {
  return Math.max(1, Math.ceil(text.length / 4));
}

function clip(text: string, maxChars: number): string {
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars)}…`;
}

function compact(facts: RawFacts): Omit<DocumentSnapshot, "tokenEstimate" | "diffFromPrevious"> {
  if (facts.host === "excel") {
    return {
      host: "excel",
      title: facts.title,
      selection: {
        sheet: facts.selection.sheet,
        address: facts.selection.address,
        values: facts.selection.values.slice(0, 30).map((row) => row.slice(0, 12))
      },
      outline: facts.sheets,
      schema: { namedRanges: facts.namedRanges ?? [], tables: facts.sheets.flatMap((s) => s.tables) }
    };
  }
  if (facts.host === "word") {
    return {
      host: "word",
      title: facts.title,
      selection: {
        text: clip(facts.selection.text, 800),
        style: facts.selection.style
      },
      outline: facts.headings.map((h) => ({ style: h.style, text: clip(h.text, 120) })),
      schema: { paragraphCount: facts.paragraphCount, surrounding: facts.surrounding.map((s) => clip(s, 240)) }
    };
  }
  return {
    host: "powerpoint",
    title: facts.title,
    selection: { ...facts.selection, text: clip(facts.selection.text, 400) },
    outline: facts.slides,
    schema: { currentNotes: facts.currentNotes ? clip(facts.currentNotes, 400) : "" }
  };
}

function fitBudget(
  snap: Omit<DocumentSnapshot, "tokenEstimate" | "diffFromPrevious">,
  budget: number
): DocumentSnapshot {
  let current = { ...snap, tokenEstimate: 0 };
  for (let i = 0; i < 8; i++) {
    const json = JSON.stringify({
      host: current.host,
      title: current.title,
      selection: current.selection,
      outline: current.outline,
      schema: current.schema
    });
    current.tokenEstimate = estimateTokens(json);
    if (current.tokenEstimate <= budget) return current;
    const factor = Math.max(0.4, budget / current.tokenEstimate);
    current = {
      ...current,
      selection: shrink(current.selection, factor),
      outline: shrink(current.outline, factor),
      schema: shrink(current.schema, factor)
    };
  }
  const fallback = {
    host: snap.host,
    title: snap.title,
    selection: shrink(snap.selection, 0.2),
    outline: Array.isArray(snap.outline) ? (snap.outline as unknown[]).slice(0, 8) : snap.outline,
    tokenEstimate: 0
  };
  fallback.tokenEstimate = Math.min(budget, estimateTokens(JSON.stringify(fallback)));
  return fallback;
}

function shrink(value: unknown, factor: number): unknown {
  if (typeof value === "string") return clip(value, Math.max(24, Math.floor(value.length * factor)));
  if (Array.isArray(value)) {
    const n = Math.max(1, Math.floor(value.length * Math.min(1, factor + 0.15)));
    return value.slice(0, n).map((item) => shrink(item, factor));
  }
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value)) out[k] = shrink(v, factor);
    return out;
  }
  return value;
}

export async function compileSnapshot(
  host: HostAdapter,
  opts: { tokenBudget?: number; previous?: DocumentSnapshot } = {}
): Promise<DocumentSnapshot> {
  const budget = opts.tokenBudget ?? 4000;
  const facts = await host.getRawFacts();
  const snap = fitBudget(compact(facts), budget);
  if (opts.previous) {
    const prev = JSON.stringify(opts.previous.selection);
    const next = JSON.stringify(snap.selection);
    if (prev !== next) {
      snap.diffFromPrevious = clip(`selection: ${next}`, 400);
      snap.tokenEstimate = Math.min(budget, snap.tokenEstimate + estimateTokens(snap.diffFromPrevious));
    } else {
      snap.diffFromPrevious = "selection unchanged";
    }
  }
  return snap;
}

export function snapshotToPrompt(snapshot: DocumentSnapshot): string {
  return [
    `Host: ${snapshot.host}`,
    `Title: ${snapshot.title}`,
    `Outline: ${JSON.stringify(snapshot.outline)}`,
    `Selection: ${JSON.stringify(snapshot.selection)}`,
    snapshot.schema ? `Schema: ${JSON.stringify(snapshot.schema)}` : "",
    snapshot.diffFromPrevious ? `Diff: ${snapshot.diffFromPrevious}` : ""
  ]
    .filter(Boolean)
    .join("\n");
}
