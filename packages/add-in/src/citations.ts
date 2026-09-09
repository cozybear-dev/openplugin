import type { HostKind } from "@openplugin/core";

const CELL = /(?:([A-Za-z][\w]{0,30})!)?(\$?[A-Z]{1,3}\$?\d+)/g;
const SLIDE = /slide\s+(\d+)/gi;

export type Citation = { text: string; kind: "cell" | "slide"; sheet?: string; address?: string; slide?: number };

export function findCitations(text: string): Citation[] {
  const found: Citation[] = [];
  for (const m of text.matchAll(CELL)) {
    found.push({ text: m[0], kind: "cell", sheet: m[1], address: m[2].replace(/\$/g, "") });
  }
  for (const m of text.matchAll(SLIDE)) {
    found.push({ text: m[0], kind: "slide", slide: Number(m[1]) });
  }
  return found;
}

export function splitCitations(text: string): Array<{ text: string; citation?: Citation }> {
  const cites = findCitations(text);
  if (!cites.length) return [{ text }];
  const out: Array<{ text: string; citation?: Citation }> = [];
  let cursor = 0;
  for (const c of cites) {
    const idx = text.indexOf(c.text, cursor);
    if (idx < 0) continue;
    if (idx > cursor) out.push({ text: text.slice(cursor, idx) });
    out.push({ text: c.text, citation: c });
    cursor = idx + c.text.length;
  }
  if (cursor < text.length) out.push({ text: text.slice(cursor) });
  return out;
}

export async function jumpTo(host: HostKind, citation: Citation): Promise<void> {
  if (host === "excel" && citation.kind === "cell" && citation.address && typeof Excel !== "undefined") {
    await Excel.run(async (context) => {
      const sheet = citation.sheet
        ? context.workbook.worksheets.getItem(citation.sheet)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(citation.address!);
      range.select();
      await context.sync();
    });
    return;
  }
  if (host === "powerpoint" && citation.kind === "slide" && citation.slide && typeof PowerPoint !== "undefined") {
    await PowerPoint.run(async (context) => {
      context.presentation.slides.getItemAt(citation.slide! - 1);
      await context.sync();
    }).catch(() => undefined);
  }
}
