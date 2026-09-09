export const MAP_CHUNK = 25;

export function chunkRows<T>(rows: T[], size = MAP_CHUNK): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < rows.length; i += size) out.push(rows.slice(i, i + size));
  return out;
}

export function serializeRange(values: unknown[][] | undefined): string {
  if (!values?.length) return "";
  return values
    .slice(0, 200)
    .map((row) => row.map((cell) => (cell == null ? "" : String(cell))).join("\t"))
    .join("\n");
}

export function cacheKey(parts: unknown[]): string {
  return JSON.stringify(parts);
}

export function buildPromptCall(prompt: string, range?: unknown[][]): { system: string; user: string } {
  const table = serializeRange(range);
  return {
    system: "You are OpenPlugin inside Excel. Reply with the answer only — no preamble.",
    user: table ? `${prompt}\n\nRange:\n${table}` : prompt
  };
}

export function buildMapCall(instruction: string, rows: unknown[][]): { system: string; user: string } {
  return {
    system: "You map each input row to one output. Return a JSON array of strings, one per row, same order. No markdown.",
    user: `${instruction}\n\nRows as JSON:\n${JSON.stringify(rows)}`
  };
}

export function parseMappedArray(text: string, expected: number): string[] {
  const json = extractJson(text);
  if (Array.isArray(json)) {
    return json.slice(0, expected).map((v) => (v == null ? "" : String(v)));
  }
  return text.split("\n").slice(0, expected);
}

export function buildExtractCall(schema: string, range?: unknown[][]): { system: string; user: string } {
  return {
    system: "Extract fields. Return a JSON array (one value per field, in schema order) or a JSON object. No markdown.",
    user: `Schema: ${schema}\n\nData:\n${serializeRange(range)}`
  };
}

export function flattenCells(range: unknown[][] | undefined): string[] {
  if (!range?.length) return [];
  return range.flatMap((row) =>
    (row ?? []).map((cell) => (cell == null ? "" : String(cell)))
  );
}

export function reshapeCells(values: string[], rows: number, cols: number): string[][] {
  const out: string[][] = [];
  let i = 0;
  for (let r = 0; r < rows; r++) {
    const row: string[] = [];
    for (let c = 0; c < cols; c++) row.push(values[i++] ?? "");
    out.push(row);
  }
  return out;
}

export function buildTranslateCall(
  target: string,
  cells: string[],
  source?: string
): { system: string; user: string } {
  const from = source?.trim() ? source.trim() : "auto-detect";
  return {
    system:
      "You translate each input cell. Return a JSON array of strings, one per cell, same order. Keep empty strings empty. Preserve meaning; do not add quotes, labels, or markdown.",
    user: `Target language: ${target}\nSource language: ${from}\n\nCells as JSON:\n${JSON.stringify(cells)}`
  };
}

export function parseExtract(text: string, schema: string): unknown[] {
  const fields = schema
    .split(/[,\n]/)
    .map((s) => s.trim())
    .filter(Boolean);
  const json = extractJson(text);
  if (Array.isArray(json)) return json;
  if (json && typeof json === "object") {
    return fields.map((f) => (json as Record<string, unknown>)[f] ?? "");
  }
  return fields.map(() => text);
}

function extractJson(text: string): unknown {
  const trimmed = text.trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  try {
    return JSON.parse(trimmed);
  } catch {
    const start = trimmed.search(/[\[{]/);
    if (start < 0) return null;
    try {
      return JSON.parse(trimmed.slice(start));
    } catch {
      return null;
    }
  }
}
