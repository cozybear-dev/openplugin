export type SlashParse =
  | { kind: "none"; rest: string }
  | { kind: "list" }
  | { kind: "skill"; name: string; rest: string }
  | { kind: "unknown"; name: string };

const SLASH = /^\/([a-z0-9-]*)(?:\s+([\s\S]*))?$/;

export function parseSlash(input: string, known: string[]): SlashParse {
  const trimmed = input.trim();
  if (!trimmed.startsWith("/")) return { kind: "none", rest: input };
  const m = trimmed.match(SLASH);
  if (!m) return { kind: "none", rest: input };
  const name = m[1];
  const rest = (m[2] ?? "").trim();
  if (!name) return { kind: "list" };
  if (known.includes(name)) return { kind: "skill", name, rest };
  return { kind: "unknown", name };
}

export function slashSuggestions(input: string, known: string[]): string[] {
  const trimmed = input.trim();
  if (!trimmed.startsWith("/")) return [];
  const prefix = trimmed.slice(1).split(/\s/)[0] ?? "";
  return known.filter((name) => name.startsWith(prefix));
}
