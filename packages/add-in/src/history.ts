import type { ChatMessage, HostKind } from "@openplugin/core";

const PREFIX = "openplugin.history.";

export type StoredLine =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string }
  | { kind: "tool"; text: string }
  | { kind: "error"; text: string };

function storageKey(host: HostKind): string {
  return `${PREFIX}${host}`;
}

function docStore(): Office.DocumentSettings | Storage | null {
  try {
    if (Office?.context?.document?.settings) return Office.context.document.settings;
  } catch {
    /* ignore */
  }
  try {
    return localStorage;
  } catch {
    return null;
  }
}

export function loadHistory(host: HostKind): { lines: StoredLine[]; messages: ChatMessage[] } {
  const store = docStore();
  if (!store) return { lines: [], messages: [] };
  try {
    const raw =
      "get" in store ? (store.get(storageKey(host)) as string | undefined) : store.getItem(storageKey(host));
    if (!raw || typeof raw !== "string") return { lines: [], messages: [] };
    const parsed = JSON.parse(raw) as { lines?: StoredLine[]; messages?: ChatMessage[] };
    return { lines: parsed.lines ?? [], messages: parsed.messages ?? [] };
  } catch {
    return { lines: [], messages: [] };
  }
}

export async function saveHistory(
  host: HostKind,
  lines: StoredLine[],
  messages: ChatMessage[]
): Promise<void> {
  const payload = JSON.stringify({ lines: lines.slice(-80), messages: messages.slice(-40) });
  const store = docStore();
  if (!store) return;
  if ("set" in store && "saveAsync" in store) {
    store.set(storageKey(host), payload);
    await new Promise<void>((resolve) => store.saveAsync(() => resolve()));
    return;
  }
  store.setItem(storageKey(host), payload);
}

export async function clearHistory(host: HostKind): Promise<void> {
  await saveHistory(host, [], []);
}
