import {
  truncateThread,
  type ActivityKind,
  type ChatMessage,
  type HostKind,
  type LineCheckpoint,
  type RestorePoint
} from "@openplugin/core";

const PREFIX = "openplugin.history.";

export type StoredLine =
  | { kind: "user"; text: string; checkpoint?: LineCheckpoint }
  | { kind: "assistant"; text: string }
  | { kind: "tool"; text: string }
  | {
      kind: "activity";
      phase: "start" | "done" | "error";
      activity: ActivityKind;
      name: string;
      label: string;
      detail?: string;
    }
  | { kind: "error"; text: string };

function storageKey(host: HostKind): string {
  return `${PREFIX}${host}`;
}

function docStore(): Office.Settings | Storage | null {
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

export function loadHistory(host: HostKind): {
  lines: StoredLine[];
  messages: ChatMessage[];
  restorePoints: RestorePoint[];
} {
  const store = docStore();
  if (!store) return { lines: [], messages: [], restorePoints: [] };
  try {
    const raw =
      "get" in store ? (store.get(storageKey(host)) as string | undefined) : store.getItem(storageKey(host));
    if (!raw || typeof raw !== "string") return { lines: [], messages: [], restorePoints: [] };
    const parsed = JSON.parse(raw) as {
      lines?: StoredLine[];
      messages?: ChatMessage[];
      restorePoints?: RestorePoint[];
    };
    return {
      lines: parsed.lines ?? [],
      messages: parsed.messages ?? [],
      restorePoints: parsed.restorePoints ?? []
    };
  } catch {
    return { lines: [], messages: [], restorePoints: [] };
  }
}

export async function saveHistory(
  host: HostKind,
  lines: StoredLine[],
  messages: ChatMessage[],
  restorePoints: RestorePoint[] = []
): Promise<void> {
  const truncated = truncateThread({
    lines,
    messages,
    restorePoints,
    maxUserTurns: 20
  });
  const payload = JSON.stringify(truncated);
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
  await saveHistory(host, [], [], []);
}
