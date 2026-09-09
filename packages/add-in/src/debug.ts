export type DebugEvent = {
  ts: string;
  level: "debug" | "info" | "warn" | "error";
  source: "agent" | "llm" | "host" | "search" | "apply" | "ui";
  message: string;
  data?: Record<string, unknown>;
};

const MAX = 500;
const KEY = "openplugin.debug";
const listeners = new Set<(events: DebugEvent[]) => void>();
let events: DebugEvent[] = [];
let persist = false;

export function loadDebug(): DebugEvent[] {
  try {
    const raw = sessionStorage.getItem(KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as DebugEvent[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

events = loadDebug();

export function debugLog(
  source: DebugEvent["source"],
  message: string,
  opts?: { level?: DebugEvent["level"]; data?: Record<string, unknown> }
): void {
  const entry: DebugEvent = {
    ts: new Date().toISOString(),
    level: opts?.level ?? "info",
    source,
    message,
    data: sanitize(opts?.data)
  };
  events = [entry, ...events].slice(0, MAX);
  if (persist) {
    try {
      sessionStorage.setItem(KEY, JSON.stringify(events));
    } catch {
      /* quota */
    }
  }
  for (const fn of listeners) fn(events);
}

export function subscribeDebug(fn: (events: DebugEvent[]) => void): () => void {
  listeners.add(fn);
  fn(events);
  return () => listeners.delete(fn);
}

export function getDebug(): DebugEvent[] {
  return events;
}

export function clearDebug(): void {
  events = [];
  try {
    sessionStorage.removeItem(KEY);
  } catch {
    /* ignore */
  }
  for (const fn of listeners) fn(events);
}

export function setDebugPersist(on: boolean): void {
  persist = on;
  if (on) {
    try {
      sessionStorage.setItem(KEY, JSON.stringify(events));
    } catch {
      /* ignore */
    }
  }
}

function sanitize(data?: Record<string, unknown>): Record<string, unknown> | undefined {
  if (!data) return undefined;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) {
    if (/key|token|secret|authorization/i.test(k)) out[k] = "[redacted]";
    else out[k] = v;
  }
  return out;
}
