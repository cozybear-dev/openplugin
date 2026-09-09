import {
  truncateThread,
  type ChatMessage,
  type HostKind,
  type RestorePoint
} from "@openplugin/core";
import { CONVERSATIONS_STORE, openOpenpluginDb } from "./idb";
import { clearHistory, loadHistory, type StoredLine } from "./history";

export type ConversationRecord = {
  id: string;
  host: HostKind;
  title: string;
  titleSource: "auto" | "user";
  createdAt: string;
  updatedAt: string;
  documentTitle?: string;
  documentUrl?: string;
  lines: StoredLine[];
  messages: ChatMessage[];
  restorePoints: RestorePoint[];
};

export const MAX_CONVERSATIONS = 40;
export const ACTIVE_KEY_PREFIX = "openplugin.activeConversation.";

export function emptyConversation(host: HostKind, now = new Date().toISOString()): ConversationRecord {
  const id =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `${now}-${Math.random().toString(16).slice(2)}`;
  return {
    id,
    host,
    title: "New chat",
    titleSource: "auto",
    createdAt: now,
    updatedAt: now,
    lines: [],
    messages: [],
    restorePoints: []
  };
}

export function autoTitleFromLines(lines: StoredLine[], max = 60): string {
  const user = lines.find((line) => line.kind === "user");
  if (!user || user.kind !== "user") return "New chat";
  const first = user.text.trim().split("\n")[0] ?? "";
  if (!first) return "New chat";
  if (first.length <= max) return first;
  return `${first.slice(0, max - 1)}…`;
}

export function conversationHasContent(lines: StoredLine[]): boolean {
  return lines.some((line) => line.kind === "user" || line.kind === "assistant");
}

export function applyAutoTitle(record: ConversationRecord, lines: StoredLine[]): ConversationRecord {
  if (record.titleSource === "user") return record;
  return { ...record, title: autoTitleFromLines(lines) };
}

export function capConversations(
  list: ConversationRecord[],
  activeId: string,
  max = MAX_CONVERSATIONS
): ConversationRecord[] {
  const sorted = [...list].sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
  const active = sorted.find((c) => c.id === activeId);
  const rest = sorted.filter((c) => c.id !== activeId);
  const room = active ? max - 1 : max;
  const trimmed = rest.slice(0, Math.max(0, room));
  return active ? [active, ...trimmed] : trimmed;
}

export function visibleConversations(list: ConversationRecord[], activeId: string): ConversationRecord[] {
  return list
    .filter((c) => c.id === activeId || conversationHasContent(c.lines))
    .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
}

export function conversationFromLegacy(
  host: HostKind,
  legacy: { lines: StoredLine[]; messages: ChatMessage[]; restorePoints: RestorePoint[] }
): ConversationRecord {
  const base = emptyConversation(host);
  return {
    ...base,
    title: autoTitleFromLines(legacy.lines),
    lines: legacy.lines,
    messages: legacy.messages,
    restorePoints: legacy.restorePoints
  };
}

export function loadActiveId(host: HostKind): string | null {
  try {
    const raw = localStorage.getItem(`${ACTIVE_KEY_PREFIX}${host}`);
    return raw && raw.trim() ? raw : null;
  } catch {
    return null;
  }
}

export function saveActiveId(host: HostKind, id: string): void {
  try {
    localStorage.setItem(`${ACTIVE_KEY_PREFIX}${host}`, id);
  } catch {
    /* quota */
  }
}

function asRecord(raw: unknown): ConversationRecord | null {
  if (!raw || typeof raw !== "object") return null;
  const value = raw as Partial<ConversationRecord>;
  if (typeof value.id !== "string" || !value.id) return null;
  if (value.host !== "excel" && value.host !== "word" && value.host !== "powerpoint") return null;
  return {
    id: value.id,
    host: value.host,
    title: typeof value.title === "string" && value.title ? value.title : "New chat",
    titleSource: value.titleSource === "user" ? "user" : "auto",
    createdAt: typeof value.createdAt === "string" ? value.createdAt : new Date().toISOString(),
    updatedAt: typeof value.updatedAt === "string" ? value.updatedAt : new Date().toISOString(),
    documentTitle: typeof value.documentTitle === "string" ? value.documentTitle : undefined,
    documentUrl: typeof value.documentUrl === "string" ? value.documentUrl : undefined,
    lines: Array.isArray(value.lines) ? value.lines : [],
    messages: Array.isArray(value.messages) ? value.messages : [],
    restorePoints: Array.isArray(value.restorePoints) ? value.restorePoints : []
  };
}

export async function listConversations(host: HostKind): Promise<ConversationRecord[]> {
  if (typeof indexedDB === "undefined") return [];
  const db = await openOpenpluginDb();
  const rows = await new Promise<unknown[]>((resolve, reject) => {
    const tx = db.transaction(CONVERSATIONS_STORE, "readonly");
    const req = tx.objectStore(CONVERSATIONS_STORE).getAll();
    req.onsuccess = () => resolve((req.result as unknown[]) ?? []);
    req.onerror = () => reject(req.error);
  });
  return rows.map(asRecord).filter((row): row is ConversationRecord => row != null && row.host === host);
}

export async function deleteConversation(id: string): Promise<void> {
  if (typeof indexedDB === "undefined") return;
  const db = await openOpenpluginDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(CONVERSATIONS_STORE, "readwrite");
    tx.objectStore(CONVERSATIONS_STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function saveConversation(record: ConversationRecord): Promise<void> {
  const truncated = truncateThread({
    lines: record.lines,
    messages: record.messages,
    restorePoints: record.restorePoints,
    maxUserTurns: 20
  });
  const toSave: ConversationRecord = { ...record, ...truncated };
  if (typeof indexedDB === "undefined") return;
  const db = await openOpenpluginDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(CONVERSATIONS_STORE, "readwrite");
    tx.objectStore(CONVERSATIONS_STORE).put(toSave);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  const listed = await listConversations(record.host);
  const kept = new Set(capConversations(listed, record.id).map((c) => c.id));
  for (const row of listed) {
    if (!kept.has(row.id)) await deleteConversation(row.id);
  }
}

export async function migrateLegacyHistory(host: HostKind): Promise<ConversationRecord | null> {
  const existing = await listConversations(host);
  if (existing.length) return null;
  const legacy = loadHistory(host);
  if (!legacy.lines.length && !legacy.messages.length) return null;
  const record = conversationFromLegacy(host, legacy);
  await saveConversation(record);
  saveActiveId(host, record.id);
  await clearHistory(host);
  return record;
}
