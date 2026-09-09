const KEY = "openplugin.audit";
const MAX = 100;

export type AuditEntry = {
  ts: string;
  action: "test" | "apply" | "reject" | "revert" | "skill" | "error";
  host: string;
  model?: string;
  summary: string;
  skill?: string;
};

export function loadAudit(): AuditEntry[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as AuditEntry[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function appendAuditLocal(entry: Omit<AuditEntry, "ts"> & { ts?: string }): AuditEntry[] {
  const next: AuditEntry = { ts: entry.ts ?? new Date().toISOString(), ...entry };
  const list = [next, ...loadAudit()].slice(0, MAX);
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    /* ignore */
  }
  return list;
}
