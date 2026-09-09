const KEY = "openplugin.audit";
const MAX = 100;

export type AuditEntry = {
  ts: string;
  action: "test" | "apply" | "reject" | "revert" | "restore" | "skill" | "search" | "error";
  host: string;
  program?: string;
  documentTitle?: string;
  documentUrl?: string;
  model?: string;
  endpoint?: string;
  summary: string;
  skill?: string;
  tools?: string[];
  ops?: string[];
  searchBackend?: string;
  tokenEstimate?: number;
};

export function programName(host: string): string {
  if (host === "excel") return "Excel";
  if (host === "word") return "Word";
  if (host === "powerpoint") return "PowerPoint";
  return host;
}

export function documentIdentity(): { documentTitle?: string; documentUrl?: string } {
  try {
    const url = Office?.context?.document?.url;
    const title =
      (typeof url === "string" && url.split(/[/\\]/).pop()) ||
      (typeof document !== "undefined" ? document.title : undefined);
    return { documentTitle: title, documentUrl: typeof url === "string" ? url : undefined };
  } catch {
    return {};
  }
}

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
