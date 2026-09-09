import type { Change, HostKind } from "@openplugin/core";

const KEY = "openplugin.revisions";

export type AppliedRevision = {
  id: string;
  at: string;
  host: HostKind;
  summary: string;
  inverse: Change[];
};

function store(): Office.DocumentSettings | Storage | null {
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

export function loadRevisions(): AppliedRevision[] {
  const s = store();
  if (!s) return [];
  try {
    const raw = "get" in s ? (s.get(KEY) as string | undefined) : s.getItem(KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as AppliedRevision[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export async function saveRevisions(list: AppliedRevision[]): Promise<void> {
  const raw = JSON.stringify(list.slice(0, 10));
  const s = store();
  if (!s) return;
  if ("set" in s && "saveAsync" in s) {
    s.set(KEY, raw);
    await new Promise<void>((resolve) => s.saveAsync(() => resolve()));
    return;
  }
  s.setItem(KEY, raw);
}

export function pushRevision(list: AppliedRevision[], rev: AppliedRevision): AppliedRevision[] {
  return [rev, ...list].slice(0, 10);
}
