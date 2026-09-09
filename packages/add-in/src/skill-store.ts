import { parseSkillMarkdown, serializeSkill, type Skill } from "@openplugin/core";
import { openOpenpluginDb, SKILLS_STORE } from "./idb";

export type SkillOrigin = "imported" | "catalog";

export type StoredSkillRecord = Skill & {
  markdown: string;
  origin: SkillOrigin;
  updatedAt: string;
};

export type ManagedSkill = Omit<StoredSkillRecord, "origin"> & {
  origin: "bundled" | SkillOrigin;
  overridesBundled: boolean;
};

export function normalizeStoredSkill(raw: unknown): StoredSkillRecord | null {
  if (!raw || typeof raw !== "object") return null;
  const value = raw as Partial<Skill> & { markdown?: unknown; origin?: unknown; updatedAt?: unknown };
  if (typeof value.name !== "string" || !value.name) return null;
  if (typeof value.description !== "string" || !value.description) return null;
  if (typeof value.body !== "string") return null;
  const skill = value as Skill;
  return {
    ...skill,
    name: skill.name,
    description: skill.description,
    body: skill.body,
    hosts: Array.isArray(skill.hosts) ? skill.hosts : ["excel", "word", "powerpoint"],
    metadata: skill.metadata ?? {},
    rootPath: skill.rootPath || skill.name,
    userInvocable: skill.userInvocable !== false,
    disableModelInvocation: Boolean(skill.disableModelInvocation),
    inject: skill.inject === "always" ? "always" : "on-demand",
    markdown: typeof value.markdown === "string" && value.markdown.trim() ? value.markdown : serializeSkill(skill),
    origin: value.origin === "catalog" ? "catalog" : "imported",
    updatedAt: typeof value.updatedAt === "string" ? value.updatedAt : ""
  };
}

export function mergeStoredRecords(
  existing: StoredSkillRecord[],
  incoming: StoredSkillRecord[]
): StoredSkillRecord[] {
  const byName = new Map(existing.map((s) => [s.name, s]));
  for (const record of incoming) {
    const current = byName.get(record.name);
    if (current?.origin === "imported" && record.origin === "catalog") continue;
    byName.set(record.name, record);
  }
  return [...byName.values()];
}

export function mergeSkillCatalog(bundled: Skill[], stored: StoredSkillRecord[]): ManagedSkill[] {
  const byName = new Map<string, ManagedSkill>();
  for (const skill of bundled) {
    byName.set(skill.name, {
      ...skill,
      markdown: serializeSkill(skill),
      origin: "bundled",
      updatedAt: "",
      overridesBundled: false
    });
  }
  for (const record of stored) {
    const previous = byName.get(record.name);
    byName.set(record.name, {
      ...record,
      overridesBundled: previous?.origin === "bundled" || Boolean(previous?.overridesBundled)
    });
  }
  return [...byName.values()];
}

export function toStoredRecord(
  skill: Skill,
  opts: { markdown?: string; origin?: SkillOrigin; updatedAt?: string } = {}
): StoredSkillRecord {
  return {
    ...skill,
    markdown: opts.markdown ?? serializeSkill(skill),
    origin: opts.origin ?? "imported",
    updatedAt: opts.updatedAt ?? new Date().toISOString()
  };
}

export async function listImportedSkills(): Promise<StoredSkillRecord[]> {
  if (typeof indexedDB === "undefined") return [];
  const db = await openOpenpluginDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SKILLS_STORE, "readonly");
    const req = tx.objectStore(SKILLS_STORE).getAll();
    req.onsuccess = () => {
      const rows = ((req.result as unknown[]) ?? [])
        .map(normalizeStoredSkill)
        .filter((s): s is StoredSkillRecord => s != null);
      resolve(rows);
    };
    req.onerror = () => reject(req.error);
  });
}

export async function removeImportedSkill(name: string): Promise<void> {
  const db = await openOpenpluginDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(SKILLS_STORE, "readwrite");
    tx.objectStore(SKILLS_STORE).delete(name);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

const DISABLED_KEY = "openplugin.disabledSkills";

export function loadDisabledSkills(): string[] {
  try {
    const raw = localStorage.getItem(DISABLED_KEY);
    return raw ? (JSON.parse(raw) as string[]) : [];
  } catch {
    return [];
  }
}

export function saveDisabledSkills(names: string[]): void {
  localStorage.setItem(DISABLED_KEY, JSON.stringify(names));
}

export async function saveImportedSkill(record: StoredSkillRecord): Promise<void> {
  const db = await openOpenpluginDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(SKILLS_STORE, "readwrite");
    tx.objectStore(SKILLS_STORE).put(record, record.name);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function importSkillFromMarkdown(
  md: string,
  folder: string,
  origin: SkillOrigin = "imported"
): Promise<StoredSkillRecord> {
  const skill = parseSkillMarkdown(md, folder);
  const record = toStoredRecord(skill, { markdown: md, origin });
  await saveImportedSkill(record);
  return record;
}

export async function importCatalog(
  catalogUrl: string,
  existing: StoredSkillRecord[] = []
): Promise<StoredSkillRecord[]> {
  const res = await fetch(catalogUrl);
  if (!res.ok) throw new Error(`Catalog failed (${res.status})`);
  const body = (await res.json()) as { skills?: Array<{ name: string; url: string }> };
  const out: StoredSkillRecord[] = [];
  const importedNames = new Set(existing.filter((s) => s.origin === "imported").map((s) => s.name));
  for (const item of body.skills ?? []) {
    if (item.name && importedNames.has(item.name)) continue;
    try {
      const record = await fetchSkillRecord(item.url, "catalog");
      if (importedNames.has(record.name)) continue;
      await saveImportedSkill(record);
      out.push(record);
    } catch {
      /* skip broken entries */
    }
  }
  return out;
}

export async function importSkillFromUrl(
  url: string,
  origin: SkillOrigin = "imported"
): Promise<StoredSkillRecord> {
  const record = await fetchSkillRecord(url, origin);
  await saveImportedSkill(record);
  return record;
}

async function fetchSkillRecord(url: string, origin: SkillOrigin): Promise<StoredSkillRecord> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Could not fetch skill (${res.status})`);
  const text = await res.text();
  const folder =
    url
      .split("/")
      .filter(Boolean)
      .at(-2)
      ?.replace(/[^a-z0-9-]/g, "") || "imported-skill";
  const skill = parseSkillMarkdown(text, folder);
  return toStoredRecord(skill, { markdown: text, origin });
}

export function downloadSkillMarkdown(name: string, markdown: string): void {
  const blob = new Blob([markdown], { type: "text/markdown" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${name}.md`;
  a.click();
  URL.revokeObjectURL(url);
}
