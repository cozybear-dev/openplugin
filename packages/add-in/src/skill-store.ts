import { parseSkillMarkdown, type Skill } from "@openplugin/core";

const DB = "openplugin";
const STORE = "skills";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function listImportedSkills(): Promise<Skill[]> {
  if (typeof indexedDB === "undefined") return [];
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const req = tx.objectStore(STORE).getAll();
    req.onsuccess = () => resolve((req.result as Skill[]) ?? []);
    req.onerror = () => reject(req.error);
  });
}

export async function removeImportedSkill(name: string): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(name);
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

export async function saveImportedSkill(skill: Skill): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(skill, skill.name);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function importSkillFromMarkdown(md: string, folder: string): Promise<Skill> {
  const skill = parseSkillMarkdown(md, folder);
  await saveImportedSkill(skill);
  return skill;
}

export async function importCatalog(catalogUrl: string): Promise<Skill[]> {
  const res = await fetch(catalogUrl);
  if (!res.ok) throw new Error(`Catalog failed (${res.status})`);
  const body = (await res.json()) as { skills?: Array<{ name: string; url: string }> };
  const out: Skill[] = [];
  for (const item of body.skills ?? []) {
    try {
      out.push(await importSkillFromUrl(item.url));
    } catch {
      /* skip broken entries */
    }
  }
  return out;
}

export async function importSkillFromUrl(url: string): Promise<Skill> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Could not fetch skill (${res.status})`);
  const text = await res.text();
  const folder =
    url
      .split("/")
      .filter(Boolean)
      .at(-2)
      ?.replace(/[^a-z0-9-]/g, "") || "imported-skill";
  return importSkillFromMarkdown(text, folder);
}
