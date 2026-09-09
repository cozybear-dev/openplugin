export const OPENPLUGIN_DB = "openplugin";
export const OPENPLUGIN_DB_VERSION = 2;
export const SKILLS_STORE = "skills";
export const CONVERSATIONS_STORE = "conversations";

export function upgradeOpenpluginDb(db: IDBDatabase): void {
  if (!db.objectStoreNames.contains(SKILLS_STORE)) db.createObjectStore(SKILLS_STORE);
  if (!db.objectStoreNames.contains(CONVERSATIONS_STORE)) {
    db.createObjectStore(CONVERSATIONS_STORE, { keyPath: "id" });
  }
}

export function openOpenpluginDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(OPENPLUGIN_DB, OPENPLUGIN_DB_VERSION);
    req.onupgradeneeded = () => upgradeOpenpluginDb(req.result);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
