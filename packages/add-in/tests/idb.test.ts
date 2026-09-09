import { describe, expect, it } from "vitest";
import { CONVERSATIONS_STORE, SKILLS_STORE, upgradeOpenpluginDb } from "../src/idb";

describe("upgradeOpenpluginDb", () => {
  it("creates both stores when names are empty", () => {
    const created: Array<{ name: string; opts?: IDBObjectStoreParameters }> = [];
    const db = {
      objectStoreNames: { contains: () => false },
      createObjectStore: (name: string, opts?: IDBObjectStoreParameters) => {
        created.push({ name, opts });
      }
    };
    upgradeOpenpluginDb(db as unknown as IDBDatabase);
    expect(created).toEqual([
      { name: SKILLS_STORE, opts: undefined },
      { name: CONVERSATIONS_STORE, opts: { keyPath: "id" } }
    ]);
  });

  it("does not recreate an existing skills store", () => {
    const created: string[] = [];
    const db = {
      objectStoreNames: { contains: (n: string) => n === SKILLS_STORE },
      createObjectStore: (name: string) => created.push(name)
    };
    upgradeOpenpluginDb(db as unknown as IDBDatabase);
    expect(created).toEqual([CONVERSATIONS_STORE]);
  });
});
