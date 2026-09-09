import { parseSkillMarkdown, serializeSkill, type Skill } from "@openplugin/core";
import { describe, expect, it } from "vitest";
import {
  mergeSkillCatalog,
  mergeStoredRecords,
  normalizeStoredSkill,
  type StoredSkillRecord
} from "../src/skill-store";

function skill(name: string, extra = ""): Skill {
  return parseSkillMarkdown(
    `---
name: ${name}
description: A test skill named ${name} used when testing the store.
metadata:
  openplugin/hosts: "excel"
---
Body for ${name}.
${extra}
`,
    name
  );
}

function record(name: string, origin: StoredSkillRecord["origin"], body = `Body for ${name}.`): StoredSkillRecord {
  const s = skill(name);
  s.body = body;
  return {
    ...s,
    body,
    markdown: serializeSkill({ ...s, body }),
    origin,
    updatedAt: "2026-01-01T00:00:00.000Z"
  };
}

describe("normalizeStoredSkill", () => {
  it("fills markdown and origin for a legacy Skill object", () => {
    const legacy = skill("range-cleanup");
    const normalized = normalizeStoredSkill(legacy);
    expect(normalized?.name).toBe("range-cleanup");
    expect(normalized?.origin).toBe("imported");
    expect(normalized?.markdown).toContain("name: range-cleanup");
    expect(normalized?.markdown).toContain("Body for range-cleanup");
  });

  it("keeps stored markdown and catalog origin", () => {
    const raw = record("house-style", "catalog");
    const normalized = normalizeStoredSkill(raw);
    expect(normalized?.origin).toBe("catalog");
    expect(normalized?.markdown).toBe(raw.markdown);
  });

  it("returns null for junk", () => {
    expect(normalizeStoredSkill(null)).toBeNull();
    expect(normalizeStoredSkill({ foo: 1 })).toBeNull();
  });
});

describe("mergeStoredRecords", () => {
  it("does not let a catalog skill overwrite a user import of the same name", () => {
    const imported = record("range-cleanup", "imported", "User body.");
    const catalog = record("range-cleanup", "catalog", "Catalog body.");
    const merged = mergeStoredRecords([imported], [catalog]);
    expect(merged).toHaveLength(1);
    expect(merged[0]?.origin).toBe("imported");
    expect(merged[0]?.body).toBe("User body.");
  });

  it("adds new catalog skills and replaces previous catalog copies", () => {
    const previous = record("house-style", "catalog", "Old.");
    const next = record("house-style", "catalog", "New.");
    const extra = record("board-deck", "catalog");
    const merged = mergeStoredRecords([previous], [next, extra]);
    expect(merged.map((s) => s.name).sort()).toEqual(["board-deck", "house-style"]);
    expect(merged.find((s) => s.name === "house-style")?.body).toBe("New.");
  });
});

describe("mergeSkillCatalog", () => {
  it("shows one row when an import overrides a bundled skill", () => {
    const bundled = skill("range-cleanup");
    bundled.body = "Bundled body.";
    const imported = record("range-cleanup", "imported", "User body.");
    const rows = mergeSkillCatalog([bundled], [imported]);
    expect(rows).toHaveLength(1);
    expect(rows[0]?.origin).toBe("imported");
    expect(rows[0]?.overridesBundled).toBe(true);
    expect(rows[0]?.body).toBe("User body.");
  });

  it("keeps bundled skills that have no stored override", () => {
    const rows = mergeSkillCatalog([skill("selection-rewrite")], [record("range-cleanup", "imported")]);
    expect(rows.map((s) => s.name).sort()).toEqual(["range-cleanup", "selection-rewrite"]);
    expect(rows.find((s) => s.name === "selection-rewrite")?.origin).toBe("bundled");
  });
});
