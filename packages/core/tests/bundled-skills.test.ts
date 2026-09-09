import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { parseSkillMarkdown, serializeSkill } from "../src/skills/parse.js";
import { skillRegistryFromDirectory } from "../src/skills/fs.js";

const repoSkills = join(dirname(fileURLToPath(import.meta.url)), "../../../skills");

describe("bundled skills", () => {
  it("parses the four shipped skills", () => {
    const registry = skillRegistryFromDirectory(repoSkills);
    const names = registry.list().map((s) => s.name).sort();
    expect(names).toEqual([
      "excel-range-cleanup",
      "ppt-outline-to-slides",
      "selection-rewrite",
      "word-memo-from-sheet"
    ]);
    expect(registry.load("selection-rewrite").body).toContain("Read only the current selection");
  });

  it("roundtrips each shipped skill through serializeSkill", () => {
    const registry = skillRegistryFromDirectory(repoSkills);
    for (const entry of registry.list()) {
      const skill = registry.load(entry.name);
      const again = parseSkillMarkdown(serializeSkill(skill), skill.rootPath);
      expect(again.name).toBe(skill.name);
      expect(again.description).toBe(skill.description);
      expect(again.hosts).toEqual(skill.hosts);
      expect(again.body).toBe(skill.body);
    }
  });
});
