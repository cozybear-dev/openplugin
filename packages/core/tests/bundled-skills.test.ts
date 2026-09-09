import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
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
});
