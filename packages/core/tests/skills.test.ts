import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parseSkillMarkdown, serializeSkill, SkillParseError } from "../src/skills/parse.js";
import { skillRegistryFromDirectory } from "../src/skills/fs.js";
import { SkillRegistry } from "../src/skills/registry.js";
import { skillsForAgent, uniqueSkillName } from "../src/skills/manage.js";

const VALID = `---
name: financial-memo
description: Draft a one-page financial memo from the active Excel selection. Use when the user wants a narrative of variance.
license: Apache-2.0
compatibility: Requires OpenPlugin Word or Excel host
metadata:
  openplugin/hosts: "excel,word"
  openplugin/tools: "excel.getSummary excel.readRange"
  openplugin/version: "1.0"
---

# Financial memo

Write a one-page memo. Read the selection first.
`;

describe("parseSkillMarkdown", () => {
  it("parses frontmatter and body per the Agent Skills spec", () => {
    const skill = parseSkillMarkdown(VALID, "/skills/financial-memo");
    expect(skill.name).toBe("financial-memo");
    expect(skill.description).toContain("financial memo");
    expect(skill.license).toBe("Apache-2.0");
    expect(skill.hosts).toEqual(["excel", "word"]);
    expect(skill.tools).toEqual(["excel.getSummary", "excel.readRange"]);
    expect(skill.body).toContain("Write a one-page memo");
    expect(skill.rootPath).toBe("/skills/financial-memo");
  });

  it("rejects a name that does not match the directory", () => {
    expect(() => parseSkillMarkdown(VALID, "/skills/other-name")).toThrow(SkillParseError);
  });

  it("rejects uppercase names", () => {
    const md = VALID.replace("name: financial-memo", "name: Financial-Memo");
    expect(() => parseSkillMarkdown(md, "/skills/Financial-Memo")).toThrow(/lowercase/);
  });

  it("defaults hosts to all three when metadata is omitted", () => {
    const md = `---
name: selection-rewrite
description: Rewrite the current selection. Use when the user wants a tone or length change.
---
Body
`;
    const skill = parseSkillMarkdown(md, "/skills/selection-rewrite");
    expect(skill.hosts).toEqual(["excel", "word", "powerpoint"]);
    expect(skill.userInvocable).toBe(true);
    expect(skill.disableModelInvocation).toBe(false);
    expect(skill.inject).toBe("on-demand");
  });

  it("parses invocation and inject flags", () => {
    const md = `---
name: house-style
description: Always-on house style for this host.
user-invocable: false
disable-model-invocation: true
metadata:
  openplugin/hosts: "word"
  openplugin/inject: always
---
Use formal tone.
`;
    const skill = parseSkillMarkdown(md, "/skills/house-style");
    expect(skill.userInvocable).toBe(false);
    expect(skill.disableModelInvocation).toBe(true);
    expect(skill.inject).toBe("always");
  });
});

describe("SkillRegistry", () => {
  function writeSkill(root: string, name: string, extra = "") {
    const dir = join(root, name);
    mkdirSync(dir, { recursive: true });
    mkdirSync(join(dir, "references"), { recursive: true });
    writeFileSync(
      join(dir, "SKILL.md"),
      `---
name: ${name}
description: A test skill named ${name} used when testing the registry.
metadata:
  openplugin/hosts: "excel"
---
Instructions for ${name}.
${extra}
`
    );
    writeFileSync(join(dir, "references", "guide.md"), "# Guide\nDetails.\n");
    return dir;
  }

  it("lists only name and description until a skill is loaded", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    writeSkill(root, "board-deck");
    const registry = skillRegistryFromDirectory(root);
    const list = registry.list("excel");
    expect(list.map((s) => s.name).sort()).toEqual(["board-deck", "range-cleanup"]);
    expect(list.every((s) => !("body" in s) || !s.body)).toBe(true);
  });

  it("filters by host", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    const registry = skillRegistryFromDirectory(root);
    expect(registry.list("word")).toEqual([]);
    expect(registry.list("excel")).toHaveLength(1);
  });

  it("loads the body on demand", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    const registry = skillRegistryFromDirectory(root);
    const loaded = registry.load("range-cleanup");
    expect(loaded.body).toContain("Instructions for range-cleanup");
  });

  it("reads a reference file relative to the skill root", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    const registry = skillRegistryFromDirectory(root);
    expect(registry.readRef("range-cleanup", "references/guide.md")).toContain("Details");
  });

  it("rejects path traversal in readRef", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    const registry = skillRegistryFromDirectory(root);
    expect(() => registry.readRef("range-cleanup", "../secret.md")).toThrow(/traversal/i);
  });

  it("splits slash, model, and always-inject catalogs", () => {
    const root = mkdtempSync(join(tmpdir(), "op-skills-"));
    writeSkill(root, "range-cleanup");
    const dir = join(root, "house-style");
    mkdirSync(dir);
    writeFileSync(
      join(dir, "SKILL.md"),
      `---
name: house-style
description: Always injected house style used when writing.
user-invocable: false
metadata:
  openplugin/hosts: "excel"
  openplugin/inject: always
---
Always use thousand separators.
`
    );
    const dir2 = join(root, "manual-only");
    mkdirSync(dir2);
    writeFileSync(
      join(dir2, "SKILL.md"),
      `---
name: manual-only
description: User-triggered cleanup used when the user types the command.
disable-model-invocation: true
metadata:
  openplugin/hosts: "excel"
---
Do the cleanup.
`
    );
    const registry = skillRegistryFromDirectory(root);
    expect(registry.listForSlash("excel").map((s) => s.name).sort()).toEqual([
      "manual-only",
      "range-cleanup"
    ]);
    expect(registry.listForModel("excel").map((s) => s.name).sort()).toEqual([
      "house-style",
      "range-cleanup"
    ]);
    expect(registry.alwaysInject("excel").map((s) => s.name)).toEqual(["house-style"]);
  });
});

describe("serializeSkill", () => {
  it("roundtrips a fully specified skill", () => {
    const skill = parseSkillMarkdown(VALID, "/skills/financial-memo");
    const again = parseSkillMarkdown(serializeSkill(skill), skill.rootPath);
    expect(again.name).toBe(skill.name);
    expect(again.description).toBe(skill.description);
    expect(again.license).toBe(skill.license);
    expect(again.compatibility).toBe(skill.compatibility);
    expect(again.hosts).toEqual(skill.hosts);
    expect(again.tools).toEqual(skill.tools);
    expect(again.body).toBe(skill.body);
    expect(again.metadata["openplugin/version"]).toBe("1.0");
  });

  it("omits default invocation flags and all-host metadata", () => {
    const md = `---
name: selection-rewrite
description: Rewrite the current selection. Use when the user wants a tone or length change.
---
Body
`;
    const serialized = serializeSkill(parseSkillMarkdown(md, "/skills/selection-rewrite"));
    expect(serialized).not.toMatch(/user-invocable/);
    expect(serialized).not.toMatch(/disable-model-invocation/);
    expect(serialized).not.toMatch(/openplugin\/hosts/);
    expect(serialized).not.toMatch(/openplugin\/inject/);
  });

  it("emits non-default flags", () => {
    const md = `---
name: house-style
description: Always-on house style for this host.
user-invocable: false
disable-model-invocation: true
metadata:
  openplugin/hosts: "word"
  openplugin/inject: always
---
Use formal tone.
`;
    const serialized = serializeSkill(parseSkillMarkdown(md, "/skills/house-style"));
    expect(serialized).toMatch(/user-invocable:\s*false/);
    expect(serialized).toMatch(/disable-model-invocation:\s*true/);
    expect(serialized).toMatch(/openplugin\/hosts:.*word/);
    expect(serialized).toMatch(/openplugin\/inject:.*always/);
    const again = parseSkillMarkdown(serialized, "/skills/house-style");
    expect(again.userInvocable).toBe(false);
    expect(again.disableModelInvocation).toBe(true);
    expect(again.inject).toBe("always");
    expect(again.hosts).toEqual(["word"]);
  });
});

describe("uniqueSkillName", () => {
  it("keeps the base name when it is free", () => {
    expect(uniqueSkillName("range-cleanup", ["other"])).toBe("range-cleanup");
  });

  it("appends -copy then -copy-2 when taken", () => {
    expect(uniqueSkillName("range-cleanup", ["range-cleanup"])).toBe("range-cleanup-copy");
    expect(uniqueSkillName("range-cleanup", ["range-cleanup", "range-cleanup-copy"])).toBe(
      "range-cleanup-copy-2"
    );
  });
});

describe("skillsForAgent", () => {
  it("drops disabled skills so they are not listed, loadable, or always-injected", () => {
    const house = parseSkillMarkdown(
      `---
name: house-style
description: Always-on house style for this host.
metadata:
  openplugin/hosts: "excel"
  openplugin/inject: always
---
Always use thousand separators.
`,
      "/skills/house-style"
    );
    const cleanup = parseSkillMarkdown(
      `---
name: range-cleanup
description: A test skill named range-cleanup used when testing the registry.
metadata:
  openplugin/hosts: "excel"
---
Instructions.
`,
      "/skills/range-cleanup"
    );
    const registry = skillsForAgent([house, cleanup], { disabled: ["house-style"] });
    expect(registry.list("excel").map((s) => s.name)).toEqual(["range-cleanup"]);
    expect(registry.alwaysInject("excel")).toEqual([]);
    expect(() => registry.load("house-style")).toThrow(/Unknown skill/);
  });

  it("drops skills outside a tenant allowlist", () => {
    const a = parseSkillMarkdown(
      `---
name: range-cleanup
description: A test skill named range-cleanup used when testing the registry.
---
Body
`,
      "/skills/range-cleanup"
    );
    const b = parseSkillMarkdown(
      `---
name: board-deck
description: A test skill named board-deck used when testing the registry.
---
Body
`,
      "/skills/board-deck"
    );
    const registry = skillsForAgent([a, b], { disabled: [], allowedSkills: ["range-cleanup"] });
    expect(registry.list().map((s) => s.name)).toEqual(["range-cleanup"]);
  });
});

describe("SkillRegistry.fromSkills", () => {
  it("overwrites by name so an imported skill replaces a bundled one", () => {
    const bundled = parseSkillMarkdown(
      `---
name: range-cleanup
description: Bundled cleanup used when testing override.
---
Bundled body.
`,
      "/skills/range-cleanup"
    );
    const imported = parseSkillMarkdown(
      `---
name: range-cleanup
description: Custom cleanup used when the bundled one is overridden.
---
Imported body.
`,
      "/skills/range-cleanup"
    );
    const registry = SkillRegistry.fromSkills([bundled, imported]);
    expect(registry.load("range-cleanup").body).toBe("Imported body.");
  });
});

