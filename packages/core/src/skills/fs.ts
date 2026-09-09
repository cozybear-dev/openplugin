import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative } from "node:path";
import { parseSkillMarkdown } from "./parse.js";
import { SkillRegistry } from "./registry.js";

function collectFiles(dir: string, base = dir, acc: Record<string, string> = {}): Record<string, string> {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const rel = relative(base, full).replace(/\\/g, "/");
    if (statSync(full).isDirectory()) collectFiles(full, base, acc);
    else if (entry !== "SKILL.md") acc[rel] = readFileSync(full, "utf8");
  }
  return acc;
}

export function skillRegistryFromDirectory(root: string): SkillRegistry {
  const registry = new SkillRegistry();
  if (!existsSync(root)) return registry;
  for (const entry of readdirSync(root)) {
    const dir = join(root, entry);
    if (!statSync(dir).isDirectory()) continue;
    const skillFile = join(dir, "SKILL.md");
    if (!existsSync(skillFile)) continue;
    const skill = parseSkillMarkdown(readFileSync(skillFile, "utf8"), dir);
    skill.files = collectFiles(dir);
    registry.add(skill);
  }
  return registry;
}
