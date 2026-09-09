import type { Skill } from "./parse.js";
import { SkillRegistry } from "./registry.js";

export function uniqueSkillName(base: string, taken: string[]): string {
  if (!taken.includes(base)) return base;
  const copy = `${base}-copy`;
  if (!taken.includes(copy)) return copy;
  let n = 2;
  while (taken.includes(`${base}-copy-${n}`)) n++;
  return `${base}-copy-${n}`;
}

export function skillsForAgent(
  skills: Skill[],
  opts: { disabled: string[]; allowedSkills?: string[] | "*" }
): SkillRegistry {
  const allowed = opts.allowedSkills;
  return SkillRegistry.fromSkills(
    skills.filter((skill) => {
      if (opts.disabled.includes(skill.name)) return false;
      if (allowed && allowed !== "*" && !allowed.includes(skill.name)) return false;
      return true;
    })
  );
}
