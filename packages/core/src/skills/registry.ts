import type { HostKind } from "../llm/types.js";
import type { Skill, SkillMeta } from "./parse.js";

export type SkillCatalogEntry = Pick<
  SkillMeta,
  "name" | "description" | "hosts" | "tools" | "userInvocable" | "disableModelInvocation" | "inject"
>;

function assertSafeRelative(path: string): string {
  const normalized = path.replace(/\\/g, "/").replace(/^\/+/, "");
  const parts = normalized.split("/");
  if (parts.some((p) => p === ".." || p === "")) {
    throw new Error("Path traversal is not allowed in skill references.");
  }
  return normalized;
}

export class SkillRegistry {
  private readonly skills = new Map<string, Skill>();

  static fromSkills(skills: Skill[]): SkillRegistry {
    const registry = new SkillRegistry();
    for (const skill of skills) registry.add(skill);
    return registry;
  }

  add(skill: Skill): void {
    this.skills.set(skill.name, skill);
  }

  list(host?: HostKind): SkillCatalogEntry[] {
    return [...this.skills.values()]
      .filter((s) => !host || s.hosts.includes(host))
      .map(
        ({
          name,
          description,
          hosts,
          tools,
          userInvocable,
          disableModelInvocation,
          inject
        }) => ({ name, description, hosts, tools, userInvocable, disableModelInvocation, inject })
      );
  }

  listForSlash(host?: HostKind): SkillCatalogEntry[] {
    return this.list(host).filter((s) => s.userInvocable);
  }

  listForModel(host?: HostKind): SkillCatalogEntry[] {
    return this.list(host).filter((s) => !s.disableModelInvocation);
  }

  alwaysInject(host?: HostKind): Skill[] {
    return [...this.skills.values()].filter(
      (s) => s.inject === "always" && (!host || s.hosts.includes(host))
    );
  }

  load(name: string): Skill {
    const skill = this.skills.get(name);
    if (!skill) throw new Error(`Unknown skill: ${name}`);
    return skill;
  }

  readRef(name: string, relativePath: string): string {
    const skill = this.load(name);
    const safe = assertSafeRelative(relativePath);
    const content = skill.files?.[safe];
    if (content == null) throw new Error(`Skill file not found: ${relativePath}`);
    return content;
  }
}
