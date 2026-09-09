import { parse as parseYaml } from "yaml";
import type { HostKind } from "../llm/types.js";

const HOSTS: HostKind[] = ["excel", "word", "powerpoint"];

export type SkillMeta = {
  name: string;
  description: string;
  license?: string;
  compatibility?: string;
  metadata: Record<string, string>;
  hosts: HostKind[];
  tools?: string[];
  rootPath: string;
};

export type Skill = SkillMeta & {
  body: string;
  files?: Record<string, string>;
};

export class SkillParseError extends Error {
  readonly name = "SkillParseError";
}

const NAME_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;

export function parseSkillMarkdown(markdown: string, rootPath: string): Skill {
  const match = markdown.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!match) throw new SkillParseError("SKILL.md must start with YAML frontmatter delimited by ---.");
  const raw = parseYaml(match[1]) as Record<string, unknown>;
  if (!raw || typeof raw !== "object") throw new SkillParseError("Frontmatter must be a YAML mapping.");

  const name = String(raw.name ?? "").trim();
  const description = String(raw.description ?? "").trim();
  if (!name) throw new SkillParseError("Frontmatter is missing name.");
  if (!description) throw new SkillParseError("Frontmatter is missing description.");
  if (name.length > 64) throw new SkillParseError("name must be at most 64 characters.");
  if (description.length > 1024) throw new SkillParseError("description must be at most 1024 characters.");
  if (!NAME_RE.test(name)) {
    throw new SkillParseError("name must be lowercase letters, numbers, and single hyphens.");
  }

  const dir = rootPath.replace(/[\\/]+$/, "").split(/[\\/]/).pop() ?? "";
  if (dir !== name) {
    throw new SkillParseError(`Skill name "${name}" must match the directory name "${dir}".`);
  }

  const metadata = stringifyMeta(raw.metadata);
  const hosts = parseHosts(metadata["openplugin/hosts"]);
  const tools = metadata["openplugin/tools"]
    ? metadata["openplugin/tools"].split(/\s+/).filter(Boolean)
    : undefined;

  return {
    name,
    description,
    license: raw.license ? String(raw.license) : undefined,
    compatibility: raw.compatibility ? String(raw.compatibility) : undefined,
    metadata,
    hosts,
    tools,
    rootPath,
    body: match[2].trim()
  };
}

function stringifyMeta(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object") return {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    if (v == null) continue;
    out[k] = typeof v === "string" ? v : String(v);
  }
  return out;
}

function parseHosts(raw?: string): HostKind[] {
  if (!raw) return [...HOSTS];
  const parts = raw.split(/[,\s]+/).filter(Boolean) as HostKind[];
  const hosts = parts.filter((h): h is HostKind => HOSTS.includes(h));
  return hosts.length ? hosts : [...HOSTS];
}
