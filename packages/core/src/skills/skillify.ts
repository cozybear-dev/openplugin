import type { ChatMessage, HostKind } from "../llm/types.js";
import { serializeSkill } from "./parse.js";

export type SkillifyTurn = { role: "user" | "assistant"; text: string };

export const SKILLIFY_MAX_TURNS = 12;
export const SKILLIFY_MAX_CHARS = 8000;

const SKILLIFY_SYSTEM = `You write a reusable OpenPlugin Agent Skill from a chat transcript.

Return only a SKILL.md document. YAML frontmatter must come first, delimited by ---. Required fields: name (lowercase letters, numbers, single hyphens, max 64 chars) and description (max 1024 chars).

description must start with "Use when" and state triggering conditions only. Do not summarize the skill's workflow in description.

Body: numbered reusable steps. No recap of the specific document or numbers from this chat. No narrative of how the chat went.

Frontmatter metadata:
- openplugin/hosts: comma-separated subset of excel,word,powerpoint
- openplugin/tools: optional space-separated tool names
- openplugin/inject: omit unless the skill must always be injected (then "always")

Optional flags: user-invocable (default true), disable-model-invocation (default false).`;

export function compactTranscript(
  turns: SkillifyTurn[],
  maxTurns = SKILLIFY_MAX_TURNS,
  maxChars = SKILLIFY_MAX_CHARS
): SkillifyTurn[] {
  let sliced = turns.filter((t) => t.text.trim()).slice(-maxTurns);
  const chars = (list: SkillifyTurn[]) => list.reduce((n, t) => n + t.text.length, 0);
  while (sliced.length > 1 && chars(sliced) >= maxChars) sliced = sliced.slice(1);
  if (sliced.length === 1 && sliced[0] && sliced[0].text.length > maxChars) {
    sliced = [{ ...sliced[0], text: sliced[0].text.slice(-maxChars) }];
  }
  return sliced;
}

export function extractSkillMarkdown(text: string): string {
  const fenced = [...text.matchAll(/```(?:markdown|md)?\s*\n([\s\S]*?)```/gi)]
    .map((m) => m[1].trim())
    .find((block) => /^---\r?\n/.test(block));
  if (fenced) return fenced.trim();

  const start = text.search(/^---[ \t]*\r?\n/m);
  if (start >= 0) {
    const slice = text.slice(start).trim();
    if (/^---[ \t]*\r?\n[\s\S]*?\r?\n---/.test(slice)) return slice;
  }
  throw new Error("Model output did not contain a SKILL.md document.");
}

export function skillifyMessages(opts: {
  host: HostKind;
  transcript: SkillifyTurn[];
  existingNames: string[];
}): ChatMessage[] {
  const compact = compactTranscript(opts.transcript);
  const transcript = compact.map((t) => `${t.role}: ${t.text}`).join("\n\n");
  const taken = opts.existingNames.length ? opts.existingNames.join(", ") : "(none)";
  return [
    { role: "system", content: SKILLIFY_SYSTEM },
    {
      role: "user",
      content: `Host: ${opts.host}
Existing skill names (pick a different name): ${taken}

Transcript:
${transcript || "(empty)"}`
    }
  ];
}

export function stubSkillMarkdown(opts: { host: HostKind; name?: string; notes?: string }): string {
  const name = opts.name ?? "from-chat";
  const notes = opts.notes?.trim() ? `\n\nNotes from the chat:\n${opts.notes.trim()}` : "";
  return serializeSkill({
    name,
    description: "Use when the user wants to repeat the workflow captured in this skill.",
    metadata: { "openplugin/hosts": opts.host },
    hosts: [opts.host],
    rootPath: name,
    userInvocable: true,
    disableModelInvocation: false,
    inject: "on-demand",
    body: `# ${name}\n\n1. Restate the goal.\n2. Inspect the current selection.\n3. Apply the same steps that worked in the source chat.${notes}`
  });
}
