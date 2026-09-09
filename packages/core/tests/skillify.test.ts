import { describe, expect, it } from "vitest";
import {
  compactTranscript,
  extractSkillMarkdown,
  skillifyMessages,
  stubSkillMarkdown
} from "../src/skills/skillify.js";

describe("extractSkillMarkdown", () => {
  const doc = `---
name: tidy-headers
description: Use when the user wants consistent Excel headers.
metadata:
  openplugin/hosts: "excel"
---

# Tidy headers

1. Read the selection.
2. Normalize header cells.
`;

  it("returns a raw SKILL.md document", () => {
    expect(extractSkillMarkdown(doc)).toBe(doc.trim());
  });

  it("extracts a fenced markdown block and ignores surrounding prose", () => {
    const wrapped = `Here is the skill:\n\n\`\`\`markdown\n${doc}\n\`\`\`\n\nLet me know if you want edits.`;
    expect(extractSkillMarkdown(wrapped)).toBe(doc.trim());
  });

  it("extracts a generic fence that contains frontmatter", () => {
    const wrapped = `\`\`\`\n${doc}\n\`\`\``;
    expect(extractSkillMarkdown(wrapped)).toBe(doc.trim());
  });

  it("throws when the model output has no SKILL.md document", () => {
    expect(() => extractSkillMarkdown("Sorry, I could not write a skill.")).toThrow(/SKILL\.md/);
  });
});

describe("compactTranscript", () => {
  it("keeps only the last N non-empty user and assistant turns", () => {
    const turns = Array.from({ length: 20 }, (_, i) => ({
      role: (i % 2 === 0 ? "user" : "assistant") as "user" | "assistant",
      text: `turn ${i}`
    }));
    const compact = compactTranscript(turns, 4, 8000);
    expect(compact.map((t) => t.text)).toEqual(["turn 16", "turn 17", "turn 18", "turn 19"]);
  });

  it("drops empty turns and trims from the start to stay under the character cap", () => {
    const compact = compactTranscript(
      [
        { role: "user", text: "   " },
        { role: "user", text: "aaaa" },
        { role: "assistant", text: "bbbb" },
        { role: "user", text: "cccc" }
      ],
      12,
      8
    );
    expect(compact.map((t) => t.text)).toEqual(["cccc"]);
  });
});

describe("skillifyMessages", () => {
  it("asks for a when-to-use description and current-host metadata", () => {
    const messages = skillifyMessages({
      host: "excel",
      transcript: [
        { role: "user", text: "Clean this table" },
        { role: "assistant", text: "I queued a write of cleaned headers." }
      ],
      existingNames: ["selection-rewrite"]
    });
    expect(messages[0]?.role).toBe("system");
    expect(messages[0]?.content).toMatch(/Use when/i);
    expect(messages[0]?.content).not.toMatch(/workflow summary of the skill/i);
    expect(messages[1]?.content).toMatch(/excel/);
    expect(messages[1]?.content).toMatch(/selection-rewrite/);
    expect(messages[1]?.content).toMatch(/Clean this table/);
  });
});

describe("stubSkillMarkdown", () => {
  it("produces parseable markdown for the current host", () => {
    const md = stubSkillMarkdown({
      host: "word",
      name: "from-chat",
      notes: "User asked for a memo."
    });
    expect(md).toMatch(/^---/);
    expect(md).toMatch(/name: from-chat/);
    expect(md).toMatch(/openplugin\/hosts:.*word/);
    expect(md).toMatch(/User asked for a memo/);
  });
});
