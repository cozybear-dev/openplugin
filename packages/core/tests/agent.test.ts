import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { runAgent } from "../src/agent/loop.js";
import { FakeExcelHost } from "../src/hosts/fake-excel.js";
import { skillRegistryFromDirectory } from "../src/skills/fs.js";
import type { ProviderConfig } from "../src/llm/types.js";

const config: ProviderConfig = {
  baseUrl: "https://api.example.com/v1",
  apiKey: "sk-test",
  model: "scripted"
};

function scriptedFetch(turns: unknown[]): typeof fetch {
  let i = 0;
  return async () => {
    const payload = turns[i++] ?? { content: "done" };
    if (typeof payload === "string") {
      return new Response(payload, { headers: { "Content-Type": "text/event-stream" } });
    }
    const choice = payload as { content?: string; tool_calls?: unknown; finish_reason?: string };
    const delta = choice.tool_calls
      ? { tool_calls: choice.tool_calls }
      : { content: choice.content ?? "done" };
    const sse = `data: ${JSON.stringify({
      choices: [{ delta, finish_reason: choice.finish_reason ?? null }]
    })}\n\ndata: [DONE]\n\n`;
    return new Response(sse, { headers: { "Content-Type": "text/event-stream" } });
  };
}

function toolDelta(name: string, args: unknown, id = "call_1") {
  return {
    tool_calls: [
      {
        index: 0,
        id,
        type: "function",
        function: { name, arguments: JSON.stringify(args) }
      }
    ]
  };
}

describe("runAgent", () => {
  it("rejects web tools when web search is disabled", async () => {
    const host = new FakeExcelHost();
    const skills = skillRegistryFromDirectory(mkdtempSync(join(tmpdir(), "op-web-disabled-")));
    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "search",
      fetchImpl: scriptedFetch([toolDelta("web.search", { query: "secret" }), { content: "done" }])
    });

    expect(result.messages).toContainEqual(
      expect.objectContaining({
        role: "tool",
        content: JSON.stringify({ error: "Web search is disabled." })
      })
    );
  });

  it("enforces the skill allowlist for always-injected skills", async () => {
    const root = mkdtempSync(join(tmpdir(), "op-skill-policy-"));
    const dir = join(root, "restricted");
    mkdirSync(dir);
    writeFileSync(
      join(dir, "SKILL.md"),
      `---
name: restricted
description: Restricted skill
metadata:
  openplugin/hosts: "excel"
  openplugin/inject: "always"
---
Do not expose this.
`
    );

    await expect(
      runAgent({
        config,
        host: new FakeExcelHost(),
        skills: skillRegistryFromDirectory(root),
        userMessage: "hello",
        policy: { denyExecuteJs: true, allowedSkills: [] },
        fetchImpl: scriptedFetch([{ content: "done" }])
      })
    ).rejects.toThrow("Skill is not allowed by policy.");
  });

  it("loads a skill then records an excel write in the changeset without applying it", async () => {
    const root = mkdtempSync(join(tmpdir(), "op-agent-"));
    const dir = join(root, "range-cleanup");
    mkdirSync(dir);
    writeFileSync(
      join(dir, "SKILL.md"),
      `---
name: range-cleanup
description: Clean a messy Excel range. Use when headers or types are wrong.
metadata:
  openplugin/hosts: "excel"
---
Turn the selection into a table.
`
    );

    const host = new FakeExcelHost();
    host.selection = { sheet: "Sheet1", address: "A1:B2", values: [["a", "b"], [1, 2]] };
    const skills = skillRegistryFromDirectory(root);

    const fetchImpl = scriptedFetch([
      toolDelta("skills.load", { name: "range-cleanup" }),
      toolDelta(
        "excel.writeRange",
        { sheet: "Sheet1", address: "A1:B2", values: [["A", "B"], [1, 2]] },
        "call_2"
      ),
      { content: "Ready to apply the cleaned range." }
    ]);

    const events: string[] = [];
    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "clean this range",
      fetchImpl,
      onEvent: (e) => events.push(e.type)
    });

    expect(result.finalText).toContain("Ready to apply");
    expect(result.changeset.changes).toHaveLength(1);
    expect(result.loadedSkills).toContain("range-cleanup");
    expect(host.sheets.Sheet1.values[0][0]).not.toBe("A");
    expect(events).toContain("tool");
    expect(events).toContain("text");
  });

  it("stops at maxSteps", async () => {
    const host = new FakeExcelHost();
    const skills = skillRegistryFromDirectory(mkdtempSync(join(tmpdir(), "op-empty-")));
    const fetchImpl = scriptedFetch(
      Array.from({ length: 12 }, (_, i) =>
        toolDelta("excel.getSummary", {}, `call_${i}`)
      )
    );

    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "loop",
      maxSteps: 3,
      fetchImpl
    });

    expect(result.steps).toBe(3);
    expect(result.stopReason).toBe("max_steps");
  });

  it("continues when finish_reason is length and concatenates text", async () => {
    const host = new FakeExcelHost();
    const skills = skillRegistryFromDirectory(mkdtempSync(join(tmpdir(), "op-len-")));
    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "explain this sheet",
      fetchImpl: scriptedFetch([
        { content: "Section one.\n\n", finish_reason: "length" },
        { content: "Section two." }
      ])
    });
    expect(result.finalText).toBe("Section one.\n\nSection two.");
    expect(result.steps).toBe(2);
    expect(result.stopReason).toBe("final");
  });

  it("nudges once when the first reply is text with no tools", async () => {
    const host = new FakeExcelHost();
    const skills = skillRegistryFromDirectory(mkdtempSync(join(tmpdir(), "op-nudge-")));
    const seen: string[][] = [];
    let i = 0;
    const turns = [
      { content: "I will clean the range." },
      toolDelta("excel.getSummary", {}),
      { content: "Ready." }
    ];
    const fetchImpl: typeof fetch = async (_url, init) => {
      const body = JSON.parse(String(init?.body)) as { messages: Array<{ role: string; content?: string }> };
      seen.push(body.messages.map((m) => `${m.role}:${typeof m.content === "string" ? m.content.slice(0, 80) : ""}`));
      const payload = turns[i++] ?? { content: "done" };
      return scriptedFetch([payload])(_url, init);
    };
    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "clean this range",
      fetchImpl
    });
    expect(seen[1]?.some((m) => m.includes("Continue. If document work remains"))).toBe(true);
    expect(result.finalText).toContain("Ready");
    expect(result.stopReason).toBe("final");
  });

  it("does not nudge after tools have already run", async () => {
    const host = new FakeExcelHost();
    const skills = skillRegistryFromDirectory(mkdtempSync(join(tmpdir(), "op-nonudge-")));
    const seen: string[] = [];
    let i = 0;
    const turns = [toolDelta("excel.getSummary", {}), { content: "Here is the summary." }];
    const fetchImpl: typeof fetch = async (_url, init) => {
      const body = JSON.parse(String(init?.body)) as { messages: Array<{ role: string; content?: string }> };
      for (const m of body.messages) {
        if (typeof m.content === "string" && m.content.startsWith("Continue.")) seen.push(m.content);
      }
      const payload = turns[i++] ?? { content: "done" };
      return scriptedFetch([payload])(_url, init);
    };
    const result = await runAgent({
      config,
      host,
      skills,
      userMessage: "summarize",
      fetchImpl
    });
    expect(seen).toEqual([]);
    expect(result.finalText).toContain("Here is the summary.");
    expect(result.steps).toBe(2);
  });
});
