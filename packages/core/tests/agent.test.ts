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
    const payload = turns[i++] ?? { choices: [{ delta: { content: "done" } }] };
    const sse =
      typeof payload === "string"
        ? payload
        : `data: ${JSON.stringify({ choices: [{ delta: payload }] })}\n\ndata: [DONE]\n\n`;
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
});
