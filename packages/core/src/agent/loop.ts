import { compileSnapshot, snapshotToPrompt, type DocumentSnapshot } from "../context/compiler.js";
import type { HostAdapter } from "../hosts/types.js";
import { chatCompletions } from "../llm/client.js";
import type { ChatMessage, ProviderConfig, ToolCall } from "../llm/types.js";
import { assertPolicy, OPEN_POLICY, type Policy } from "../policy.js";
import {
  nativeSearchTool,
  resolveSearchBackend,
  runWebFetch,
  runWebSearch,
  type SearchBackend
} from "../search/index.js";
import type { SkillRegistry } from "../skills/registry.js";
import { Changeset } from "../tools/changeset.js";
import { describeActivity, type ActivityKind } from "../tools/describe.js";
import { executeHostTool, listToolDefinitions } from "../tools/registry.js";

export const SYSTEM_PROMPT = `You are OpenPlugin, an agent that edits Microsoft Office documents through typed tools.

Rules:
- Prefer tools over guessing document content.
- Load a skill with skills.load when one matches the task.
- Mutations go into a changeset; the user will apply them. Do not claim they are already applied.
- Never dump or request the entire document. Use summaries and ranged reads.
- Keep tool arguments compact.
- If a tool errors, choose a different tool or ask a concise question.`;

export type AgentEvent =
  | { type: "text"; delta: string }
  | { type: "tool"; name: string; args: unknown; result?: unknown }
  | {
      type: "activity";
      phase: "start" | "done" | "error";
      activity: ActivityKind;
      name: string;
      label: string;
      detail?: string;
    }
  | { type: "skill"; name: string }
  | { type: "error"; message: string };

export type WebSearchOptions = {
  enabled: boolean;
  backend: SearchBackend;
  exaApiKey?: string;
  customUrl?: string;
  extraHeaders?: Record<string, unknown> | Record<string, string>;
  proxy?: import("../search/index.js").SearchProxy;
};

export type AgentResult = {
  finalText: string;
  changeset: Changeset;
  messages: ChatMessage[];
  loadedSkills: string[];
  steps: number;
  stopReason: "final" | "max_steps" | "abort";
  snapshot: DocumentSnapshot;
};

export async function runAgent(opts: {
  config: ProviderConfig;
  host: HostAdapter;
  skills: SkillRegistry;
  userMessage: string;
  history?: ChatMessage[];
  maxSteps?: number;
  tokenBudget?: number;
  executeJsEnabled?: boolean;
  policy?: Policy;
  fetchImpl?: typeof fetch;
  onEvent?: (event: AgentEvent) => void;
  signal?: AbortSignal;
  previousSnapshot?: DocumentSnapshot;
  webSearch?: WebSearchOptions;
}): Promise<AgentResult> {
  const maxSteps = opts.maxSteps ?? 8;
  const policy = opts.policy ?? OPEN_POLICY;
  assertPolicy(policy, {
    baseUrl: opts.config.baseUrl,
    model: opts.config.model,
    executeJs: opts.executeJsEnabled
  });
  const changeset = new Changeset();
  const loadedSkills: string[] = [];
  const snapshot = await compileSnapshot(opts.host, {
    tokenBudget: opts.tokenBudget,
    previous: opts.previousSnapshot
  });

  const catalog = opts.skills
    .listForModel(opts.host.kind)
    .map((s) => `- ${s.name}: ${s.description}`)
    .join("\n");
  const injected = opts.skills.alwaysInject(opts.host.kind);
  for (const skill of injected) {
    assertPolicy(policy, {
      baseUrl: opts.config.baseUrl,
      model: opts.config.model,
      skill: skill.name
    });
    if (!loadedSkills.includes(skill.name)) loadedSkills.push(skill.name);
  }
  const injectedBlock = injected.length
    ? `\n\nAlways-on skills:\n${injected.map((s) => `## ${s.name}\n${s.body}`).join("\n\n")}`
    : "";

  const messages: ChatMessage[] = [
    {
      role: "system",
      content: `${SYSTEM_PROMPT}\n\nInstalled skills:\n${catalog || "(none)"}${injectedBlock}`
    },
    ...(opts.history ?? []),
    {
      role: "user",
      content: `${opts.userMessage}\n\nDocument snapshot:\n${snapshotToPrompt(snapshot)}`
    }
  ];

  const searchResolved =
    opts.webSearch?.enabled === true
      ? resolveSearchBackend({
          backend: opts.webSearch.backend,
          baseUrl: opts.config.baseUrl,
          exaApiKey: opts.webSearch.exaApiKey
        })
      : undefined;
  const nativeTool = searchResolved === "native" ? nativeSearchTool(opts.config.baseUrl) : undefined;
  const tools = listToolDefinitions(opts.host.kind, {
    executeJsEnabled: opts.executeJsEnabled,
    webSearch: nativeTool ? nativeTool : searchResolved ? "function" : false
  });
  let finalText = "";
  let steps = 0;

  for (; steps < maxSteps; steps++) {
    if (opts.signal?.aborted) {
      return { finalText, changeset, messages, loadedSkills, steps, stopReason: "abort", snapshot };
    }

    const result = await chatCompletions({
      config: opts.config,
      messages,
      tools,
      fetchImpl: opts.fetchImpl,
      signal: opts.signal,
      onEvent: (e) => {
        if (e.type === "text") opts.onEvent?.({ type: "text", delta: e.delta });
      }
    });

    messages.push(result.message);
    const calls = result.message.role === "assistant" ? result.message.tool_calls : undefined;
    if (!calls || calls.length === 0) {
      finalText = (result.message.role === "assistant" && result.message.content) || "";
      return { finalText, changeset, messages, loadedSkills, steps: steps + 1, stopReason: "final", snapshot };
    }

    for (const call of calls) {
      const toolResult = await dispatchTool(call, {
        host: opts.host,
        skills: opts.skills,
        changeset,
        loadedSkills,
        onEvent: opts.onEvent,
        policy,
        baseUrl: opts.config.baseUrl,
        model: opts.config.model,
        fetchImpl: opts.fetchImpl,
        webSearch: opts.webSearch,
        searchResolved
      });
      messages.push({
        role: "tool",
        tool_call_id: call.id,
        content: typeof toolResult === "string" ? toolResult : JSON.stringify(toolResult)
      });
    }
  }

  return { finalText, changeset, messages, loadedSkills, steps, stopReason: "max_steps", snapshot };
}

async function dispatchTool(
  call: ToolCall,
  ctx: {
    host: HostAdapter;
    skills: SkillRegistry;
    changeset: Changeset;
    loadedSkills: string[];
    onEvent?: (event: AgentEvent) => void;
    policy: Policy;
    baseUrl: string;
    model: string;
    fetchImpl?: typeof fetch;
    webSearch?: WebSearchOptions;
    searchResolved?: ReturnType<typeof resolveSearchBackend>;
  }
): Promise<unknown> {
  let args: Record<string, unknown> = {};
  try {
    args = call.function.arguments ? (JSON.parse(call.function.arguments) as Record<string, unknown>) : {};
  } catch {
    args = {};
  }
  const name = call.function.name;
  const described = describeActivity(name, args);
  ctx.onEvent?.({ type: "tool", name, args });
  ctx.onEvent?.({
    type: "activity",
    phase: "start",
    activity: described.activity,
    name,
    label: described.label,
    detail: described.detail
  });

  try {
    if (name === "skills.list") return emitDone(ctx, described, name, ctx.skills.listForModel(ctx.host.kind));
    if (name === "skills.load") {
      assertPolicy(ctx.policy, {
        baseUrl: ctx.baseUrl,
        model: ctx.model,
        skill: String(args.name)
      });
      const skill = ctx.skills.load(String(args.name));
      if (skill.inject === "always") {
        const already = { name: skill.name, body: skill.body, tools: skill.tools, alreadyInContext: true };
        return emitDone(ctx, described, name, already);
      }
      if (!ctx.loadedSkills.includes(skill.name)) ctx.loadedSkills.push(skill.name);
      ctx.onEvent?.({ type: "skill", name: skill.name });
      return emitDone(ctx, described, name, { name: skill.name, body: skill.body, tools: skill.tools });
    }
    if (name === "skills.readRef") {
      return emitDone(ctx, described, name, ctx.skills.readRef(String(args.name), String(args.path)));
    }
    if (name === "changeset.preview") return emitDone(ctx, described, name, { preview: ctx.changeset.preview() });
    if (name === "web.search" || name === "web.fetch") {
      if (ctx.webSearch?.enabled !== true) {
        throw new Error("Web search is disabled.");
      }
      const backend = ctx.searchResolved ?? "duckduckgo";
      const result =
        name === "web.search"
          ? await runWebSearch({
              backend: backend === "native" ? "duckduckgo" : backend,
              query: String(args.query ?? ""),
              apiKey: ctx.webSearch?.exaApiKey,
              customUrl: ctx.webSearch?.customUrl,
              extraHeaders: ctx.webSearch?.extraHeaders as Record<string, string> | undefined,
              fetchImpl: ctx.fetchImpl,
              proxy: ctx.webSearch?.proxy
            })
          : await runWebFetch({
              backend: backend === "native" ? "duckduckgo" : backend,
              url: String(args.url ?? ""),
              apiKey: ctx.webSearch?.exaApiKey,
              fetchImpl: ctx.fetchImpl,
              proxy: ctx.webSearch?.proxy
            });
      if (!result.ok) {
        ctx.onEvent?.({
          type: "activity",
          phase: "error",
          activity: described.activity,
          name,
          label: described.label,
          detail: result.error
        });
        return result;
      }
      return emitDone(ctx, described, name, result);
    }
    const result = await executeHostTool(ctx.host, name, args, ctx.changeset);
    ctx.onEvent?.({ type: "tool", name, args, result });
    return emitDone(ctx, described, name, result);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    ctx.onEvent?.({ type: "error", message });
    ctx.onEvent?.({
      type: "activity",
      phase: "error",
      activity: described.activity,
      name,
      label: described.label,
      detail: message
    });
    return { error: message };
  }
}

function emitDone(
  ctx: { onEvent?: (event: AgentEvent) => void },
  described: ReturnType<typeof describeActivity>,
  name: string,
  result: unknown
): unknown {
  ctx.onEvent?.({
    type: "activity",
    phase: "done",
    activity: described.activity,
    name,
    label: described.label,
    detail: described.detail
  });
  return result;
}
