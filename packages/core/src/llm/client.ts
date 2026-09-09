import {
  LlmError,
  type ChatMessage,
  type ChatResult,
  type LlmErrorCode,
  type ProviderConfig,
  type StreamEvent,
  type ToolCall,
  type ToolDefinition
} from "./types.js";

export { LlmError } from "./types.js";

export function chatUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim();
  if (/\/chat\/completions(\?|$)/i.test(trimmed)) return trimmed;
  return `${trimmed.replace(/\/+$/, "")}/chat/completions`;
}

export function classifyHttpError(status: number, body: string): LlmError {
  const lower = body.toLowerCase();
  if (status === 401 || status === 403) {
    return new LlmError("The endpoint rejected the API key.", "auth", status);
  }
  if (status === 404) {
    return new LlmError(`Endpoint not found (${status}). Check the base URL.`, "http", status);
  }
  if (
    lower.includes("maximum context length") ||
    lower.includes("context_length") ||
    lower.includes("context window") ||
    lower.includes("too many tokens")
  ) {
    return new LlmError("The prompt exceeded the model's context length.", "context_length", status);
  }
  return new LlmError(`HTTP ${status}: ${body.slice(0, 400) || "request failed"}`, "http", status);
}

function isAbort(err: unknown): boolean {
  return (
    (err instanceof DOMException && err.name === "AbortError") ||
    (err instanceof Error && (err.name === "AbortError" || err.message.toLowerCase().includes("aborted")))
  );
}

function isCorsLike(err: unknown): boolean {
  if (!(err instanceof Error)) return false;
  const msg = err.message.toLowerCase();
  return (
    err.name === "TypeError" ||
    msg.includes("failed to fetch") ||
    msg.includes("load failed") ||
    msg.includes("networkerror") ||
    msg.includes("cors")
  );
}

type PendingTool = { id: string; name: string; arguments: string };

export async function chatCompletions(opts: {
  config: ProviderConfig;
  messages: ChatMessage[];
  tools?: ToolDefinition[];
  toolChoice?: "auto" | "none" | "required";
  stream?: boolean;
  fetchImpl?: typeof fetch;
  signal?: AbortSignal;
  onEvent?: (event: StreamEvent) => void;
}): Promise<ChatResult> {
  const fetchImpl = opts.fetchImpl ?? globalThis.fetch;
  if (!fetchImpl) throw new LlmError("fetch is not available", "network");

  const headers = new Headers({
    "Content-Type": "application/json",
    ...(opts.config.headers ?? {})
  });
  if (opts.config.apiKey && !headers.has("api-key") && !headers.has("Authorization")) {
    headers.set("Authorization", `Bearer ${opts.config.apiKey}`);
  }
  if (/openrouter\.ai/i.test(opts.config.baseUrl)) {
    if (!headers.has("HTTP-Referer")) headers.set("HTTP-Referer", "https://openplugin.local");
    if (!headers.has("X-Title")) headers.set("X-Title", "OpenPlugin");
  }

  const body: Record<string, unknown> = {
    model: opts.config.model,
    messages: opts.messages,
    stream: opts.stream !== false
  };
  if (opts.tools && opts.tools.length > 0) {
    body.tools = opts.tools;
    body.tool_choice = opts.toolChoice ?? "auto";
  }
  if (opts.config.maxOutputTokens) body.max_tokens = opts.config.maxOutputTokens;

  const controller = new AbortController();
  const timeout = setTimeout(
    () => controller.abort(),
    opts.config.timeoutMs ?? 120_000
  );
  const onOuterAbort = () => controller.abort();
  opts.signal?.addEventListener("abort", onOuterAbort);

  let response: Response;
  try {
    response = await fetchImpl(chatUrl(opts.config.baseUrl), {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: controller.signal
    });
  } catch (err) {
    clearTimeout(timeout);
    opts.signal?.removeEventListener("abort", onOuterAbort);
    if (isAbort(err)) throw new LlmError("Request aborted.", "abort");
    if (isCorsLike(err)) {
      throw new LlmError(
        "The endpoint blocked the request (CORS or network). Enable CORS on the server, or use the OpenPlugin companion.",
        "cors"
      );
    }
    throw new LlmError(err instanceof Error ? err.message : "Network error", "network");
  }

  if (!response.ok) {
    clearTimeout(timeout);
    opts.signal?.removeEventListener("abort", onOuterAbort);
    const text = await response.text().catch(() => "");
    throw classifyHttpError(response.status, text);
  }

  try {
    if (opts.stream === false) {
      const json = (await response.json()) as {
        choices?: Array<{ message?: { content?: string; tool_calls?: ToolCall[] } }>;
      };
      const msg = json.choices?.[0]?.message;
      const message: Extract<ChatMessage, { role: "assistant" }> = {
        role: "assistant",
        content: msg?.content ?? "",
        ...(msg?.tool_calls ? { tool_calls: msg.tool_calls } : {})
      };
      opts.onEvent?.({ type: "done", message });
      return { message };
    }
    return await readSse(response, opts.onEvent);
  } catch (err) {
    if (err instanceof LlmError) throw err;
    if (isAbort(err)) throw new LlmError("Request aborted.", "abort");
    throw new LlmError(err instanceof Error ? err.message : "Failed to parse the model stream.", "parse");
  } finally {
    clearTimeout(timeout);
    opts.signal?.removeEventListener("abort", onOuterAbort);
  }
}

async function readSse(
  response: Response,
  onEvent?: (event: StreamEvent) => void
): Promise<ChatResult> {
  const text = await response.text();
  const tools = new Map<number, PendingTool>();
  let content = "";

  for (const block of text.split(/\n\n/)) {
    for (const rawLine of block.split("\n")) {
      const line = rawLine.trim();
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).trim();
      if (!data || data === "[DONE]") continue;
      let parsed: {
        error?: { message?: string };
        choices?: Array<{
          delta?: {
            content?: string | null;
            tool_calls?: Array<{
              index?: number;
              id?: string;
              type?: string;
              function?: { name?: string; arguments?: string };
            }>;
          };
          message?: { content?: string; tool_calls?: ToolCall[] };
        }>;
      };
      try {
        parsed = JSON.parse(data);
      } catch {
        continue;
      }
      if (parsed.error?.message) {
        throw new LlmError(parsed.error.message, "http");
      }
      const delta = parsed.choices?.[0]?.delta;
      const whole = parsed.choices?.[0]?.message;
      if (typeof delta?.content === "string" && delta.content.length > 0) {
        content += delta.content;
        onEvent?.({ type: "text", delta: delta.content });
      }
      if (typeof whole?.content === "string" && !delta?.content) {
        content += whole.content;
        onEvent?.({ type: "text", delta: whole.content });
      }
      for (const tc of delta?.tool_calls ?? []) {
        const index = tc.index ?? 0;
        const current = tools.get(index) ?? { id: "", name: "", arguments: "" };
        if (tc.id) current.id = tc.id;
        if (tc.function?.name) current.name = tc.function.name;
        if (tc.function?.arguments) current.arguments += tc.function.arguments;
        tools.set(index, current);
        onEvent?.({
          type: "tool_call",
          index,
          id: tc.id,
          name: tc.function?.name,
          argumentsDelta: tc.function?.arguments ?? ""
        });
      }
      if (whole?.tool_calls) {
        whole.tool_calls.forEach((tc, index) => {
          tools.set(index, {
            id: tc.id,
            name: tc.function.name,
            arguments: tc.function.arguments
          });
        });
      }
    }
  }

  const tool_calls: ToolCall[] = [...tools.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, t]) => ({
      id: t.id || `call_${t.name}`,
      type: "function" as const,
      function: { name: t.name, arguments: t.arguments }
    }))
    .filter((t) => t.function.name);

  const message: Extract<ChatMessage, { role: "assistant" }> = {
    role: "assistant",
    content: content || (tool_calls.length ? null : ""),
    ...(tool_calls.length ? { tool_calls } : {})
  };
  onEvent?.({ type: "done", message });
  return { message };
}

export function asErrorCode(err: unknown): LlmErrorCode | undefined {
  return err instanceof LlmError ? err.code : undefined;
}
