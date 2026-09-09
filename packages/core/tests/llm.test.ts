import { describe, expect, it } from "vitest";
import { LlmError, chatCompletions, chatUrl, classifyHttpError } from "../src/llm/client.js";
import type { ChatMessage, ProviderConfig } from "../src/llm/types.js";

const config: ProviderConfig = {
  baseUrl: "https://api.example.com/v1",
  apiKey: "sk-test",
  model: "test-model"
};

function sseResponse(chunks: string[], status = 200, headers?: HeadersInit): Response {
  const body = chunks.join("");
  return new Response(body, {
    status,
    headers: { "Content-Type": "text/event-stream", ...headers }
  });
}

describe("chatUrl", () => {
  it("appends /chat/completions to a versioned base URL", () => {
    expect(chatUrl("https://api.x.ai/v1")).toBe("https://api.x.ai/v1/chat/completions");
  });

  it("does not duplicate when the base already targets chat/completions", () => {
    expect(
      chatUrl("https://res.openai.azure.com/openai/deployments/gpt/chat/completions?api-version=2024-02-15-preview")
    ).toBe(
      "https://res.openai.azure.com/openai/deployments/gpt/chat/completions?api-version=2024-02-15-preview"
    );
  });
});

describe("classifyHttpError", () => {
  it("maps 401 to auth", () => {
    expect(classifyHttpError(401, "nope").code).toBe("auth");
  });

  it("maps context-length bodies to context_length", () => {
    expect(classifyHttpError(400, "This model's maximum context length is 8192 tokens").code).toBe(
      "context_length"
    );
  });
});

describe("chatCompletions", () => {
  const messages: ChatMessage[] = [{ role: "user", content: "hi" }];

  it("streams text deltas and returns the assembled assistant message", async () => {
    const fetchImpl: typeof fetch = async () =>
      sseResponse([
        'data: {"choices":[{"delta":{"content":"Hel"}}]}\n\n',
        'data: {"choices":[{"delta":{"content":"lo"}}]}\n\n',
        "data: [DONE]\n\n"
      ]);

    const deltas: string[] = [];
    const result = await chatCompletions({
      config,
      messages,
      fetchImpl,
      onEvent: (e) => {
        if (e.type === "text") deltas.push(e.delta);
      }
    });

    expect(deltas.join("")).toBe("Hello");
    expect(result.message).toEqual({ role: "assistant", content: "Hello" });
  });

  it("assembles streamed tool calls", async () => {
    const fetchImpl: typeof fetch = async () =>
      sseResponse([
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","type":"function","function":{"name":"excel.readRange","arguments":""}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"a"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"ddress\\":\\"A1\\"}"}}]}}]}\n\n',
        "data: [DONE]\n\n"
      ]);

    const result = await chatCompletions({ config, messages, fetchImpl });
    expect(result.message.role).toBe("assistant");
    if (result.message.role !== "assistant") throw new Error("expected assistant");
    expect(result.message.tool_calls).toEqual([
      {
        id: "call_1",
        type: "function",
        function: { name: "excel.readRange", arguments: '{"address":"A1"}' }
      }
    ]);
  });

  it("throws LlmError cors when fetch reports a CORS failure", async () => {
    const fetchImpl: typeof fetch = async () => {
      throw new TypeError("Failed to fetch");
    };
    await expect(chatCompletions({ config, messages, fetchImpl })).rejects.toMatchObject({
      name: "LlmError",
      code: "cors"
    } satisfies Partial<LlmError>);
  });

  it("throws auth on 401", async () => {
    const fetchImpl: typeof fetch = async () => new Response("unauthorized", { status: 401 });
    await expect(chatCompletions({ config, messages, fetchImpl })).rejects.toMatchObject({
      code: "auth"
    });
  });

  it("sends Authorization bearer and tools payload", async () => {
    let captured: { url: string; headers: Headers; body: unknown } | undefined;
    const fetchImpl: typeof fetch = async (input, init) => {
      captured = {
        url: String(input),
        headers: new Headers(init?.headers),
        body: JSON.parse(String(init?.body))
      };
      return sseResponse(["data: {\"choices\":[{\"delta\":{\"content\":\"ok\"}}]}\n\n", "data: [DONE]\n\n"]);
    };

    await chatCompletions({
      config,
      messages,
      tools: [
        {
          type: "function",
          function: { name: "skills.list", description: "List skills", parameters: { type: "object", properties: {} } }
        }
      ],
      fetchImpl
    });

    expect(captured?.url).toBe("https://api.example.com/v1/chat/completions");
    expect(captured?.headers.get("Authorization")).toBe("Bearer sk-test");
    expect(captured?.body).toMatchObject({
      model: "test-model",
      stream: true,
      tools: [{ type: "function", function: { name: "skills.list" } }]
    });
  });

  it("uses api-key header instead of bearer when provided", async () => {
    let headers: Headers | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      headers = new Headers(init?.headers);
      return sseResponse(["data: {\"choices\":[{\"delta\":{}}]}\n\n", "data: [DONE]\n\n"]);
    };
    await chatCompletions({
      config: { ...config, headers: { "api-key": "azure-key" } },
      messages,
      fetchImpl
    });
    expect(headers?.get("api-key")).toBe("azure-key");
    expect(headers?.get("Authorization")).toBeNull();
  });

  it("adds OpenRouter etiquette headers", async () => {
    let headers: Headers | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      headers = new Headers(init?.headers);
      return sseResponse(["data: {\"choices\":[{\"delta\":{}}]}\n\n", "data: [DONE]\n\n"]);
    };
    await chatCompletions({
      config: { ...config, baseUrl: "https://openrouter.ai/api/v1", apiKey: "sk-or" },
      messages,
      fetchImpl
    });
    expect(headers?.get("X-Title")).toBe("OpenPlugin");
    expect(headers?.get("HTTP-Referer")).toBe("https://openplugin.local");
  });

  it("aborts when the signal fires", async () => {
    const controller = new AbortController();
    const fetchImpl: typeof fetch = async (_input, init) => {
      return new Promise((_resolve, reject) => {
        init?.signal?.addEventListener("abort", () => {
          const err = new Error("aborted");
          err.name = "AbortError";
          reject(err);
        });
      });
    };
    const pending = chatCompletions({ config, messages, fetchImpl, signal: controller.signal });
    controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "abort" });
  });

  it("returns finish_reason from an SSE choice", async () => {
    const fetchImpl: typeof fetch = async () =>
      sseResponse([
        'data: {"choices":[{"delta":{"content":"Section one."},"finish_reason":"length"}]}\n\n',
        "data: [DONE]\n\n"
      ]);
    const result = await chatCompletions({ config, messages, fetchImpl });
    expect(result.message.content).toBe("Section one.");
    expect(result.finishReason).toBe("length");
  });

  it("returns finish_reason from a non-stream JSON body", async () => {
    const fetchImpl: typeof fetch = async () =>
      new Response(
        JSON.stringify({
          choices: [{ message: { role: "assistant", content: "Hi" }, finish_reason: "stop" }]
        }),
        { headers: { "Content-Type": "application/json" } }
      );
    const result = await chatCompletions({ config, messages, stream: false, fetchImpl });
    expect(result.finishReason).toBe("stop");
  });

  it("isTruncatedFinish treats length and max_tokens as truncated", async () => {
    const { isTruncatedFinish } = await import("../src/llm/client.js");
    expect(isTruncatedFinish("length")).toBe(true);
    expect(isTruncatedFinish("max_tokens")).toBe(true);
    expect(isTruncatedFinish("stop")).toBe(false);
    expect(isTruncatedFinish(undefined)).toBe(false);
  });
});
