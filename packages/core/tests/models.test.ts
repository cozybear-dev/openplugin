import { describe, expect, it } from "vitest";
import { LlmError, listModels, modelsUrl } from "../src/llm/client.js";
import type { ProviderConfig } from "../src/llm/types.js";

const config: ProviderConfig = {
  baseUrl: "https://api.example.com/v1",
  apiKey: "sk-test",
  model: "test-model"
};

describe("modelsUrl", () => {
  it("appends /models to a versioned base URL", () => {
    expect(modelsUrl("https://api.x.ai/v1")).toBe("https://api.x.ai/v1/models");
  });

  it("replaces /chat/completions with /models and keeps Azure query string", () => {
    expect(
      modelsUrl(
        "https://res.openai.azure.com/openai/deployments/gpt/chat/completions?api-version=2024-02-15-preview"
      )
    ).toBe("https://res.openai.azure.com/openai/deployments/gpt/models?api-version=2024-02-15-preview");
  });

  it("does not double-append when the base already ends in /models", () => {
    expect(modelsUrl("https://api.example.com/v1/models")).toBe("https://api.example.com/v1/models");
  });

  it("does not double-append when /models is followed by a query string", () => {
    expect(modelsUrl("https://api.example.com/v1/models?api-version=2024-02-15-preview")).toBe(
      "https://api.example.com/v1/models?api-version=2024-02-15-preview"
    );
  });
});

describe("listModels", () => {
  it("GETs /models with Bearer and returns sorted unique ids", async () => {
    let captured: { url: string; method?: string; headers: Headers } | undefined;
    const fetchImpl: typeof fetch = async (input, init) => {
      captured = {
        url: String(input),
        method: init?.method,
        headers: new Headers(init?.headers)
      };
      return new Response(
        JSON.stringify({
          data: [
            { id: "z-model", name: "Z" },
            { id: "a-model" },
            { id: "a-model" },
            { id: "m-model", name: "M" }
          ]
        }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    };

    const models = await listModels({ config, fetchImpl });

    expect(captured?.url).toBe("https://api.example.com/v1/models");
    expect(captured?.method).toBe("GET");
    expect(captured?.headers.get("Authorization")).toBe("Bearer sk-test");
    expect(models.map((m) => m.id)).toEqual(["a-model", "m-model", "z-model"]);
    expect(models.find((m) => m.id === "m-model")?.name).toBe("M");
  });

  it("accepts Ollama-style { models: [{ name }] } payloads", async () => {
    const fetchImpl: typeof fetch = async () =>
      new Response(JSON.stringify({ models: [{ name: "llama3" }, { name: "mistral" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });

    const models = await listModels({ config, fetchImpl });
    expect(models).toEqual([
      { id: "llama3", name: "llama3" },
      { id: "mistral", name: "mistral" }
    ]);
  });

  it("uses api-key header instead of bearer when provided", async () => {
    let headers: Headers | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      headers = new Headers(init?.headers);
      return new Response(JSON.stringify({ data: [] }), { status: 200 });
    };
    await listModels({
      config: { ...config, headers: { "api-key": "azure-key" } },
      fetchImpl
    });
    expect(headers?.get("api-key")).toBe("azure-key");
    expect(headers?.get("Authorization")).toBeNull();
  });

  it("adds OpenRouter etiquette headers", async () => {
    let headers: Headers | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      headers = new Headers(init?.headers);
      return new Response(JSON.stringify({ data: [] }), { status: 200 });
    };
    await listModels({
      config: { ...config, baseUrl: "https://openrouter.ai/api/v1", apiKey: "sk-or" },
      fetchImpl
    });
    expect(headers?.get("X-Title")).toBe("OpenPlugin");
    expect(headers?.get("HTTP-Referer")).toBe("https://openplugin.local");
  });

  it("throws LlmError cors when fetch reports a CORS failure", async () => {
    const fetchImpl: typeof fetch = async () => {
      throw new TypeError("Failed to fetch");
    };
    await expect(listModels({ config, fetchImpl })).rejects.toMatchObject({
      name: "LlmError",
      code: "cors"
    } satisfies Partial<LlmError>);
  });

  it("throws auth on 401", async () => {
    const fetchImpl: typeof fetch = async () => new Response("unauthorized", { status: 401 });
    await expect(listModels({ config, fetchImpl })).rejects.toMatchObject({
      code: "auth"
    });
  });
});
