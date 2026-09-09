import { describe, expect, it } from "vitest";
import { nativeSearchTool, resolveSearchBackend, runWebSearch } from "../src/search/index.js";

describe("resolveSearchBackend", () => {
  it("picks native for OpenRouter in auto mode", () => {
    expect(
      resolveSearchBackend({
        backend: "auto",
        baseUrl: "https://openrouter.ai/api/v1",
        exaApiKey: undefined
      })
    ).toBe("native");
  });

  it("picks duckduckgo when native is unavailable", () => {
    expect(
      resolveSearchBackend({
        backend: "auto",
        baseUrl: "http://127.0.0.1:11434/v1",
        exaApiKey: undefined
      })
    ).toBe("duckduckgo");
  });

  it("picks exa in auto when native is unavailable and a key is set", () => {
    expect(
      resolveSearchBackend({
        backend: "auto",
        baseUrl: "http://127.0.0.1:11434/v1",
        exaApiKey: "exa-key"
      })
    ).toBe("exa");
  });

  it("honors an explicit backend", () => {
    expect(
      resolveSearchBackend({
        backend: "exa",
        baseUrl: "https://openrouter.ai/api/v1",
        exaApiKey: "exa-key"
      })
    ).toBe("exa");
  });
});

describe("nativeSearchTool", () => {
  it("returns an OpenRouter server tool", () => {
    expect(nativeSearchTool("https://openrouter.ai/api/v1")).toEqual({ type: "openrouter:web_search" });
  });
});

describe("runWebSearch", () => {
  it("returns a settings hint when the free engine is rate-limited", async () => {
    const result = await runWebSearch({
      backend: "duckduckgo",
      query: "test",
      fetchImpl: async () => new Response("blocked", { status: 429 })
    });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error("expected failure");
    expect(result.error).toMatch(/rate-limited/i);
    expect(result.error).toMatch(/Exa/i);
  });

  it("calls Exa with highlights", async () => {
    const calls: Array<{ url: string; body: unknown }> = [];
    const result = await runWebSearch({
      backend: "exa",
      query: "gdp",
      apiKey: "exa-secret",
      fetchImpl: async (input, init) => {
        calls.push({ url: String(input), body: JSON.parse(String(init?.body)) });
        return new Response(
          JSON.stringify({
            results: [{ title: "GDP", url: "https://example.com", highlights: ["grew"] }]
          }),
          { status: 200, headers: { "Content-Type": "application/json" } }
        );
      }
    });
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error("expected ok");
    expect(result.results[0]).toMatchObject({ title: "GDP", url: "https://example.com" });
    expect(calls[0]?.url).toContain("api.exa.ai/search");
    expect(calls[0]?.body).toMatchObject({ query: "gdp" });
  });
});
