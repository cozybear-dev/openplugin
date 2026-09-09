import { describe, expect, it } from "vitest";
import { forwardHeaders, joinTarget } from "../src/proxy.ts";

describe("joinTarget", () => {
  it("avoids doubling /v1 when the target already ends in /v1", () => {
    expect(joinTarget("http://127.0.0.1:11434/v1", "/v1/chat/completions")).toBe(
      "http://127.0.0.1:11434/v1/chat/completions"
    );
  });

  it("appends the path when the target has no /v1 suffix", () => {
    expect(joinTarget("https://openrouter.ai/api/v1", "/v1/models")).toBe(
      "https://openrouter.ai/api/v1/models"
    );
  });
});

describe("forwardHeaders", () => {
  it("strips hop-by-hop and pairing headers", () => {
    const headers = new Headers({
      Authorization: "Bearer x",
      Host: "127.0.0.1:8788",
      Connection: "keep-alive",
      "x-openplugin-token": "secret",
      "x-openplugin-target": "http://127.0.0.1:11434/v1"
    });
    const out = forwardHeaders(headers);
    expect(out.get("Authorization")).toBe("Bearer x");
    expect(out.get("host")).toBeNull();
    expect(out.get("connection")).toBeNull();
    expect(out.get("x-openplugin-token")).toBeNull();
    expect(out.get("x-openplugin-target")).toBeNull();
  });
});
