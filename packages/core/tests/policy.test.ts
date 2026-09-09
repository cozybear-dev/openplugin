import { describe, expect, it } from "vitest";
import { assertPolicy, OPEN_POLICY } from "../src/policy.js";

describe("assertPolicy", () => {
  it("allows any endpoint under the open policy", () => {
    expect(() =>
      assertPolicy(OPEN_POLICY, { baseUrl: "https://api.x.ai/v1", model: "grok-4.5" })
    ).not.toThrow();
  });

  it("blocks executeOfficeJs when denyExecuteJs is set", () => {
    expect(() =>
      assertPolicy(OPEN_POLICY, {
        baseUrl: "https://api.x.ai/v1",
        model: "grok-4.5",
        executeJs: true
      })
    ).toThrow(/executeOfficeJs/);
  });

  it("enforces endpoint and model allowlists", () => {
    const policy = {
      allowedEndpoints: ["https://llm.corp.example"],
      allowedModels: ["qwen2.5-32b"],
      denyExecuteJs: true
    };
    expect(() =>
      assertPolicy(policy, { baseUrl: "https://evil.example/v1", model: "qwen2.5-32b" })
    ).toThrow(/endpoint/i);
    expect(() =>
      assertPolicy(policy, { baseUrl: "https://llm.corp.example/v1", model: "gpt-4o" })
    ).toThrow(/model/i);
    expect(() =>
      assertPolicy(policy, { baseUrl: "https://llm.corp.example/v1", model: "qwen2.5-32b" })
    ).not.toThrow();
  });
});
