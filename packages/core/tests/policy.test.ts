import { describe, expect, it } from "vitest";
import { assertPolicy, mergePolicy, OPEN_POLICY } from "../src/policy.js";

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

describe("mergePolicy", () => {
  it("does not let a user widen the tenant allowlist", () => {
    const tenant = {
      allowedEndpoints: ["https://llm.corp.example"],
      allowedModels: ["qwen2.5-32b"],
      denyExecuteJs: true
    };
    const merged = mergePolicy(tenant, {
      allowedEndpoints: ["https://evil.example"],
      allowedModels: ["gpt-4o", "qwen2.5-32b"],
      denyExecuteJs: false
    });
    expect(merged.allowedEndpoints).toEqual(["https://llm.corp.example"]);
    expect(merged.allowedModels).toEqual(["qwen2.5-32b"]);
    expect(merged.denyExecuteJs).toBe(true);
  });
});
