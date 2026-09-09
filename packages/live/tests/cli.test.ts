import { describe, expect, it } from "vitest";
import { parseArgs } from "../src/parse.js";

describe("parseArgs", () => {
  it("parses host, nested action, rest, and flags", () => {
    const cmd = parseArgs(["excel", "app", "send", "add a total", "--mock"]);
    expect(cmd.host).toBe("excel");
    expect(cmd.action).toBe("app");
    expect(cmd.sub).toBe("send");
    expect(cmd.rest).toEqual(["add a total"]);
    expect(cmd.flags.mock).toBe(true);
  });

  it("parses host range with sheet and address", () => {
    const cmd = parseArgs(["excel", "host", "range", "Sheet1", "A1:C3"]);
    expect(cmd.action).toBe("host");
    expect(cmd.sub).toBe("range");
    expect(cmd.rest).toEqual(["Sheet1", "A1:C3"]);
  });

  it("parses real-provider configure flags", () => {
    const cmd = parseArgs([
      "excel",
      "app",
      "configure",
      "--base-url",
      "https://openrouter.ai/api/v1",
      "--model",
      "deepseek/deepseek-v4-flash",
      "--api-key",
      "sk-test"
    ]);
    expect(cmd.flags["base-url"]).toBe("https://openrouter.ai/api/v1");
    expect(cmd.flags.model).toBe("deepseek/deepseek-v4-flash");
    expect(cmd.flags["api-key"]).toBe("sk-test");
  });

  it("rejects an unknown host", () => {
    expect(() => parseArgs(["outlook", "start"])).toThrow(/excel\|word\|powerpoint/);
  });
});
