import { LlmError } from "@openplugin/core";
import { describe, expect, it } from "vitest";
import { maxStepsLine, shouldShowTimeout, timeoutLine } from "../src/turn-status";

describe("turn-status", () => {
  it("formats max-steps and timeout copy", () => {
    expect(maxStepsLine(16)).toBe("Stopped after 16 steps. Send another message to continue.");
    expect(timeoutLine()).toBe("The model timed out. Send a message to continue.");
  });

  it("shows timeout only when the outer signal was not aborted", () => {
    const err = new LlmError("Request aborted.", "abort");
    expect(shouldShowTimeout(err, false)).toBe(true);
    expect(shouldShowTimeout(err, true)).toBe(false);
    expect(shouldShowTimeout(new Error("nope"), false)).toBe(false);
  });
});
