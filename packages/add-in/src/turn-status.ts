import { LlmError } from "@openplugin/core";

export function maxStepsLine(steps: number): string {
  return `Stopped after ${steps} steps. Send another message to continue.`;
}

export function timeoutLine(): string {
  return "The model timed out. Send a message to continue.";
}

export function shouldShowTimeout(err: unknown, outerAborted: boolean): boolean {
  return err instanceof LlmError && err.code === "abort" && !outerAborted;
}
