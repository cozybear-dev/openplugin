import type { LiveCommand } from "./parse.js";
import { MOCK_LLM_ORIGIN } from "./mock-llm.js";
import { requireState } from "./state.js";
import { withTaskpane } from "./session.js";
import type { Page } from "playwright";

type LiveDump = {
  hostKind: string;
  inOffice: boolean;
  tab: string;
  busy: boolean;
  input: string;
  provider: { baseUrl: string; model: string };
  pending: Array<{ id: string; title: string }>;
  lastLine?: string;
};

export async function runApp(cmd: LiveCommand): Promise<void> {
  const sub = cmd.sub;
  if (sub === "state") {
    const state = await withTaskpane((page) => liveCall<LiveDump>(page, "getState"));
    console.log(JSON.stringify(state, null, 2));
    return;
  }
  if (sub === "configure") {
    const session = await requireState();
    const realBase = typeof cmd.flags["base-url"] === "string" ? cmd.flags["base-url"] : "";
    const origin = realBase
      || (typeof cmd.flags.mock === "string" ? cmd.flags.mock : session.mockOrigin || MOCK_LLM_ORIGIN);
    const model = realBase
      ? String(cmd.flags.model || "")
      : "openplugin-live";
    const apiKey = realBase ? String(cmd.flags["api-key"] || "") : "live";
    const dump = await withTaskpane(async (page) => {
      await liveCall(page, "configure", {
        baseUrl: origin,
        model,
        apiKey
      });
      return liveCall<LiveDump>(page, "getState");
    });
    console.log(JSON.stringify({ configured: origin, state: dump }, null, 2));
    return;
  }
  if (sub === "send") {
    const text = cmd.rest.join(" ");
    if (!text) throw new Error("Usage: live <host> app send <prompt>");
    const dump = await withTaskpane(async (page) => {
      await liveCall(page, "send", text);
      await waitUntilIdle(page);
      return liveCall<LiveDump>(page, "getState");
    });
    console.log(JSON.stringify(dump, null, 2));
    return;
  }
  if (sub === "apply") {
    const dump = await withTaskpane(async (page) => {
      await liveCall(page, "apply");
      await waitUntilIdle(page);
      return liveCall<LiveDump>(page, "getState");
    });
    console.log(JSON.stringify(dump, null, 2));
    return;
  }
  throw new Error(`Unknown app command '${sub}'. Use state|configure|send|apply`);
}

async function liveCall<T>(page: Page, method: string, ...args: unknown[]): Promise<T> {
  await page.waitForFunction(() => Boolean((window as { __openpluginLive?: unknown }).__openpluginLive), {
    timeout: 30_000
  });
  return page.evaluate(
    async ({ methodName, callArgs }) => {
      const api = (window as unknown as { __openpluginLive: Record<string, (...a: unknown[]) => unknown> })
        .__openpluginLive;
      const fn = api[methodName];
      if (!fn) throw new Error(`live bridge has no method ${methodName}`);
      return await fn(...callArgs);
    },
    { methodName: method, callArgs: args }
  ) as Promise<T>;
}

async function waitUntilIdle(page: Page, timeoutMs = 180_000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const state = await liveCall<LiveDump>(page, "getState");
    if (!state.busy) return;
    await page.waitForTimeout(400);
  }
  throw new Error("Timed out waiting for the OpenPlugin turn to finish.");
}
