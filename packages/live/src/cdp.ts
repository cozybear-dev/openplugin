import { chromium, type Browser, type Page } from "playwright";
import { CDP_URL, RECOVERY_HINT } from "./constants.js";

export type CdpTarget = {
  id?: string;
  type?: string;
  url: string;
  title?: string;
  webSocketDebuggerUrl?: string;
};

export function pickTaskpaneTarget(targets: CdpTarget[]): CdpTarget | undefined {
  const pages = targets.filter((t) => (t.type ?? "page") === "page");
  return (
    pages.find((t) => /localhost:3000\/taskpane\.html/i.test(t.url)) ??
    pages.find((t) => /localhost:3000/i.test(t.url) && /taskpane/i.test(t.url)) ??
    pages.find((t) => /localhost:3000/i.test(t.url))
  );
}

export async function listTargets(cdpUrl = CDP_URL): Promise<CdpTarget[]> {
  const res = await fetch(`${cdpUrl.replace(/\/$/, "")}/json`);
  if (!res.ok) throw new Error(`CDP list failed: ${res.status}. ${RECOVERY_HINT}`);
  return (await res.json()) as CdpTarget[];
}

export async function waitForCdp(opts?: { timeoutMs?: number; cdpUrl?: string }): Promise<CdpTarget[]> {
  const timeoutMs = opts?.timeoutMs ?? 60_000;
  const cdpUrl = opts?.cdpUrl ?? CDP_URL;
  const start = Date.now();
  let last = "";
  while (Date.now() - start < timeoutMs) {
    try {
      const targets = await listTargets(cdpUrl);
      if (targets.length) return targets;
      last = "CDP is up but has no targets yet";
    } catch (err) {
      last = err instanceof Error ? err.message : String(err);
    }
    await delay(500);
  }
  throw new Error(`Timed out waiting for WebView2 CDP at ${cdpUrl}. Last: ${last}. ${RECOVERY_HINT}`);
}

export async function waitForTaskpane(opts?: { timeoutMs?: number; cdpUrl?: string }): Promise<CdpTarget> {
  const timeoutMs = opts?.timeoutMs ?? 90_000;
  const cdpUrl = opts?.cdpUrl ?? CDP_URL;
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const picked = pickTaskpaneTarget(await listTargets(cdpUrl));
      if (picked) return picked;
    } catch {
      /* not up yet */
    }
    await delay(500);
  }
  throw new Error(`Timed out waiting for the OpenPlugin task pane. ${RECOVERY_HINT}`);
}

type PlaywrightConn = {
  close?: () => void;
  _transport?: { close?: () => void; _ws?: { close?: () => void } };
};

function playwrightConnection(browser: object): PlaywrightConn | undefined {
  let proto: object | null = browser;
  for (let i = 0; i < 6 && proto; i++) {
    const conn = (proto as { _connection?: PlaywrightConn })._connection;
    if (conn && (typeof conn.close === "function" || conn._transport)) return conn;
    proto = Object.getPrototypeOf(proto);
  }
  return undefined;
}

/** Drop the Playwright CDP socket. Do not send Browser.close (that kills Office). */
export async function disconnectCdp(browser: Browser): Promise<void> {
  const conn = playwrightConnection(browser);
  try {
    conn?._transport?._ws?.close?.();
  } catch {
    /* already gone */
  }
  try {
    conn?._transport?.close?.();
  } catch {
    /* already gone */
  }
  // Do not call conn.close() — that disposes Playwright's default context and can
  // close the Office WebView. The CLI uses process.exit after the command.
}

export async function connectTaskpane(opts?: { timeoutMs?: number; cdpUrl?: string }): Promise<{
  browser: Browser;
  page: Page;
}> {
  const cdpUrl = opts?.cdpUrl ?? CDP_URL;
  await waitForTaskpane({ timeoutMs: opts?.timeoutMs, cdpUrl });
  const browser = await chromium.connectOverCDP(cdpUrl, { isWebView: true } as never);
  const pages = browser.contexts().flatMap((ctx) => ctx.pages());
  const page =
    pages.find((p) => /localhost:3000\/taskpane\.html/i.test(p.url())) ??
    pages.find((p) => /localhost:3000/i.test(p.url()));
  if (!page) {
    await disconnectCdp(browser);
    throw new Error(`Connected to CDP but no task pane page was found. ${RECOVERY_HINT}`);
  }
  return { browser, page };
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
