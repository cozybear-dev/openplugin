import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import net from "node:net";
import { join } from "node:path";
import { generateSideloadFile, registerAddIn, unregisterAddIn } from "office-addin-dev-settings";
import { OfficeAddinManifest, OfficeApp } from "office-addin-manifest";
import { connectTaskpane, disconnectCdp, waitForCdp, waitForTaskpane } from "./cdp.js";
import { disableOfficeJsDebugger, watchBlockingDialogs } from "./dialogs.js";
import { ADDIN_ID, OFFICE_EXES, OFFICE_PATHS, RECOVERY_HINT, VITE_PORT, WEBVIEW_ARGS } from "./constants.js";
import { MOCK_LLM_ORIGIN, MOCK_LLM_PORT } from "./mock-llm.js";
import { addInDir, liveManifestPath, outputDir, repoRoot } from "./paths.js";
import type { LiveHost } from "./parse.js";
import { ensureMockLlm } from "./server.js";
import { clearState, writeState } from "./state.js";

const OFFICE_APP: Record<LiveHost, OfficeApp> = {
  excel: OfficeApp.Excel,
  word: OfficeApp.Word,
  powerpoint: OfficeApp.PowerPoint
};

export async function startLive(host: LiveHost, flags: Record<string, string | boolean>): Promise<void> {
  const force = Boolean(flags.force);
  await ensureCerts();
  const viteStarted = await ensureVite();
  const mockStarted = await ensureMock();
  if (force) await killOffice(host);
  else await assertOfficeNotBlocking(host);
  console.error(`[live] vite=${viteStarted ? "started" : "reused"} mock=${mockStarted ? "started" : "reused"}`);

  const manifestPath = liveManifestPath();
  await registerAddIn(manifestPath);
  // Direct debugger (the default for enableDebugging) shows "WebView Stop On Load"
  // and blocks the task pane until a human clicks. CDP attach does not need it.
  await disableOfficeJsDebugger(ADDIN_ID);

  const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
  const document = await generateSideloadFile(OFFICE_APP[host], manifest);
  const exe = officeExe(host);
  const stopWatch = watchBlockingDialogs();
  try {
    launchOffice(exe, document);
    await waitForCdp({ timeoutMs: 90_000 });
    const target = await waitForTaskpane({ timeoutMs: 90_000 });
    console.error(`[live] task pane ${target.url}`);
    const { browser, page } = await connectTaskpane({ timeoutMs: 15_000 });
    try {
      const shot = join(outputDir(), `start-${host}.png`);
      await mkdir(outputDir(), { recursive: true });
      await page.screenshot({ path: shot, fullPage: true });
      console.error(`[live] wrote ${shot}`);
    } finally {
      await disconnectCdp(browser);
    }
  } finally {
    stopWatch();
  }

  await mkdir(outputDir(), { recursive: true });
  await writeState({
    host,
    cdp: "http://127.0.0.1:9222",
    mockOrigin: MOCK_LLM_ORIGIN,
    document,
    startedAt: new Date().toISOString(),
    viteStarted,
    mockStarted
  });
  console.log(`Live ${host} session ready. Task pane attached. Mock LLM at ${MOCK_LLM_ORIGIN}`);
}

export async function stopLive(host: LiveHost, flags: Record<string, string | boolean>): Promise<void> {
  try {
    await unregisterAddIn(liveManifestPath());
  } catch {
    /* not registered */
  }
  if (flags.force !== false) await killOffice(host);
  await clearState();
  console.log(`Stopped live ${host} session.`);
}

export async function statusLive(): Promise<void> {
  const { readState } = await import("./state.js");
  const state = await readState();
  if (!state) {
    console.log("No live session.");
    return;
  }
  let cdp = false;
  try {
    const res = await fetch("http://127.0.0.1:9222/json");
    cdp = res.ok;
  } catch {
    cdp = false;
  }
  console.log(JSON.stringify({ ...state, cdp }, null, 2));
}

function launchOffice(exe: string, document: string): void {
  const child = spawn(exe, [document], {
    detached: true,
    stdio: "ignore",
    env: {
      ...process.env,
      WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: WEBVIEW_ARGS
    }
  });
  child.on("error", (err) => {
    console.error(`Failed to launch ${exe}: ${err.message}`);
  });
  child.unref();
}

function officeExe(host: LiveHost): string {
  const configured = OFFICE_PATHS[host];
  if (existsSync(configured)) return configured;
  throw new Error(`${host} was not found at ${configured}. Install Microsoft 365 desktop.`);
}

async function ensureCerts(): Promise<void> {
  try {
    const devCerts = await import("office-addin-dev-certs");
    const api = (devCerts.default ?? devCerts) as { ensureCertificatesAreInstalled?: () => Promise<unknown> };
    await api.ensureCertificatesAreInstalled?.();
  } catch (err) {
    console.warn("Could not ensure Office add-in HTTPS certificates.", err);
  }
}

async function ensureVite(): Promise<boolean> {
  if (await isListening(VITE_PORT)) return false;
  const viteBin = join(repoRoot(), "node_modules", "vite", "bin", "vite.js");
  spawn(process.execPath, [viteBin, "--port", String(VITE_PORT)], {
    cwd: addInDir(),
    detached: true,
    stdio: "ignore"
  }).unref();
  const start = Date.now();
  while (Date.now() - start < 45_000) {
    if (await isListening(VITE_PORT)) return true;
    await delay(400);
  }
  throw new Error(`Vite did not start on port ${VITE_PORT}. From ${addInDir()} run npm run dev.`);
}

async function ensureMock(): Promise<boolean> {
  try {
    const res = await fetch(`http://127.0.0.1:${MOCK_LLM_PORT}/v1/models`);
    if (res.ok) return false;
  } catch {
    /* start */
  }
  const tsxCli = join(repoRoot(), "node_modules", "tsx", "dist", "cli.mjs");
  const child = spawn(process.execPath, [tsxCli, join(repoRoot(), "packages", "live", "src", "mock-main.ts")], {
    detached: true,
    stdio: "ignore",
    env: process.env
  });
  child.unref();
  const start = Date.now();
  while (Date.now() - start < 10_000) {
    try {
      const res = await fetch(`http://127.0.0.1:${MOCK_LLM_PORT}/v1/models`);
      if (res.ok) return true;
    } catch {
      /* wait */
    }
    await delay(200);
  }
  await ensureMockLlm();
  return true;
}

async function assertOfficeNotBlocking(host: LiveHost): Promise<void> {
  if (!(await isOfficeRunning(host))) return;
  try {
    const res = await fetch("http://127.0.0.1:9222/json");
    if (res.ok) return;
  } catch {
    /* */
  }
  throw new Error(
    `${OFFICE_EXES[host]} is already running without the WebView2 debug port. Close it, or pass --force. ${RECOVERY_HINT}`
  );
}

async function isOfficeRunning(host: LiveHost): Promise<boolean> {
  return new Promise((resolve) => {
    const child = spawn("tasklist", ["/FI", `IMAGENAME eq ${OFFICE_EXES[host]}`, "/NH"], { windowsHide: true });
    let out = "";
    child.stdout?.on("data", (d) => {
      out += String(d);
    });
    child.on("close", () => resolve(out.toLowerCase().includes(OFFICE_EXES[host].toLowerCase())));
    child.on("error", () => resolve(false));
  });
}

async function killOffice(host: LiveHost): Promise<void> {
  await new Promise<void>((resolve) => {
    const child = spawn("taskkill", ["/F", "/IM", OFFICE_EXES[host]], { windowsHide: true, stdio: "ignore" });
    child.on("close", () => resolve());
    child.on("error", () => resolve());
  });
  await delay(800);
}

function isListening(port: number): Promise<boolean> {
  return Promise.all([tryConnect("127.0.0.1", port), tryConnect("::1", port)]).then((hits) => hits.some(Boolean));
}

function tryConnect(host: string, port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const socket = net.connect({ port, host }, () => {
      socket.end();
      resolve(true);
    });
    socket.on("error", () => resolve(false));
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
