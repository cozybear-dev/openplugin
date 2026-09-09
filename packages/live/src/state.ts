import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { LiveHost } from "./parse.js";

export type LiveState = {
  host: LiveHost;
  cdp: string;
  mockOrigin: string;
  document?: string;
  startedAt: string;
  viteStarted: boolean;
  mockStarted: boolean;
};

function dir(): string {
  return join(tmpdir(), "openplugin-live");
}

export function statePath(): string {
  return join(dir(), "state.json");
}

export async function writeState(state: LiveState): Promise<void> {
  await mkdir(dir(), { recursive: true });
  await writeFile(statePath(), JSON.stringify(state, null, 2), "utf8");
}

export async function readState(): Promise<LiveState | null> {
  try {
    const raw = await readFile(statePath(), "utf8");
    return JSON.parse(raw) as LiveState;
  } catch {
    return null;
  }
}

export async function clearState(): Promise<void> {
  await rm(statePath(), { force: true });
}

export async function requireState(): Promise<LiveState> {
  const state = await readState();
  if (!state) throw new Error("No live session. Run: npm run live -- excel start");
  return state;
}
