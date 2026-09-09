import { existsSync, mkdirSync, readFileSync, writeFileSync, appendFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { homedir } from "node:os";
import { randomBytes } from "node:crypto";
import type { Policy } from "@openplugin/core";

export function dataDir(): string {
  const override = process.env.OPENPLUGIN_HOME;
  if (override) return override;
  if (process.platform === "win32") {
    return join(process.env.LOCALAPPDATA || join(homedir(), "AppData", "Local"), "OpenPlugin");
  }
  return join(homedir(), ".openplugin");
}

export type CompanionConfig = { token: string };

export function ensureConfig(): CompanionConfig {
  const dir = dataDir();
  mkdirSync(dir, { recursive: true });
  const file = join(dir, "companion.json");
  if (existsSync(file)) {
    return JSON.parse(readFileSync(file, "utf8")) as CompanionConfig;
  }
  const cfg = { token: randomBytes(24).toString("hex") };
  writeFileSync(file, JSON.stringify(cfg, null, 2), { mode: 0o600 });
  return cfg;
}

export function readPolicy(): Policy | null {
  const file = join(dataDir(), "policy.json");
  if (!existsSync(file)) return null;
  return JSON.parse(readFileSync(file, "utf8")) as Policy;
}

export function appendAudit(entry: Record<string, unknown>): void {
  const file = join(dataDir(), "audit.jsonl");
  mkdirSync(dirname(file), { recursive: true });
  appendFileSync(file, `${JSON.stringify(entry)}\n`);
}

export function readAudit(limit = 100): Record<string, unknown>[] {
  const file = join(dataDir(), "audit.jsonl");
  if (!existsSync(file)) return [];
  const lines = readFileSync(file, "utf8").trim().split(/\n/).filter(Boolean);
  return lines
    .slice(-limit)
    .reverse()
    .map((line) => JSON.parse(line) as Record<string, unknown>);
}
