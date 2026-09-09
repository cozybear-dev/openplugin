import type { Policy, ProviderConfig } from "@openplugin/core";
import { DEFAULT_PROVIDER } from "./presets";

const KEY = "openplugin.provider";
const INSTRUCTIONS_KEY = "openplugin.instructions";
const POLICY_KEY = "openplugin.userPolicy";

function parseProvider(raw: unknown): ProviderConfig | null {
  if (!raw || typeof raw !== "string") return null;
  try {
    const value = JSON.parse(raw) as ProviderConfig;
    if (typeof value.baseUrl !== "string") return null;
    return {
      timeoutMs: 120000,
      maxOutputTokens: 2048,
      ...value
    };
  } catch {
    return null;
  }
}

function read(key: string): string | null {
  try {
    const roaming = Office?.context?.roamingSettings?.get(key);
    if (typeof roaming === "string" && roaming) return roaming;
  } catch {
    /* not in Office */
  }
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

async function write(key: string, raw: string): Promise<void> {
  try {
    localStorage.setItem(key, raw);
  } catch {
    /* ignore quota */
  }
  try {
    await OfficeRuntime?.storage?.setItem(key, raw);
  } catch {
    /* functions runtime only */
  }
  const settings = Office?.context?.roamingSettings;
  if (!settings) return;
  settings.set(key, raw);
  await new Promise<void>((resolve) => {
    settings.saveAsync(() => resolve());
  });
}

export function loadProvider(): ProviderConfig {
  return parseProvider(read(KEY)) ?? { ...DEFAULT_PROVIDER };
}

export async function saveProvider(config: ProviderConfig): Promise<void> {
  await write(KEY, JSON.stringify(config));
}

export function loadInstructions(): string {
  return read(INSTRUCTIONS_KEY) ?? "";
}

export async function saveInstructions(text: string): Promise<void> {
  await write(INSTRUCTIONS_KEY, text);
}

export function loadUserPolicy(): Partial<Policy> | null {
  const raw = read(POLICY_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as Partial<Policy>;
  } catch {
    return null;
  }
}
