import type { ProviderConfig } from "@openplugin/core";
import { DEFAULT_PROVIDER } from "./presets";

const KEY = "openplugin.provider";

function parse(raw: unknown): ProviderConfig | null {
  if (!raw || typeof raw !== "string") return null;
  try {
    const value = JSON.parse(raw) as ProviderConfig;
    if (!value.baseUrl || !value.model) return null;
    return value;
  } catch {
    return null;
  }
}

export function loadProvider(): ProviderConfig {
  try {
    const roaming = Office?.context?.roamingSettings?.get(KEY);
    const fromRoaming = parse(roaming);
    if (fromRoaming) return fromRoaming;
  } catch {
    /* not in Office */
  }
  try {
    return parse(localStorage.getItem(KEY)) ?? { ...DEFAULT_PROVIDER };
  } catch {
    return { ...DEFAULT_PROVIDER };
  }
}

export async function saveProvider(config: ProviderConfig): Promise<void> {
  const raw = JSON.stringify(config);
  try {
    localStorage.setItem(KEY, raw);
  } catch {
    /* ignore quota */
  }
  const settings = Office?.context?.roamingSettings;
  if (!settings) return;
  settings.set(KEY, raw);
  await new Promise<void>((resolve) => {
    settings.saveAsync(() => resolve());
  });
}
