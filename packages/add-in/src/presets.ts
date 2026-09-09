import type { ProviderConfig } from "@openplugin/core";

export type EndpointPreset = {
  id: "openrouter" | "ollama" | "custom";
  name: string;
  baseUrl: string;
  model: string;
  needsKey: boolean;
};

export const PRESETS: EndpointPreset[] = [
  {
    id: "openrouter",
    name: "OpenRouter",
    baseUrl: "https://openrouter.ai/api/v1",
    model: "openai/gpt-4o-mini",
    needsKey: true
  },
  {
    id: "ollama",
    name: "Ollama (local)",
    baseUrl: "http://127.0.0.1:11434/v1",
    model: "llama3.1",
    needsKey: false
  },
  {
    id: "custom",
    name: "Custom OpenAI-compatible",
    baseUrl: "",
    model: "",
    needsKey: true
  }
];

export const COMPANION_ORIGIN = "http://127.0.0.1:8788";

export const DEFAULT_PROVIDER: ProviderConfig = {
  baseUrl: "",
  model: "",
  timeoutMs: 120000,
  maxOutputTokens: 2048
};

export function matchPreset(baseUrl: string): EndpointPreset {
  const normalized = baseUrl.replace(/\/+$/, "");
  const found = PRESETS.find(
    (p) => p.id !== "custom" && p.baseUrl.replace(/\/+$/, "") === normalized
  );
  return found ?? PRESETS.find((p) => p.id === "custom")!;
}

export function isOllamaUrl(baseUrl: string): boolean {
  return /127\.0\.0\.1:11434|localhost:11434/i.test(baseUrl);
}

export function isOpenRouterUrl(baseUrl: string): boolean {
  return /openrouter\.ai/i.test(baseUrl);
}
