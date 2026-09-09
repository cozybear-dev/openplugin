import type { ProviderConfig } from "@openplugin/core";

export type EndpointPreset = {
  id: string;
  name: string;
  baseUrl: string;
  model: string;
};

export const PRESETS: EndpointPreset[] = [
  { id: "xai", name: "SpaceXAI (xAI)", baseUrl: "https://api.x.ai/v1", model: "grok-4.5" },
  { id: "groq", name: "Groq", baseUrl: "https://api.groq.com/openai/v1", model: "llama-3.3-70b-versatile" },
  { id: "openrouter", name: "OpenRouter", baseUrl: "https://openrouter.ai/api/v1", model: "openai/gpt-4o-mini" },
  { id: "ollama", name: "Ollama (local)", baseUrl: "http://127.0.0.1:11434/v1", model: "llama3.1" },
  { id: "custom", name: "Custom OpenAI-compatible", baseUrl: "https://", model: "" }
];

export const DEFAULT_PROVIDER: ProviderConfig = {
  baseUrl: PRESETS[0].baseUrl,
  model: PRESETS[0].model,
  timeoutMs: 120000,
  maxOutputTokens: 2048
};
