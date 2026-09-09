import {
  buildExtractCall,
  buildMapCall,
  buildPromptCall,
  cacheKey,
  chatCompletions,
  chunkRows,
  parseExtract,
  parseMappedArray,
  type ProviderConfig
} from "@openplugin/core";

const cache = new Map<string, string | string[]>();

async function loadConfig(): Promise<ProviderConfig> {
  try {
    const raw = await OfficeRuntime.storage.getItem("openplugin.provider");
    if (raw) return JSON.parse(raw) as ProviderConfig;
  } catch {
    /* ignore */
  }
  throw new Error("OpenPlugin is not configured. Open the task pane and save Settings.");
}

async function complete(config: ProviderConfig, system: string, user: string): Promise<string> {
  const result = await chatCompletions({
    config: { ...config, maxOutputTokens: config.maxOutputTokens ?? 512 },
    messages: [
      { role: "system", content: system },
      { role: "user", content: user }
    ],
    stream: false
  });
  return result.message.role === "assistant" ? result.message.content || "" : "";
}

async function prompt(promptText: string, range?: unknown[][]): Promise<string> {
  const config = await loadConfig();
  const key = cacheKey(["prompt", config.model, promptText, range ?? null]);
  const hit = cache.get(key);
  if (typeof hit === "string") return hit;
  const { system, user } = buildPromptCall(String(promptText), range);
  const out = await complete(config, system, user);
  cache.set(key, out);
  return out;
}

async function map(range: unknown[][], instruction: string): Promise<string[][]> {
  const config = await loadConfig();
  const key = cacheKey(["map", config.model, instruction, range]);
  const hit = cache.get(key);
  if (Array.isArray(hit)) return hit.map((v) => [v]);
  const rows = range ?? [];
  const out: string[] = [];
  for (const chunk of chunkRows(rows, 25)) {
    const { system, user } = buildMapCall(String(instruction), chunk);
    const text = await complete(config, system, user);
    out.push(...parseMappedArray(text, chunk.length));
  }
  cache.set(key, out);
  return out.map((v) => [v]);
}

async function extract(range: unknown[][], schema: string): Promise<unknown[][]> {
  const config = await loadConfig();
  const { system, user } = buildExtractCall(String(schema), range);
  const text = await complete(config, system, user);
  return [parseExtract(text, String(schema))];
}

function associate(): void {
  const cf = (globalThis as { CustomFunctions?: { associate: (id: string, fn: unknown) => void } }).CustomFunctions;
  if (!cf) return;
  cf.associate("PROMPT", prompt);
  cf.associate("MAP", map);
  cf.associate("EXTRACT", extract);
}

associate();
