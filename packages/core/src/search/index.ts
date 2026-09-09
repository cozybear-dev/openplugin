import type { NativeSearchToolDefinition, ToolDefinition } from "../llm/types.js";

export type SearchBackend = "auto" | "native" | "duckduckgo" | "exa" | "custom";
export type ResolvedSearchBackend = Exclude<SearchBackend, "auto">;

export type SearchHit = {
  title: string;
  url: string;
  snippet: string;
};

export type SearchResult =
  | { ok: true; backend: ResolvedSearchBackend; results: SearchHit[] }
  | { ok: false; backend: ResolvedSearchBackend; error: string; rateLimited?: boolean };

const RATE_LIMIT_HINT =
  "Free web search is rate-limited. Add an Exa API key in Settings, or switch the search backend.";

export function supportsNativeSearch(baseUrl: string): boolean {
  return /openrouter\.ai/i.test(baseUrl) || /api\.openai\.com/i.test(baseUrl);
}

export function resolveSearchBackend(opts: {
  backend: SearchBackend;
  baseUrl: string;
  exaApiKey?: string;
}): ResolvedSearchBackend {
  if (opts.backend !== "auto") return opts.backend;
  if (supportsNativeSearch(opts.baseUrl)) return "native";
  if (opts.exaApiKey) return "exa";
  return "duckduckgo";
}

export function nativeSearchTool(baseUrl: string): NativeSearchToolDefinition | undefined {
  if (/openrouter\.ai/i.test(baseUrl)) return { type: "openrouter:web_search" };
  if (/api\.openai\.com/i.test(baseUrl)) return { type: "web_search" };
  return undefined;
}

export const WEB_SEARCH_TOOLS: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "web.search",
      description: "Search the public web. Use for current facts, citations, and URLs.",
      parameters: {
        type: "object",
        properties: { query: { type: "string" } },
        required: ["query"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "web.fetch",
      description: "Fetch a URL and return extracted text.",
      parameters: {
        type: "object",
        properties: { url: { type: "string" } },
        required: ["url"]
      }
    }
  }
];

export type SearchProxy = (req: {
  action: "search" | "fetch";
  backend: ResolvedSearchBackend;
  query?: string;
  url?: string;
}) => Promise<SearchResult>;

export async function runWebSearch(opts: {
  backend: ResolvedSearchBackend;
  query: string;
  apiKey?: string;
  customUrl?: string;
  extraHeaders?: Record<string, string>;
  fetchImpl?: typeof fetch;
  proxy?: SearchProxy;
}): Promise<SearchResult> {
  const fetchImpl = opts.fetchImpl ?? globalThis.fetch;
  if (opts.backend === "native") {
    return { ok: false, backend: "native", error: "Native search is executed by the model provider." };
  }
  const direct =
    opts.backend === "exa"
      ? await searchExa(opts.query, opts.apiKey ?? "", fetchImpl)
      : opts.backend === "custom"
        ? await searchCustom(opts.query, opts.customUrl ?? "", opts.apiKey, opts.extraHeaders, fetchImpl)
        : await searchDuckDuckGo(opts.query, fetchImpl);
  if (direct.ok || direct.rateLimited || !opts.proxy) return direct;
  try {
    return await opts.proxy({ action: "search", backend: opts.backend, query: opts.query });
  } catch {
    return direct;
  }
}

export async function runWebFetch(opts: {
  backend: ResolvedSearchBackend;
  url: string;
  apiKey?: string;
  fetchImpl?: typeof fetch;
  proxy?: SearchProxy;
}): Promise<SearchResult> {
  const fetchImpl = opts.fetchImpl ?? globalThis.fetch;
  if (opts.backend === "exa" && opts.apiKey) return fetchExa(opts.url, opts.apiKey, fetchImpl);
  try {
    const res = await fetchImpl(opts.url, { headers: { Accept: "text/html,application/json,text/plain" } });
    if (isRateLimited(res)) {
      return { ok: false, backend: opts.backend, error: RATE_LIMIT_HINT, rateLimited: true };
    }
    if (!res.ok) return { ok: false, backend: opts.backend, error: `Fetch failed (${res.status}).` };
    const text = stripHtml(await res.text()).slice(0, 8000);
    return { ok: true, backend: opts.backend, results: [{ title: opts.url, url: opts.url, snippet: text }] };
  } catch (err) {
    if (opts.proxy) {
      try {
        return await opts.proxy({ action: "fetch", backend: opts.backend, url: opts.url });
      } catch {
        /* fall through */
      }
    }
    return { ok: false, backend: opts.backend, error: err instanceof Error ? err.message : String(err) };
  }
}

async function searchDuckDuckGo(query: string, fetchImpl: typeof fetch): Promise<SearchResult> {
  const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
  try {
    const res = await fetchImpl(url, { headers: { "User-Agent": "OpenPlugin/0.1" } });
    if (isRateLimited(res)) {
      return { ok: false, backend: "duckduckgo", error: RATE_LIMIT_HINT, rateLimited: true };
    }
    if (!res.ok) {
      if (res.status === 403 || res.status === 429) {
        return { ok: false, backend: "duckduckgo", error: RATE_LIMIT_HINT, rateLimited: true };
      }
      return { ok: false, backend: "duckduckgo", error: `Search failed (${res.status}).` };
    }
    const html = await res.text();
    if (/captcha|anomaly-modal|bots/i.test(html) && !/result__a/i.test(html)) {
      return { ok: false, backend: "duckduckgo", error: RATE_LIMIT_HINT, rateLimited: true };
    }
    const results = parseDuckDuckGo(html).slice(0, 5);
    if (!results.length) {
      return { ok: false, backend: "duckduckgo", error: RATE_LIMIT_HINT, rateLimited: true };
    }
    return { ok: true, backend: "duckduckgo", results };
  } catch (err) {
    return { ok: false, backend: "duckduckgo", error: err instanceof Error ? err.message : String(err) };
  }
}

export function parseDuckDuckGo(html: string): SearchHit[] {
  const hits: SearchHit[] = [];
  const re =
    /<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?(?:class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html))) {
    hits.push({
      url: decodeDuckHref(m[1]),
      title: stripHtml(m[2]),
      snippet: stripHtml(m[3] ?? "")
    });
  }
  return hits;
}

function decodeDuckHref(href: string): string {
  try {
    const u = new URL(href, "https://duckduckgo.com");
    return u.searchParams.get("uddg") ?? href;
  } catch {
    return href;
  }
}

async function searchExa(query: string, apiKey: string, fetchImpl: typeof fetch): Promise<SearchResult> {
  if (!apiKey) return { ok: false, backend: "exa", error: "Exa API key is missing. Add it in Settings." };
  try {
    const res = await fetchImpl("https://api.exa.ai/search", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "x-api-key": apiKey
      },
      body: JSON.stringify({
        query,
        numResults: 5,
        contents: { highlights: true }
      })
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      return { ok: false, backend: "exa", error: `Exa search failed (${res.status}): ${body.slice(0, 200)}` };
    }
    const json = (await res.json()) as {
      results?: Array<{ title?: string; url?: string; highlights?: string[]; text?: string }>;
    };
    const results = (json.results ?? []).map((r) => ({
      title: r.title ?? r.url ?? "Result",
      url: r.url ?? "",
      snippet: (r.highlights ?? []).join(" ") || (r.text ?? "").slice(0, 400)
    }));
    return { ok: true, backend: "exa", results };
  } catch (err) {
    return { ok: false, backend: "exa", error: err instanceof Error ? err.message : String(err) };
  }
}

async function fetchExa(url: string, apiKey: string, fetchImpl: typeof fetch): Promise<SearchResult> {
  try {
    const res = await fetchImpl("https://api.exa.ai/contents", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "x-api-key": apiKey
      },
      body: JSON.stringify({ urls: [url], text: true })
    });
    if (!res.ok) return { ok: false, backend: "exa", error: `Exa fetch failed (${res.status}).` };
    const json = (await res.json()) as { results?: Array<{ title?: string; url?: string; text?: string }> };
    const r = json.results?.[0];
    return {
      ok: true,
      backend: "exa",
      results: [{ title: r?.title ?? url, url: r?.url ?? url, snippet: (r?.text ?? "").slice(0, 8000) }]
    };
  } catch (err) {
    return { ok: false, backend: "exa", error: err instanceof Error ? err.message : String(err) };
  }
}

async function searchCustom(
  query: string,
  url: string,
  apiKey: string | undefined,
  extraHeaders: Record<string, string> | undefined,
  fetchImpl: typeof fetch
): Promise<SearchResult> {
  if (!url) return { ok: false, backend: "custom", error: "Custom search URL is missing." };
  try {
    const headers: Record<string, string> = { "Content-Type": "application/json", ...(extraHeaders ?? {}) };
    if (apiKey && !headers.Authorization) headers.Authorization = `Bearer ${apiKey}`;
    const res = await fetchImpl(url, {
      method: "POST",
      headers,
      body: JSON.stringify({ query, maxResults: 5 })
    });
    if (isRateLimited(res)) {
      return { ok: false, backend: "custom", error: RATE_LIMIT_HINT, rateLimited: true };
    }
    if (!res.ok) return { ok: false, backend: "custom", error: `Custom search failed (${res.status}).` };
    const json = (await res.json()) as { results?: SearchHit[] };
    return { ok: true, backend: "custom", results: json.results ?? [] };
  } catch (err) {
    return { ok: false, backend: "custom", error: err instanceof Error ? err.message : String(err) };
  }
}

function isRateLimited(res: Response): boolean {
  return res.status === 429 || res.status === 403;
}

function stripHtml(value: string): string {
  return value
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}
