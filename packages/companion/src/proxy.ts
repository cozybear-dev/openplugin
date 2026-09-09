const HOP = new Set([
  "connection",
  "keep-alive",
  "proxy-authenticate",
  "proxy-authorization",
  "te",
  "trailers",
  "transfer-encoding",
  "upgrade",
  "host",
  "content-length"
]);

export function joinTarget(baseUrl: string, path: string, search = ""): string {
  const base = baseUrl.replace(/\/+$/, "");
  if (path.startsWith("/v1/") && /\/v1$/i.test(base) && path.startsWith("/v1")) {
    return `${base}${path.slice(3)}${search}`;
  }
  return `${base}${path}${search}`;
}

export function forwardHeaders(input: Headers): Headers {
  const out = new Headers();
  input.forEach((value, key) => {
    if (HOP.has(key.toLowerCase())) return;
    if (key.toLowerCase() === "x-openplugin-token") return;
    if (key.toLowerCase() === "x-openplugin-target") return;
    out.set(key, value);
  });
  return out;
}

export async function proxyTo(
  targetUrl: string,
  init: { method: string; headers: Headers; body?: ArrayBuffer | string | null; signal?: AbortSignal },
  fetchImpl: typeof fetch = fetch
): Promise<Response> {
  return fetchImpl(targetUrl, {
    method: init.method,
    headers: init.headers,
    body: init.method === "GET" || init.method === "HEAD" ? undefined : init.body ?? undefined,
    signal: init.signal
  });
}
