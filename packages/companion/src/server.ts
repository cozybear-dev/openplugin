import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { appendAudit, ensureConfig, readPolicy } from "./store.ts";
import { forwardHeaders, joinTarget, proxyTo } from "./proxy.ts";

const PORT = Number(process.env.OPENPLUGIN_COMPANION_PORT ?? 8788);
const HOST = "127.0.0.1";

const cfg = ensureConfig();

function readBody(req: IncomingMessage): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (c) => chunks.push(c as Buffer));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function json(res: ServerResponse, status: number, body: unknown): void {
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Authorization, Content-Type, x-openplugin-token, x-openplugin-target, api-key, HTTP-Referer, X-Title",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS"
  });
  res.end(JSON.stringify(body));
}

function cors(res: ServerResponse): void {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "Authorization, Content-Type, x-openplugin-token, x-openplugin-target, api-key, HTTP-Referer, X-Title"
  );
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
}

function authorized(req: IncomingMessage): boolean {
  const token = req.headers["x-openplugin-token"];
  return typeof token === "string" && token === cfg.token;
}

const server = createServer(async (req, res) => {
  cors(res);
  const url = new URL(req.url || "/", `http://${HOST}:${PORT}`);
  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }
  try {
    if (req.method === "GET" && url.pathname === "/health") {
      json(res, 200, { ok: true, version: "0.1.0", paired: true, port: PORT });
      return;
    }
    if (req.method === "POST" && url.pathname === "/pair") {
      json(res, 200, { token: cfg.token });
      return;
    }
    if (req.method === "GET" && url.pathname === "/policy") {
      json(res, 200, { policy: readPolicy() });
      return;
    }
    if (req.method === "POST" && url.pathname === "/audit") {
      if (!authorized(req)) {
        json(res, 401, { error: "bad token" });
        return;
      }
      const body = JSON.parse((await readBody(req)).toString("utf8") || "{}") as Record<string, unknown>;
      appendAudit({ ts: new Date().toISOString(), ...body });
      json(res, 204, {});
      return;
    }
    if (url.pathname.startsWith("/v1/")) {
      if (!authorized(req)) {
        json(res, 401, { error: "bad token" });
        return;
      }
      const target = String(req.headers["x-openplugin-target"] || "");
      if (!target) {
        json(res, 400, { error: "missing x-openplugin-target" });
        return;
      }
      const dest = joinTarget(target, url.pathname, url.search);
      const headers = forwardHeaders(new Headers(req.headers as Record<string, string>));
      const body = req.method === "GET" ? undefined : await readBody(req);
      const upstream = await proxyTo(dest, {
        method: req.method || "GET",
        headers,
        body: body && body.length ? body : undefined
      });
      res.writeHead(upstream.status, {
        "Content-Type": upstream.headers.get("content-type") || "application/json",
        "Access-Control-Allow-Origin": "*"
      });
      const buf = Buffer.from(await upstream.arrayBuffer());
      res.end(buf);
      return;
    }
    json(res, 404, { error: "not found" });
  } catch (err) {
    json(res, 502, { error: err instanceof Error ? err.message : String(err) });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`OpenPlugin companion on http://${HOST}:${PORT}`);
});
