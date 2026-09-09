import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { completionForMessages, jsonBody, sseBody, MOCK_LLM_PORT } from "./mock-llm.js";

function cors(res: ServerResponse): void {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type, api-key, HTTP-Referer, X-Title");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
}

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (c) => chunks.push(c as Buffer));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

export function startMockLlm(port = MOCK_LLM_PORT): Promise<{ close: () => Promise<void>; port: number }> {
  const server = createServer(async (req, res) => {
    cors(res);
    const url = new URL(req.url || "/", `http://127.0.0.1:${port}`);
    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }
    try {
      if (req.method === "GET" && url.pathname.endsWith("/models")) {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ data: [{ id: "openplugin-live", object: "model" }] }));
        return;
      }
      if (req.method === "POST" && /chat\/completions/.test(url.pathname)) {
        const raw = await readBody(req);
        const body = JSON.parse(raw || "{}") as { messages?: unknown[]; stream?: boolean; model?: string };
        const completion = completionForMessages((body.messages ?? []) as Parameters<typeof completionForMessages>[0]);
        if (body.stream === false) {
          res.writeHead(200, { "Content-Type": "application/json" });
          res.end(jsonBody(completion, body.model));
          return;
        }
        res.writeHead(200, { "Content-Type": "text/event-stream" });
        res.end(sseBody(completion));
        return;
      }
      res.writeHead(404, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: { message: "not found" } }));
    } catch (err) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: { message: err instanceof Error ? err.message : String(err) } }));
    }
  });
  return new Promise((resolve, reject) => {
    server.listen(port, "127.0.0.1", () =>
      resolve({
        port,
        close: () =>
          new Promise((done, fail) => server.close((e) => (e ? fail(e) : done())))
      })
    );
    server.on("error", reject);
  });
}

export async function ensureMockLlm(port = MOCK_LLM_PORT): Promise<void> {
  try {
    const res = await fetch(`http://127.0.0.1:${port}/v1/models`);
    if (res.ok) return;
  } catch {
    /* start */
  }
  await startMockLlm(port);
}
