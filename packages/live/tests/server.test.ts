import { afterAll, describe, expect, it } from "vitest";
import { startMockLlm } from "../src/server.js";

describe("mock LLM HTTP", () => {
  let close: (() => Promise<void>) | undefined;
  const port = 18790;

  afterAll(async () => {
    await close?.();
  });

  it("serves models and a non-stream completion", async () => {
    const server = await startMockLlm(port);
    close = server.close;
    const models = await fetch(`http://127.0.0.1:${port}/v1/models`);
    expect(models.ok).toBe(true);
    const listed = (await models.json()) as { data: Array<{ id: string }> };
    expect(listed.data[0]?.id).toBe("openplugin-live");

    const chat = await fetch(`http://127.0.0.1:${port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        stream: false,
        messages: [{ role: "user", content: "Reply with ok" }]
      })
    });
    expect(chat.ok).toBe(true);
    const json = (await chat.json()) as { choices: Array<{ message: { content: string } }> };
    expect(json.choices[0]?.message.content).toBe("ok");
  });
});
