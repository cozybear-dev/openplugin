import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import type { Browser, Page } from "playwright";
import { connectTaskpane, disconnectCdp } from "./cdp.js";
import { outputDir } from "./paths.js";

export async function withTaskpane<T>(fn: (page: Page) => Promise<T>): Promise<T> {
  const { browser, page } = await connectTaskpane();
  try {
    await page.waitForSelector("[data-testid='openplugin-app']", { timeout: 30_000 }).catch(() => undefined);
    return await fn(page);
  } finally {
    await disconnectCdp(browser);
  }
}

export async function screenshotPath(name = "taskpane"): Promise<string> {
  await mkdir(outputDir(), { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return join(outputDir(), `${name}-${stamp}.png`);
}

export type { Browser, Page };
