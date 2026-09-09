import { screenshotPath, withTaskpane } from "./session.js";
import type { LiveCommand } from "./parse.js";

function testIdSelector(id: string): string {
  return `[data-testid="${id}"]`;
}

export async function runUi(cmd: LiveCommand): Promise<void> {
  const sub = cmd.sub;
  if (sub === "snapshot") {
    const snap = await withTaskpane(async (page) => {
      const aria = await page.locator("body").ariaSnapshot();
      const url = page.url();
      const title = await page.title();
      return { url, title, aria };
    });
    console.log(JSON.stringify(snap, null, 2));
    return;
  }
  if (sub === "screenshot") {
    const dest = cmd.rest[0] || (await screenshotPath());
    await withTaskpane(async (page) => {
      await page.screenshot({ path: dest, fullPage: true });
    });
    console.log(dest);
    return;
  }
  if (sub === "click") {
    const id = cmd.rest[0];
    if (!id) throw new Error("Usage: live <host> ui click <testid>");
    await withTaskpane(async (page) => {
      await page.locator(testIdSelector(id)).first().click();
    });
    console.log(`clicked ${id}`);
    return;
  }
  if (sub === "fill") {
    const id = cmd.rest[0];
    const text = cmd.rest.slice(1).join(" ");
    if (!id) throw new Error("Usage: live <host> ui fill <testid> <text>");
    await withTaskpane(async (page) => {
      const root = page.locator(testIdSelector(id)).first();
      const input = root.locator("textarea, input").first();
      if (await input.count()) await input.fill(text);
      else await root.fill(text);
    });
    console.log(`filled ${id}`);
    return;
  }
  if (sub === "press") {
    const key = cmd.rest[0];
    if (!key) throw new Error("Usage: live <host> ui press <key>");
    await withTaskpane(async (page) => {
      await page.keyboard.press(key);
    });
    console.log(`pressed ${key}`);
    return;
  }
  throw new Error(`Unknown ui command '${sub}'. Use snapshot|screenshot|click|fill|press`);
}

export async function runConsole(): Promise<void> {
  const messages = await withTaskpane(async (page) => {
    const errors: string[] = [];
    page.on("pageerror", (err) => errors.push(err.message));
    page.on("console", (msg) => {
      if (msg.type() === "error") errors.push(msg.text());
    });
    await page.waitForTimeout(500);
    return errors;
  });
  console.log(JSON.stringify(messages, null, 2));
}
