import { parseArgs } from "./parse.js";
import { startLive, statusLive, stopLive } from "./launch.js";
import { runUi, runConsole } from "./ui.js";
import { runHost } from "./host.js";
import { runApp } from "./app.js";
import { runSmoke } from "./smoke.js";

async function main(argv: string[]): Promise<void> {
  const cmd = parseArgs(argv);
  switch (cmd.action) {
    case "start":
      await startLive(cmd.host, cmd.flags);
      return;
    case "stop":
      await stopLive(cmd.host, cmd.flags);
      return;
    case "status":
      await statusLive();
      return;
    case "ui":
      await runUi(cmd);
      return;
    case "host":
      await runHost(cmd);
      return;
    case "app":
      await runApp(cmd);
      return;
    case "console":
      await runConsole();
      return;
    case "smoke":
      await runSmoke(cmd.host, cmd.flags);
      return;
  }
}

const args = process.argv.slice(2);
main(args)
  .then(() => {
    // Playwright's CDP socket keeps the event loop alive. Do not browser.close()
    // (that kills Office). Force the one-shot CLI to exit instead.
    process.exit(0);
  })
  .catch((err) => {
    console.error(err instanceof Error ? err.stack ?? err.message : err);
    process.exit(1);
  });
