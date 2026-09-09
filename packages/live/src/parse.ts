export type LiveHost = "excel" | "word" | "powerpoint";

export type LiveAction = "start" | "stop" | "status" | "ui" | "host" | "app" | "console" | "smoke";

export type LiveCommand = {
  host: LiveHost;
  action: LiveAction;
  sub?: string;
  rest: string[];
  flags: Record<string, string | boolean>;
};

const HOSTS = new Set<LiveHost>(["excel", "word", "powerpoint"]);
const ACTIONS = new Set<LiveAction>(["start", "stop", "status", "ui", "host", "app", "console", "smoke"]);

export function parseArgs(argv: string[]): LiveCommand {
  const flags: Record<string, string | boolean> = {};
  const positional: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i] ?? "";
    if (token === "--") continue;
    if (token.startsWith("--")) {
      const key = token.slice(2);
      const next = argv[i + 1];
      if (next && !next.startsWith("-")) {
        flags[key] = next;
        i++;
      } else {
        flags[key] = true;
      }
      continue;
    }
    positional.push(token);
  }

  const host = positional[0] as LiveHost | undefined;
  if (!host || !HOSTS.has(host)) {
    throw new Error("Usage: live <excel|word|powerpoint> <start|stop|status|ui|host|app|console|smoke> ...");
  }
  const action = (positional[1] ?? "status") as LiveAction;
  if (!ACTIONS.has(action)) {
    throw new Error(`Unknown action '${action}'.`);
  }
  const rest = positional.slice(2);
  const nested = action === "ui" || action === "host" || action === "app";
  return {
    host,
    action,
    sub: nested ? rest[0] : undefined,
    rest: nested ? rest.slice(1) : rest,
    flags
  };
}
