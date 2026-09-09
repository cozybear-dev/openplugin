import { COMPANION_ORIGIN } from "./presets";

const TOKEN_KEY = "openplugin.companionToken";

export type CompanionStatus =
  | { state: "unknown" }
  | { state: "missing" }
  | { state: "connected"; token: string };

export async function probeCompanion(): Promise<CompanionStatus> {
  try {
    const health = await fetch(`${COMPANION_ORIGIN}/health`, { method: "GET" });
    if (!health.ok) return { state: "missing" };
    const pair = await fetch(`${COMPANION_ORIGIN}/pair`, { method: "POST" });
    const body = (await pair.json()) as { token?: string };
    if (!body.token) return { state: "missing" };
    try {
      localStorage.setItem(TOKEN_KEY, body.token);
    } catch {
      /* ignore */
    }
    return { state: "connected", token: body.token };
  } catch {
    return { state: "missing" };
  }
}

export async function fetchPolicy(token: string): Promise<unknown> {
  const res = await fetch(`${COMPANION_ORIGIN}/policy`, {
    headers: { "x-openplugin-token": token }
  });
  if (!res.ok) return null;
  const body = (await res.json()) as { policy?: unknown };
  return body.policy ?? null;
}

export async function postSearch(
  token: string,
  body: Record<string, unknown>
): Promise<unknown> {
  const res = await fetch(`${COMPANION_ORIGIN}/search`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-openplugin-token": token
    },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`Companion search failed (${res.status})`);
  return res.json();
}

export async function postAudit(token: string, entry: Record<string, unknown>): Promise<void> {
  await fetch(`${COMPANION_ORIGIN}/audit`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-openplugin-token": token
    },
    body: JSON.stringify(entry)
  }).catch(() => undefined);
}
