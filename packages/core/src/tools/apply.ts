import type { HostAdapter } from "../hosts/types.js";
import { Changeset, type Change } from "./changeset.js";

export type ApplyOpResult =
  | { change: Change; ok: true }
  | { change: Change; ok: false; error: string };

export async function applyChangeset(host: HostAdapter, changeset: Changeset): Promise<ApplyOpResult[]> {
  const results: ApplyOpResult[] = [];
  for (const change of changeset.changes) {
    const one = new Changeset();
    one.add(change);
    try {
      await host.apply(one);
      results.push({ change, ok: true });
    } catch (err) {
      results.push({
        change,
        ok: false,
        error: err instanceof Error ? err.message : String(err)
      });
    }
  }
  return results;
}
