import type { HostAdapter } from "../hosts/types.js";
import { Changeset, type Change } from "./changeset.js";
import { rebaseExcelChanges } from "./rebase.js";

export type ApplyOpResult =
  | { change: Change; ok: true }
  | { change: Change; ok: false; error: string };

const COSMETIC_OPS = new Set(["formatRange", "resizeRange"]);

export async function applyChangeset(host: HostAdapter, changeset: Changeset): Promise<ApplyOpResult[]> {
  const results: ApplyOpResult[] = [];
  const changes = host.kind === "excel" ? rebaseExcelChanges(changeset.changes) : changeset.changes;
  for (const change of changes) {
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

export function applyFailureMessage(
  succeeded: ApplyOpResult[],
  failed: Array<Extract<ApplyOpResult, { ok: false }>>
): { kind: "error" | "warning"; text: string } {
  const detail = failed.map((r) => r.error).join("; ");
  const cosmeticOnly = failed.every((r) => COSMETIC_OPS.has(r.change.op));
  if (cosmeticOnly && succeeded.length) {
    const n = failed.length;
    return {
      kind: "warning",
      text: `Formatting skipped on ${n} range${n === 1 ? "" : "s"}: ${detail}`
    };
  }
  return {
    kind: "error",
    text: `Could not apply ${failed.length} change(s): ${detail}`
  };
}
