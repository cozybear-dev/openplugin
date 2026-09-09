import type { ApplyOpResult } from "./tools/apply.js";
import type { Change } from "./tools/changeset.js";

export type RestorePoint = {
  id: string;
  at: string;
  userLineIndex: number;
  messageCount: number;
  revisionIds: string[];
};

export type RestorableRevision = {
  id: string;
  inverse: Change[];
};

export type LineCheckpoint = {
  id: string;
  at: string;
  revisionIds: string[];
};

export function planRestore<L, M, R extends RestorableRevision>(opts: {
  lines: L[];
  messages: M[];
  restorePoint: RestorePoint;
  revisions: R[];
}): {
  lines: L[];
  messages: M[];
  remainingRevisions: R[];
  undos: R[];
  irreversible: R[];
} {
  const keepIds = new Set(opts.restorePoint.revisionIds);
  const later = opts.revisions.filter((r) => !keepIds.has(r.id));
  const undos = later.filter((r) => r.inverse.length > 0);
  const irreversible = later.filter((r) => r.inverse.length === 0);
  return {
    lines: opts.lines.slice(0, opts.restorePoint.userLineIndex + 1),
    messages: opts.messages.slice(0, opts.restorePoint.messageCount),
    remainingRevisions: opts.revisions.filter((r) => keepIds.has(r.id)),
    undos,
    irreversible
  };
}

export function countMessagesThroughUserTurn(
  messages: Array<{ role: string }>,
  userTurnIndex: number
): number {
  let seen = 0;
  for (let i = 0; i < messages.length; i++) {
    if (messages[i]?.role !== "user") continue;
    if (seen === userTurnIndex) return i + 1;
    seen += 1;
  }
  return messages.length;
}

export function resolveRestorePoint<L extends { kind: string }, M extends { role: string }>(opts: {
  lines: L[];
  messages: M[];
  restorePoints: RestorePoint[];
  userLineIndex: number;
  lineCheckpoint?: LineCheckpoint;
}): RestorePoint {
  const userTurnIndex = opts.lines
    .slice(0, opts.userLineIndex)
    .filter((line) => line.kind === "user").length;
  const messageCount = countMessagesThroughUserTurn(opts.messages, userTurnIndex);
  const stored =
    opts.lineCheckpoint ??
    [...opts.restorePoints].reverse().find((point) => point.userLineIndex === opts.userLineIndex);
  return {
    id: stored?.id ?? `synthetic-${opts.userLineIndex}`,
    at: stored?.at ?? new Date().toISOString(),
    userLineIndex: opts.userLineIndex,
    messageCount,
    revisionIds: stored?.revisionIds ?? []
  };
}

function lastUserStart<T>(items: T[], isUser: (item: T) => boolean, maxUserTurns: number): number {
  let seen = 0;
  for (let i = items.length - 1; i >= 0; i--) {
    if (!isUser(items[i]!)) continue;
    seen += 1;
    if (seen >= maxUserTurns) return i;
  }
  return 0;
}

export function truncateThread<
  L extends { kind: string },
  M extends { role: string },
  P extends RestorePoint
>(opts: {
  lines: L[];
  messages: M[];
  restorePoints: P[];
  maxUserTurns: number;
}): { lines: L[]; messages: M[]; restorePoints: P[] } {
  const lineStart = lastUserStart(opts.lines, (line) => line.kind === "user", opts.maxUserTurns);
  const messageStart = lastUserStart(opts.messages, (message) => message.role === "user", opts.maxUserTurns);
  return {
    lines: opts.lines.slice(lineStart),
    messages: opts.messages.slice(messageStart),
    restorePoints: opts.restorePoints
      .filter((point) => point.userLineIndex >= lineStart && point.messageCount >= messageStart)
      .map((point) => ({
        ...point,
        userLineIndex: point.userLineIndex - lineStart,
        messageCount: point.messageCount - messageStart
      }))
  };
}

export type RestoreUndoOutcome<R extends RestorableRevision> = {
  remainingRevisions: R[];
  errors: string[];
};

export async function applyRestoreUndos<R extends RestorableRevision>(opts: {
  undos: R[];
  remainingRevisions: R[];
  apply: (revision: R) => Promise<ApplyOpResult[]>;
}): Promise<RestoreUndoOutcome<R>> {
  const { undos, remainingRevisions, apply } = opts;
  for (let index = 0; index < undos.length; index += 1) {
    const revision = undos[index]!;
    try {
      const results = await apply(revision);
      const failed = results.filter((result) => !result.ok);
      if (!failed.length) continue;

      const detail = failed.map((result) => result.error).join("; ");
      const retryable = { ...revision, inverse: failed.map((result) => result.change) } as R;
      return {
        remainingRevisions: [...remainingRevisions, retryable, ...undos.slice(index + 1)],
        errors: [`Could not auto-undo: ${revision.id}${detail ? `: ${detail}` : ""}`]
      };
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      return {
        remainingRevisions: [...remainingRevisions, revision, ...undos.slice(index + 1)],
        errors: [`Could not auto-undo: ${revision.id}${detail ? `: ${detail}` : ""}`]
      };
    }
  }
  return { remainingRevisions, errors: [] };
}
