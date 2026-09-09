import { describe, expect, it } from "vitest";
import { applyRestoreUndos, planRestore, resolveRestorePoint, truncateThread } from "../src/restore.js";

describe("planRestore", () => {
  it("truncates chat and lists later revisions newest-first", () => {
    const plan = planRestore({
      lines: [
        { id: "u1" },
        { id: "a1" },
        { id: "u2" },
        { id: "a2" }
      ],
      messages: [{ i: 0 }, { i: 1 }, { i: 2 }, { i: 3 }],
      restorePoint: {
        id: "p1",
        at: "t",
        userLineIndex: 0,
        messageCount: 1,
        revisionIds: ["r1"]
      },
      revisions: [
        { id: "r2", inverse: [{ host: "word", op: "replaceSelection", text: "old" }] },
        { id: "r1", inverse: [{ host: "word", op: "replaceSelection", text: "first" }] }
      ]
    });
    expect(plan.lines.map((l) => l.id)).toEqual(["u1"]);
    expect(plan.messages).toEqual([{ i: 0 }]);
    expect(plan.remainingRevisions.map((r) => r.id)).toEqual(["r1"]);
    expect(plan.undos.map((u) => u.id)).toEqual(["r2"]);
  });

  it("separates irreversible revisions instead of throwing", () => {
    const plan = planRestore({
      lines: [{ id: "u1" }, { id: "a1" }],
      messages: [],
      restorePoint: {
        id: "p1",
        at: "t",
        userLineIndex: 0,
        messageCount: 0,
        revisionIds: []
      },
      revisions: [
        { id: "r-chart", inverse: [] },
        { id: "r-text", inverse: [{ host: "word", op: "replaceSelection", text: "old" }] }
      ]
    });
    expect(plan.undos.map((u) => u.id)).toEqual(["r-text"]);
    expect(plan.irreversible.map((u) => u.id)).toEqual(["r-chart"]);
  });
});

describe("resolveRestorePoint", () => {
  const lines = [
    { kind: "user" },
    { kind: "assistant" },
    { kind: "user" },
    { kind: "assistant" }
  ];
  const messages = [
    { role: "user" },
    { role: "assistant" },
    { role: "user" },
    { role: "assistant" }
  ];

  it("keeps the restored user message in history even when a stored point omitted it", () => {
    const point = resolveRestorePoint({
      lines,
      messages,
      restorePoints: [
        {
          id: "p1",
          at: "t",
          userLineIndex: 0,
          messageCount: 0,
          revisionIds: ["r1"]
        }
      ],
      userLineIndex: 0
    });
    expect(point.messageCount).toBe(1);
    expect(point.revisionIds).toEqual(["r1"]);
    expect(point.userLineIndex).toBe(0);
  });

  it("synthesizes a checkpoint when none is stored so restore can still truncate", () => {
    const point = resolveRestorePoint({
      lines,
      messages,
      restorePoints: [],
      userLineIndex: 2
    });
    expect(point.userLineIndex).toBe(2);
    expect(point.messageCount).toBe(3);
    expect(point.revisionIds).toEqual([]);
    const plan = planRestore({
      lines,
      messages,
      restorePoint: point,
      revisions: []
    });
    expect(plan.lines).toEqual(lines.slice(0, 3));
    expect(plan.messages).toEqual(messages.slice(0, 3));
  });

  it("prefers checkpoint metadata stamped on the user line for revision ids", () => {
    const point = resolveRestorePoint({
      lines,
      messages,
      restorePoints: [],
      userLineIndex: 2,
      lineCheckpoint: {
        id: "stamped",
        at: "t",
        revisionIds: ["r-keep"]
      }
    });
    expect(point.id).toBe("stamped");
    expect(point.revisionIds).toEqual(["r-keep"]);
    expect(point.messageCount).toBe(3);
  });
});

describe("truncateThread", () => {
  it("drops whole user turns together instead of slicing lines and messages independently", () => {
    const truncated = truncateThread({
      lines: [
        { kind: "user", id: "u1" },
        { kind: "assistant", id: "a1" },
        { kind: "user", id: "u2" },
        { kind: "assistant", id: "a2" },
        { kind: "user", id: "u3" },
        { kind: "assistant", id: "a3" }
      ],
      messages: [
        { role: "user", id: "m1" },
        { role: "tool", id: "t1" },
        { role: "assistant", id: "m2" },
        { role: "user", id: "m3" },
        { role: "assistant", id: "m4" },
        { role: "user", id: "m5" },
        { role: "assistant", id: "m6" }
      ],
      restorePoints: [
        { id: "p1", at: "t", userLineIndex: 0, messageCount: 1, revisionIds: [] },
        { id: "p2", at: "t", userLineIndex: 2, messageCount: 4, revisionIds: [] },
        { id: "p3", at: "t", userLineIndex: 4, messageCount: 6, revisionIds: [] }
      ],
      maxUserTurns: 2
    });
    expect(truncated.lines.map((l) => l.id)).toEqual(["u2", "a2", "u3", "a3"]);
    expect(truncated.messages.map((m) => m.id)).toEqual(["m3", "m4", "m5", "m6"]);
    expect(truncated.restorePoints.map((p) => ({ id: p.id, userLineIndex: p.userLineIndex, messageCount: p.messageCount }))).toEqual([
      { id: "p2", userLineIndex: 0, messageCount: 1 },
      { id: "p3", userLineIndex: 2, messageCount: 3 }
    ]);
  });
});

describe("applyRestoreUndos", () => {
  it("keeps remaining chat truncation even when a later inverse fails", async () => {
    const outcome = await applyRestoreUndos({
      remainingRevisions: [{ id: "r1", inverse: [{ host: "word", op: "replaceSelection", text: "keep" }] }],
      undos: [
        { id: "r3", inverse: [{ host: "word", op: "replaceSelection", text: "newer" }] },
        { id: "r2", inverse: [{ host: "word", op: "replaceSelection", text: "older" }] }
      ],
      apply: async (revision) => {
        if (revision.id === "r2") {
          return [{ change: revision.inverse[0]!, ok: false, error: "blocked" }];
        }
        return revision.inverse.map((change) => ({ change, ok: true as const }));
      }
    });
    expect(outcome.errors[0]).toMatch(/r2/);
    expect(outcome.remainingRevisions.map((r) => r.id)).toEqual(["r1", "r2"]);
  });
});

