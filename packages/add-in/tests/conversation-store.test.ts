import { describe, expect, it } from "vitest";
import type { StoredLine } from "../src/history";
import {
  applyAutoTitle,
  autoTitleFromLines,
  capConversations,
  conversationHasContent,
  emptyConversation,
  visibleConversations,
  type ConversationRecord
} from "../src/conversation-store";

const user = (text: string): StoredLine => ({ kind: "user", text });

function rec(id: string, extra: Partial<ConversationRecord> = {}): ConversationRecord {
  return {
    ...emptyConversation("excel", "2026-01-01T00:00:00.000Z"),
    id,
    updatedAt: extra.updatedAt ?? "2026-01-01T00:00:00.000Z",
    ...extra
  };
}

describe("autoTitleFromLines", () => {
  it("uses the first line of the first user message and ellipsizes at 60", () => {
    expect(autoTitleFromLines([])).toBe("New chat");
    expect(autoTitleFromLines([user("Clean column A\nmore")])).toBe("Clean column A");
    expect(autoTitleFromLines([user("x".repeat(70))]).length).toBe(60);
    expect(autoTitleFromLines([user("x".repeat(70))]).endsWith("…")).toBe(true);
  });
});

describe("applyAutoTitle", () => {
  it("does not overwrite a user title", () => {
    const named = rec("1", { title: "Keep me", titleSource: "user", lines: [user("other")] });
    expect(applyAutoTitle(named, named.lines).title).toBe("Keep me");
  });
});

describe("cap and visible", () => {
  it("keeps the active conversation when capping", () => {
    const list = Array.from({ length: 45 }, (_, i) =>
      rec(`c${i}`, { updatedAt: `2026-01-${String((i % 28) + 1).padStart(2, "0")}T00:00:00.000Z` })
    );
    const capped = capConversations(list, "c0", 40);
    expect(capped).toHaveLength(40);
    expect(capped.some((c) => c.id === "c0")).toBe(true);
  });

  it("hides empty threads except the active one", () => {
    const empty = rec("e");
    const full = rec("f", { lines: [user("hi")], updatedAt: "2026-02-01T00:00:00.000Z" });
    expect(visibleConversations([empty, full], "e").map((c) => c.id)).toEqual(["f", "e"]);
    expect(visibleConversations([empty, full], "f").map((c) => c.id)).toEqual(["f"]);
  });
});

describe("conversationHasContent", () => {
  it("is false for only activity lines", () => {
    expect(conversationHasContent([])).toBe(false);
    expect(conversationHasContent([{ kind: "error", text: "x" }])).toBe(false);
    expect(conversationHasContent([user("hi")])).toBe(true);
  });
});
