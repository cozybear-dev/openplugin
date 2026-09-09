import { describe, expect, it } from "vitest";
import { isPinnedToBottom, scrollToBottom } from "../src/ui/scroll";

describe("isPinnedToBottom", () => {
  it("is true when flush with the bottom", () => {
    expect(isPinnedToBottom({ scrollTop: 600, scrollHeight: 1000, clientHeight: 400 })).toBe(true);
  });

  it("is true within the pin threshold", () => {
    expect(isPinnedToBottom({ scrollTop: 550, scrollHeight: 1000, clientHeight: 400 })).toBe(true);
  });

  it("is false when scrolled up past the threshold", () => {
    expect(isPinnedToBottom({ scrollTop: 0, scrollHeight: 1000, clientHeight: 400 })).toBe(false);
  });
});

describe("scrollToBottom", () => {
  it("sets scrollTop to scrollHeight", () => {
    const el = { scrollTop: 12, scrollHeight: 800 };
    scrollToBottom(el);
    expect(el.scrollTop).toBe(800);
  });
});
