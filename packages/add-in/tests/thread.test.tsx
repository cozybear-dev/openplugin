import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { fireEvent, render } from "@testing-library/react";
import { type ReactNode } from "react";
import { describe, expect, it } from "vitest";
import type { StoredLine } from "../src/history";
import { Thread } from "../src/ui/Thread";

function wrap(ui: ReactNode) {
  return <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>;
}

const user: StoredLine = { kind: "user", text: "hello" };
const assistant: StoredLine = { kind: "assistant", text: "working on it" };
const more: StoredLine = { kind: "activity", phase: "start", activity: "read", name: "excel.readRange", label: "Reading A1" };

function scroller(container: HTMLElement): HTMLDivElement {
  const el = container.querySelector(".op-thread");
  if (!(el instanceof HTMLDivElement)) throw new Error("missing .op-thread");
  return el;
}

function mockLayout(el: HTMLElement, scrollHeight: number, clientHeight: number): void {
  Object.defineProperty(el, "scrollHeight", { configurable: true, get: () => scrollHeight });
  Object.defineProperty(el, "clientHeight", { configurable: true, get: () => clientHeight });
}

describe("Thread auto-scroll", () => {
  it("scrolls to the bottom when new lines arrive while pinned", () => {
    const { container, rerender } = render(wrap(<Thread hostKind="excel" lines={[user]} busy={false} />));
    const el = scroller(container);
    mockLayout(el, 1000, 400);
    rerender(wrap(<Thread hostKind="excel" lines={[user, assistant]} busy={false} />));
    expect(el.scrollTop).toBe(1000);
  });

  it("does not steal scroll after the user moves up", () => {
    const { container, rerender } = render(wrap(<Thread hostKind="excel" lines={[user]} busy={false} />));
    const el = scroller(container);
    mockLayout(el, 1000, 400);
    el.scrollTop = 0;
    fireEvent.scroll(el);
    rerender(wrap(<Thread hostKind="excel" lines={[user, assistant, more]} busy={false} />));
    expect(el.scrollTop).toBe(0);
  });

  it("re-pins when a new user message is appended", () => {
    const { container, rerender } = render(wrap(<Thread hostKind="excel" lines={[user, assistant]} busy={false} />));
    const el = scroller(container);
    mockLayout(el, 1000, 400);
    el.scrollTop = 0;
    fireEvent.scroll(el);
    const next: StoredLine = { kind: "user", text: "do it again" };
    rerender(wrap(<Thread hostKind="excel" lines={[user, assistant, next]} busy={false} />));
    expect(el.scrollTop).toBe(1000);
  });
});
