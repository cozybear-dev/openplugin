import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { MarkdownMessage } from "../src/ui/MarkdownMessage";

describe("MarkdownMessage", () => {
  it("renders emphasis, headings, lists, and fenced code", () => {
    render(
      <MarkdownMessage
        hostKind="excel"
        text={"## Title\n\nUse **bold** and `code`.\n\n- one\n- two\n\n```ts\nconst x = 1;\n```"}
      />
    );
    expect(screen.getByRole("heading", { level: 2, name: "Title" })).toBeTruthy();
    expect(screen.getByText("bold").tagName).toBe("STRONG");
    expect(screen.getByText("code").tagName).toBe("CODE");
    expect(screen.getByRole("list")).toBeTruthy();
    expect(screen.getByText("const x = 1;").closest("pre")).toBeTruthy();
  });

  it("renders GFM tables", () => {
    render(
      <MarkdownMessage
        hostKind="excel"
        text={"| A | B |\n| --- | --- |\n| 1 | 2 |"}
      />
    );
    expect(screen.getByRole("table")).toBeTruthy();
    expect(screen.getByText("1")).toBeTruthy();
  });

  it("makes prose citations clickable buttons", () => {
    render(<MarkdownMessage hostKind="excel" text="Total is in Sheet1!C4." />);
    const btn = screen.getByRole("button", { name: "Sheet1!C4" });
    expect(btn.className).toContain("op-cite");
  });

  it("does not turn code-span cell refs into citation buttons", () => {
    render(<MarkdownMessage hostKind="excel" text="The formula is `=A1+1`." />);
    expect(screen.queryByRole("button")).toBeNull();
    expect(screen.getByText("=A1+1").tagName).toBe("CODE");
  });

  it("does not execute raw HTML", () => {
    render(
      <MarkdownMessage
        hostKind="word"
        text={'<script>window.__xss=1</script>\n\n<img src=x onerror="window.__xss=1">\n\n**ok**'}
      />
    );
    expect(screen.getByText("ok").tagName).toBe("STRONG");
    expect(document.querySelector("script")).toBeNull();
    expect(document.querySelector("img")).toBeNull();
  });

  it("keeps http(s) links with a new-tab target", () => {
    render(<MarkdownMessage hostKind="word" text="See [docs](https://example.com/a)." />);
    const link = screen.getByRole("link", { name: "docs" });
    expect(link.getAttribute("href")).toBe("https://example.com/a");
    expect(link.getAttribute("target")).toBe("_blank");
    expect(link.getAttribute("rel")).toBe("noreferrer");
  });

  it("drops javascript: links and remote images", () => {
    render(
      <MarkdownMessage
        hostKind="word"
        text={"[x](javascript:alert(1))\n\n![alt](https://example.com/x.png)"}
      />
    );
    expect(document.querySelector('a[href^="javascript"]')).toBeNull();
    expect(document.querySelector("img")).toBeNull();
  });
});
