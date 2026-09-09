import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { fireEvent, render, screen } from "@testing-library/react";
import { type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import { Header } from "../src/ui/Header";

function wrap(ui: ReactNode) {
  return <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>;
}

describe("Header", () => {
  it("exposes History and New chat", () => {
    const onOpenHistory = vi.fn();
    const onNewChat = vi.fn();
    render(
      wrap(
        <Header
          hostKind="excel"
          contextLabel="Sheet1!A1"
          onNewChat={onNewChat}
          onOpenHistory={onOpenHistory}
          onOpenSkills={() => undefined}
          onOpenSettings={() => undefined}
        />
      )
    );
    fireEvent.click(screen.getByRole("button", { name: "History" }));
    expect(onOpenHistory).toHaveBeenCalled();
    fireEvent.click(screen.getByRole("button", { name: "New chat" }));
    expect(onNewChat).toHaveBeenCalled();
    expect(screen.getByTestId("header-history")).toBeTruthy();
    expect(screen.getByTestId("header-settings")).toBeTruthy();
  });
});
