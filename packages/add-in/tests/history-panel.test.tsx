import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { fireEvent, render, screen } from "@testing-library/react";
import { type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import { emptyConversation } from "../src/conversation-store";
import { HistoryPanel } from "../src/ui/HistoryPanel";

function wrap(ui: ReactNode) {
  return <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>;
}

describe("HistoryPanel", () => {
  it("selects a chat and confirms delete", () => {
    const onSelect = vi.fn();
    const onDelete = vi.fn();
    const conversations = [
      {
        ...emptyConversation("excel"),
        id: "a",
        title: "Clean column A",
        lines: [{ kind: "user" as const, text: "Clean column A" }]
      },
      { ...emptyConversation("excel"), id: "b", title: "New chat" }
    ];
    render(
      wrap(
        <HistoryPanel
          conversations={conversations}
          activeId="b"
          error={null}
          onClose={() => undefined}
          onSelect={onSelect}
          onRename={() => undefined}
          onDelete={onDelete}
        />
      )
    );
    fireEvent.click(screen.getByRole("button", { name: "Clean column A" }));
    expect(onSelect).toHaveBeenCalledWith("a");
    fireEvent.click(screen.getByRole("button", { name: "Actions for Clean column A" }));
    fireEvent.click(screen.getByRole("menuitem", { name: "Delete" }));
    fireEvent.click(screen.getByRole("button", { name: "Delete" }));
    expect(onDelete).toHaveBeenCalledWith("a");
  });

  it("shows save error copy", () => {
    render(
      wrap(
        <HistoryPanel
          conversations={[]}
          activeId="x"
          error="Chats could not be saved."
          onClose={() => undefined}
          onSelect={() => undefined}
          onRename={() => undefined}
          onDelete={() => undefined}
        />
      )
    );
    expect(screen.getByText("Chats could not be saved.")).toBeTruthy();
    expect(screen.getByText("No saved chats yet")).toBeTruthy();
  });
});
