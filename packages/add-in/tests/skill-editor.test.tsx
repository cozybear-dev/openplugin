import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { fireEvent, render, screen } from "@testing-library/react";
import { type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import { SkillEditor } from "../src/ui/SkillEditor";

function wrap(ui: ReactNode) {
  return <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>;
}

const VALID = `---
name: range-cleanup
description: Use when the user wants to tidy an Excel range.
metadata:
  openplugin/hosts: excel
---

# Cleanup

1. Read the selection.
`;

describe("SkillEditor", () => {
  it("does not save when the name is invalid", () => {
    const onSave = vi.fn();
    const invalid = VALID.replace("name: range-cleanup", "name: Range_Cleanup");
    render(
      wrap(
        <SkillEditor
          markdown={invalid}
          title="Edit skill"
          onChange={() => undefined}
          onSave={onSave}
          onCancel={() => undefined}
        />
      )
    );
    fireEvent.click(screen.getByRole("button", { name: "SKILL.md" }));
    expect(screen.getByText(/lowercase letters, numbers, and single hyphens/i)).toBeTruthy();
    expect((screen.getByRole("button", { name: "Save" }) as HTMLButtonElement).disabled).toBe(true);
    fireEvent.click(screen.getByRole("button", { name: "Save" }));
    expect(onSave).not.toHaveBeenCalled();
  });

  it("saves valid markdown", () => {
    const onSave = vi.fn();
    render(
      wrap(
        <SkillEditor
          markdown={VALID}
          title="Edit skill"
          onChange={() => undefined}
          onSave={onSave}
          onCancel={() => undefined}
        />
      )
    );
    fireEvent.click(screen.getByRole("button", { name: "Save" }));
    expect(onSave).toHaveBeenCalledTimes(1);
  });

  it("hides save in view mode", () => {
    render(
      wrap(
        <SkillEditor
          markdown={VALID}
          title="View skill"
          readOnly
          onChange={() => undefined}
          onSave={() => undefined}
          onCancel={() => undefined}
        />
      )
    );
    expect(screen.queryByRole("button", { name: "Save" })).toBeNull();
  });
});
