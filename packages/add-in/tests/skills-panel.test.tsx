import { parseSkillMarkdown, serializeSkill } from "@openplugin/core";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { fireEvent, render, screen } from "@testing-library/react";
import { type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import type { ManagedSkill } from "../src/skill-store";
import { SkillsPanel } from "../src/ui/SkillsPanel";

function wrap(ui: ReactNode) {
  return <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>;
}

function managed(name: string, origin: ManagedSkill["origin"], overridesBundled = false): ManagedSkill {
  const skill = parseSkillMarkdown(
    `---
name: ${name}
description: A test skill named ${name} used when testing the panel.
metadata:
  openplugin/hosts: excel
---
Body
`,
    name
  );
  return {
    ...skill,
    markdown: serializeSkill(skill),
    origin,
    updatedAt: "",
    overridesBundled
  };
}

describe("SkillsPanel", () => {
  it("offers Remove for imported skills and Customize for bundled skills", () => {
    const onRemove = vi.fn(async () => undefined);
    const onOpenEditor = vi.fn();
    render(
      wrap(
        <SkillsPanel
          skills={[managed("range-cleanup", "imported"), managed("selection-rewrite", "bundled")]}
          disabledSkills={[]}
          editor={null}
          onClose={() => undefined}
          onToggle={() => undefined}
          onRemove={onRemove}
          onExport={() => undefined}
          onOpenEditor={onOpenEditor}
          onChangeEditor={() => undefined}
          onSaveEditor={() => undefined}
          onCloseEditor={() => undefined}
          onImportUrl={async () => undefined}
          onImportMarkdown={async () => undefined}
          onNew={() => undefined}
        />
      )
    );

    fireEvent.click(screen.getByRole("button", { name: "Actions for range-cleanup" }));
    fireEvent.click(screen.getByRole("menuitem", { name: "Remove" }));
    fireEvent.click(screen.getByRole("button", { name: "Remove" }));
    expect(onRemove).toHaveBeenCalledWith("range-cleanup");

    fireEvent.click(screen.getByRole("button", { name: "Actions for selection-rewrite" }));
    expect(screen.queryByRole("menuitem", { name: "Remove" })).toBeNull();
    fireEvent.click(screen.getByRole("menuitem", { name: "Customize" }));
    expect(onOpenEditor).toHaveBeenCalled();
    expect(onOpenEditor.mock.calls[0]?.[0]).toMatchObject({ mode: "edit" });
  });
});
