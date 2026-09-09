import {
  Button,
  Caption1,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Textarea
} from "@fluentui/react-components";
import { MoreHorizontal24Regular } from "@fluentui/react-icons";
import { parseSkillMarkdown, serializeSkill, uniqueSkillName } from "@openplugin/core";
import { useMemo, useState } from "react";
import type { ManagedSkill } from "../skill-store";
import { SkillEditor, type SkillEditorMode } from "./SkillEditor";

export type SkillEditorState = {
  markdown: string;
  previousName?: string;
  mode: SkillEditorMode;
  generating: boolean;
  error: string | null;
};

export function SkillsPanel(props: {
  skills: ManagedSkill[];
  disabledSkills: string[];
  editor: SkillEditorState | null;
  onClose: () => void;
  onToggle: (name: string, enabled: boolean) => void;
  onRemove: (name: string) => Promise<void>;
  onExport: (skill: ManagedSkill) => void;
  onOpenEditor: (state: SkillEditorState) => void;
  onChangeEditor: (markdown: string) => void;
  onSaveEditor: () => void;
  onCloseEditor: () => void;
  onImportUrl: (url: string) => Promise<void>;
  onImportMarkdown: (md: string) => Promise<void>;
  onNew: () => void;
}) {
  const [query, setQuery] = useState("");
  const [importUrl, setImportUrl] = useState("");
  const [paste, setPaste] = useState("");
  const [importMsg, setImportMsg] = useState<string | null>(null);
  const [pendingRemove, setPendingRemove] = useState<string | null>(null);
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return props.skills;
    return props.skills.filter(
      (s) => s.name.includes(q) || s.description.toLowerCase().includes(q)
    );
  }, [props.skills, query]);

  if (props.editor) {
    return (
      <SkillEditor
        markdown={props.editor.markdown}
        title={editorTitle(props.editor)}
        generating={props.editor.generating}
        error={props.editor.error}
        readOnly={props.editor.mode === "view"}
        onChange={props.onChangeEditor}
        onSave={props.onSaveEditor}
        onCancel={props.onCloseEditor}
      />
    );
  }

  return (
    <div className="op-settings">
      <div className="op-settings-head">
        <div>
          <div className="op-eyebrow">SHORTCUTS YOU KEEP</div>
          <h1>Skills</h1>
        </div>
        <Button appearance="subtle" size="small" onClick={props.onClose}>
          Done
        </Button>
      </div>

      <div className="op-skill-toolbar">
        <Button appearance="primary" size="small" onClick={props.onNew}>
          New skill
        </Button>
      </div>

      <Field label="Filter">
        <Input value={query} onChange={(_, d) => setQuery(d.value)} placeholder="Name or description" />
      </Field>

      <ul className="op-skill-list">
        {filtered.map((s) => (
          <li key={s.name} className="op-skill-row">
            <div className="op-skill-row-main">
              <Checkbox
                checked={!props.disabledSkills.includes(s.name)}
                label={`/${s.name}`}
                onChange={(_, d) => props.onToggle(s.name, Boolean(d.checked))}
              />
              <Menu>
                <MenuTrigger disableButtonEnhancement>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<MoreHorizontal24Regular />}
                    aria-label={`Actions for ${s.name}`}
                  />
                </MenuTrigger>
                <MenuPopover>
                  <MenuList>
                    <MenuItem onClick={() => props.onOpenEditor(viewState(s))}>View</MenuItem>
                    {s.origin === "imported" && (
                      <MenuItem onClick={() => props.onOpenEditor(editState(s))}>Edit</MenuItem>
                    )}
                    {(s.origin === "bundled" || s.origin === "catalog") && (
                      <MenuItem onClick={() => props.onOpenEditor(customizeState(s))}>Customize</MenuItem>
                    )}
                    <MenuItem onClick={() => props.onOpenEditor(duplicateState(s, props.skills.map((x) => x.name)))}>
                      Duplicate
                    </MenuItem>
                    <MenuItem onClick={() => props.onExport(s)}>Export</MenuItem>
                    {s.overridesBundled && (
                      <MenuItem onClick={() => setPendingRemove(s.name)}>Reset to bundled</MenuItem>
                    )}
                    {s.origin === "imported" && !s.overridesBundled && (
                      <MenuItem onClick={() => setPendingRemove(s.name)}>Remove</MenuItem>
                    )}
                  </MenuList>
                </MenuPopover>
              </Menu>
            </div>
            <Caption1>
              {sourceLabel(s)} · {invocationLabel(s)} · {s.hosts.join(", ")}
            </Caption1>
            <Caption1>{s.description}</Caption1>
          </li>
        ))}
      </ul>

      <details className="op-settings-section">
        <summary>
          Import<span>URL or paste a SKILL.md</span>
        </summary>
        <Field label="Import skill from URL">
          <div className="op-import">
            <Input value={importUrl} onChange={(_, d) => setImportUrl(d.value)} placeholder="https://…/SKILL.md" />
            <Button
              disabled={!importUrl.trim()}
              onClick={async () => {
                setImportMsg(null);
                try {
                  await props.onImportUrl(importUrl.trim());
                  setImportMsg("Skill imported.");
                  setImportUrl("");
                } catch (err) {
                  setImportMsg(err instanceof Error ? err.message : String(err));
                }
              }}
            >
              Import
            </Button>
          </div>
        </Field>
        <Field label="Or paste a SKILL.md">
          <Textarea value={paste} onChange={(_, d) => setPaste(d.value)} placeholder={"---\nname: my-skill\n..."} />
          <Button
            disabled={!paste.trim()}
            onClick={async () => {
              setImportMsg(null);
              try {
                await props.onImportMarkdown(paste);
                setPaste("");
                setImportMsg("Skill imported.");
              } catch (err) {
                setImportMsg(err instanceof Error ? err.message : String(err));
              }
            }}
          >
            Import paste
          </Button>
        </Field>
        {importMsg && <Caption1>{importMsg}</Caption1>}
      </details>

      <Dialog
        modalType="alert"
        open={pendingRemove !== null}
        onOpenChange={(_, data) => {
          if (!data.open) setPendingRemove(null);
        }}
      >
        <DialogSurface className="op-restore-dialog">
          <DialogBody>
            <DialogTitle>
              {props.skills.find((s) => s.name === pendingRemove)?.overridesBundled
                ? `Reset “${pendingRemove}” to bundled?`
                : `Remove “${pendingRemove}”?`}
            </DialogTitle>
            <DialogContent>
              {props.skills.find((s) => s.name === pendingRemove)?.overridesBundled
                ? "Your customized copy will be deleted. The bundled skill comes back."
                : "This cannot be undone."}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setPendingRemove(null)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => {
                  const name = pendingRemove;
                  setPendingRemove(null);
                  if (name) void props.onRemove(name);
                }}
              >
                {props.skills.find((s) => s.name === pendingRemove)?.overridesBundled ? "Reset" : "Remove"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

function editorTitle(state: SkillEditorState): string {
  if (state.mode === "skillify") return "Save chat as skill";
  if (state.mode === "create") return "New skill";
  if (state.mode === "view") return "View skill";
  return "Edit skill";
}

function viewState(skill: ManagedSkill): SkillEditorState {
  return { markdown: skill.markdown, mode: "view", generating: false, error: null };
}

function editState(skill: ManagedSkill): SkillEditorState {
  return {
    markdown: skill.markdown,
    previousName: skill.name,
    mode: "edit",
    generating: false,
    error: null
  };
}

function customizeState(skill: ManagedSkill): SkillEditorState {
  return { markdown: skill.markdown, mode: "edit", generating: false, error: null };
}

function duplicateState(skill: ManagedSkill, taken: string[]): SkillEditorState {
  const parsed = parseSkillMarkdown(skill.markdown, skill.name);
  const name = uniqueSkillName(skill.name, taken);
  const markdown = serializeSkill({ ...parsed, name, rootPath: name });
  return { markdown, mode: "create", generating: false, error: null };
}

function sourceLabel(s: ManagedSkill): string {
  if (s.overridesBundled) return "imported · overrides bundled";
  return s.origin;
}

function invocationLabel(s: ManagedSkill): string {
  if (s.inject === "always") return "Always injected";
  if (s.disableModelInvocation) return "Slash only";
  if (s.userInvocable === false) return "Model loads";
  return "Slash";
}
