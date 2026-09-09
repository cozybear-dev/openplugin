import {
  Button,
  Checkbox,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Spinner,
  Textarea
} from "@fluentui/react-components";
import { ArrowLeft24Regular } from "@fluentui/react-icons";
import {
  parseSkillMarkdown,
  serializeSkill,
  SKILL_HOSTS,
  type HostKind,
  type Skill
} from "@openplugin/core";
import { useMemo, useState } from "react";

export type SkillEditorMode = "create" | "edit" | "view" | "skillify";

export function parseSkillDraft(markdown: string): { skill: Skill; error: null } | { skill: null; error: string } {
  try {
    const name = markdown.match(/^name:\s*(\S+)/m)?.[1]?.trim() ?? "";
    const folder = /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name) ? name : "draft-skill";
    return { skill: parseSkillMarkdown(markdown, folder), error: null };
  } catch (err) {
    return { skill: null, error: err instanceof Error ? err.message : String(err) };
  }
}

export function SkillEditor(props: {
  markdown: string;
  title: string;
  generating?: boolean;
  error?: string | null;
  readOnly?: boolean;
  onChange: (markdown: string) => void;
  onSave: () => void;
  onCancel: () => void;
}) {
  const [surface, setSurface] = useState<"form" | "markdown">("form");
  const parsed = useMemo(() => parseSkillDraft(props.markdown), [props.markdown]);
  const skill = parsed.skill;
  const parseError = parsed.error;
  const locked = Boolean(props.readOnly || props.generating);

  function patch(partial: Partial<Skill>) {
    if (!skill || locked) return;
    const next: Skill = {
      ...skill,
      ...partial,
      rootPath: partial.name ?? skill.name
    };
    props.onChange(serializeSkill(next));
  }

  return (
    <div className="op-settings op-skill-editor">
      <div className="op-settings-head">
        <div>
          <Button appearance="subtle" size="small" icon={<ArrowLeft24Regular />} onClick={props.onCancel}>
            Skills
          </Button>
          <h1>{props.title}</h1>
        </div>
        <div className="op-skill-editor-actions">
          {!props.readOnly && (
            <Button
              appearance="primary"
              size="small"
              disabled={locked || Boolean(parseError)}
              onClick={props.onSave}
            >
              Save
            </Button>
          )}
        </div>
      </div>

      {props.generating && <Spinner size="tiny" label="Drafting skill from chat…" />}
      {(props.error || parseError) && (
        <MessageBar intent="error">
          <MessageBarBody>{props.error || parseError}</MessageBarBody>
        </MessageBar>
      )}

      <div className="op-skill-surfaces">
        <Button
          size="small"
          appearance={surface === "form" ? "primary" : "subtle"}
          onClick={() => setSurface("form")}
        >
          Form
        </Button>
        <Button
          size="small"
          appearance={surface === "markdown" ? "primary" : "subtle"}
          onClick={() => setSurface("markdown")}
        >
          SKILL.md
        </Button>
      </div>

      {surface === "markdown" ? (
        <Field label="SKILL.md">
          <Textarea
            value={props.markdown}
            disabled={locked}
            textarea={{ className: "op-skill-markdown" }}
            onChange={(_, d) => props.onChange(d.value)}
          />
        </Field>
      ) : (
        <>
          <Field label="Name">
            <Input
              value={skill?.name ?? ""}
              disabled={locked || !skill}
              onChange={(_, d) => patch({ name: d.value })}
            />
          </Field>
          <Field label="Description">
            <Textarea
              value={skill?.description ?? ""}
              disabled={locked || !skill}
              onChange={(_, d) => patch({ description: d.value })}
            />
          </Field>
          <Field label="Hosts">
            <div className="op-skill-hosts">
              {SKILL_HOSTS.map((host: HostKind) => (
                <Checkbox
                  key={host}
                  label={host}
                  disabled={locked || !skill}
                  checked={skill?.hosts.includes(host) ?? false}
                  onChange={(_, d) => {
                    if (!skill) return;
                    const on = Boolean(d.checked);
                    const hosts = on
                      ? [...new Set([...skill.hosts, host])]
                      : skill.hosts.filter((h) => h !== host);
                    patch({ hosts: hosts.length ? hosts : [host] });
                  }}
                />
              ))}
            </div>
          </Field>
          <Checkbox
            label="Show in / menu"
            disabled={locked || !skill}
            checked={skill?.userInvocable ?? true}
            onChange={(_, d) => patch({ userInvocable: Boolean(d.checked) })}
          />
          <Checkbox
            label="Hide from the model"
            disabled={locked || !skill}
            checked={skill?.disableModelInvocation ?? false}
            onChange={(_, d) => patch({ disableModelInvocation: Boolean(d.checked) })}
          />
          <Checkbox
            label="Always inject"
            disabled={locked || !skill}
            checked={skill?.inject === "always"}
            onChange={(_, d) => patch({ inject: d.checked ? "always" : "on-demand" })}
          />
          <Field label="Instructions">
            <Textarea
              value={skill?.body ?? ""}
              disabled={locked || !skill}
              textarea={{ className: "op-skill-body" }}
              onChange={(_, d) => patch({ body: d.value })}
            />
          </Field>
        </>
      )}
    </div>
  );
}
