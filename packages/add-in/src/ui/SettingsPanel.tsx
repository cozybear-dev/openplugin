import {
  Body1,
  Button,
  Caption1,
  Checkbox,
  Combobox,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Textarea
} from "@fluentui/react-components";
import { CheckmarkFilled, DismissCircleFilled } from "@fluentui/react-icons";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  chatCompletions,
  LlmError,
  type ModelInfo,
  type ProviderConfig,
  type SkillCatalogEntry
} from "@openplugin/core";
import type { AuditEntry } from "../audit";
import type { CompanionStatus } from "../companion";
import { matchPreset, PRESETS } from "../presets";

export function SettingsPanel(props: {
  provider: ProviderConfig;
  instructions: string;
  companion: CompanionStatus;
  policyNote?: string;
  skills: Array<SkillCatalogEntry & { source: "bundled" | "imported" }>;
  disabledSkills: string[];
  audit: AuditEntry[];
  autoApply: boolean;
  models: ModelInfo[];
  modelsError: string | null;
  modelsLoading: boolean;
  onClose: () => void;
  onChange: (next: ProviderConfig) => Promise<void>;
  onRefreshModels: () => void;
  onInstructions: (text: string) => Promise<void>;
  onImportUrl: (url: string) => Promise<void>;
  onImportMarkdown: (md: string) => Promise<void>;
  onRemoveSkill: (name: string) => Promise<void>;
  onToggleSkill: (name: string, enabled: boolean) => void;
  onAutoApply: (on: boolean) => void;
  onTestLogged: (ok: boolean, summary: string) => void;
}) {
  const preset = matchPreset(props.provider.baseUrl);
  const [testing, setTesting] = useState(false);
  const [testStatus, setTestStatus] = useState<"idle" | "ok" | "fail">("idle");
  const [testError, setTestError] = useState<string | null>(null);
  const [importUrl, setImportUrl] = useState("");
  const [paste, setPaste] = useState("");
  const [importMsg, setImportMsg] = useState<string | null>(null);
  const [modelText, setModelText] = useState(props.provider.model);
  const ignoreNextCommit = useRef(false);

  useEffect(() => {
    setTestStatus("idle");
    setTestError(null);
  }, [props.provider.baseUrl, props.provider.apiKey, props.provider.model]);

  useEffect(() => {
    setModelText(props.provider.model);
  }, [props.provider.model]);

  const filteredModels = useMemo(
    () => filterModels(props.models, modelText, props.provider.model),
    [props.models, modelText, props.provider.model]
  );

  function commitTypedModel() {
    if (ignoreNextCommit.current) {
      ignoreNextCommit.current = false;
      return;
    }
    const typed = modelText.trim();
    const id = typed ? modelIdFromInput(props.models, typed) : "";
    setModelText(id);
    if (id !== props.provider.model) patch({ model: id });
  }

  function patch(partial: Partial<ProviderConfig>) {
    void props.onChange({ ...props.provider, ...partial });
  }

  async function testConnection() {
    setTesting(true);
    setTestError(null);
    try {
      await chatCompletions({
        config: { ...props.provider, maxOutputTokens: 8 },
        messages: [{ role: "user", content: "Reply with ok" }]
      });
      setTestStatus("ok");
      props.onTestLogged(true, `Connected to ${props.provider.model}`);
    } catch (err) {
      setTestStatus("fail");
      const msg = formatError(err, props.companion.state === "connected");
      setTestError(msg);
      props.onTestLogged(false, msg);
    } finally {
      setTesting(false);
    }
  }

  const testLabel =
    testing ? "Testing…" : testStatus === "ok" ? `Connected · ${props.provider.model || "ok"}` : "Test connection";

  return (
    <div className="op-settings">
      <div className="op-settings-head">
        <Body1>Settings</Body1>
        <Button appearance="subtle" size="small" onClick={props.onClose}>
          Done
        </Button>
      </div>

      <Field label="Endpoint">
        <Dropdown
          value={preset.name}
          onOptionSelect={(_, data) => {
            const next = PRESETS.find((p) => p.id === data.optionValue);
            if (!next) return;
            void props.onChange({
              ...props.provider,
              baseUrl: next.baseUrl,
              model: next.model || props.provider.model
            });
          }}
        >
          {PRESETS.map((p) => (
            <Option key={p.id} value={p.id}>
              {p.name}
            </Option>
          ))}
        </Dropdown>
      </Field>

      <Field label="Base URL">
        <Input
          value={props.provider.baseUrl}
          placeholder="https://host/v1"
          onChange={(_, d) => patch({ baseUrl: d.value })}
        />
      </Field>
      <Field label="Default model">
        <div className="op-import">
          <Combobox
            aria-label="Default model"
            freeform
            placeholder="Model"
            value={modelText}
            selectedOptions={props.provider.model ? [props.provider.model] : []}
            onOptionSelect={(_, data) => {
              if (data.optionValue == null) return;
              ignoreNextCommit.current = true;
              setModelText(data.optionValue);
              if (data.optionValue !== props.provider.model) patch({ model: data.optionValue });
              queueMicrotask(() => {
                ignoreNextCommit.current = false;
              });
            }}
            onChange={(e) => setModelText(e.target.value)}
            onBlur={commitTypedModel}
          >
            {filteredModels.map((m) => (
              <Option key={m.id} value={m.id} text={m.name ?? m.id}>
                {m.name ?? m.id}
              </Option>
            ))}
          </Combobox>
          <Button
            size="small"
            disabled={props.modelsLoading || !props.provider.baseUrl}
            onClick={() => props.onRefreshModels()}
          >
            Refresh
          </Button>
        </div>
      </Field>
      {props.modelsError && (
        <MessageBar intent="error">
          <MessageBarBody>{props.modelsError}</MessageBarBody>
        </MessageBar>
      )}
      {preset.needsKey && (
        <Field label="API key">
          <Input
            type="password"
            value={props.provider.apiKey ?? ""}
            onChange={(_, d) => patch({ apiKey: d.value })}
          />
        </Field>
      )}

      <MessageBar intent={props.companion.state === "connected" ? "success" : "info"}>
        <MessageBarBody>
          {props.companion.state === "connected"
            ? "Companion connected on 127.0.0.1:8788. Local models can skip CORS."
            : "Companion not running. Needed for Ollama unless the server sends CORS headers. npm run companion"}
        </MessageBarBody>
      </MessageBar>

      <Button
        className={testStatus === "ok" ? "op-test-ok" : undefined}
        appearance={testStatus === "fail" ? "outline" : testStatus === "ok" ? "primary" : "secondary"}
        icon={
          testStatus === "ok" ? (
            <CheckmarkFilled />
          ) : testStatus === "fail" ? (
            <DismissCircleFilled />
          ) : undefined
        }
        disabled={testing || !props.provider.baseUrl}
        onClick={() => void testConnection()}
      >
        {testLabel}
      </Button>
      {testStatus === "fail" && testError && (
        <MessageBar intent="error">
          <MessageBarBody>{testError}</MessageBarBody>
        </MessageBar>
      )}

      <Checkbox
        checked={props.autoApply}
        label="Auto-apply this session (skip review)"
        onChange={(_, d) => props.onAutoApply(Boolean(d.checked))}
      />

      <Field label="Instructions for this host">
        <Textarea
          value={props.instructions}
          placeholder="Always use thousand separators. Keep formulas."
          onChange={(_, d) => void props.onInstructions(d.value)}
        />
      </Field>

      <Body1>Skills</Body1>
      <ul className="op-skill-list">
        {props.skills.map((s) => (
          <li key={s.name}>
            <Checkbox
              checked={!props.disabledSkills.includes(s.name)}
              label={`/${s.name}`}
              onChange={(_, d) => props.onToggleSkill(s.name, Boolean(d.checked))}
            />
            <Caption1>
              {s.source} · {s.description}
            </Caption1>
            {s.source === "imported" && (
              <Button size="small" appearance="subtle" onClick={() => void props.onRemoveSkill(s.name)}>
                Remove
              </Button>
            )}
          </li>
        ))}
      </ul>

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
        <Textarea value={paste} onChange={(_, d) => setPaste(d.value)} placeholder="---&#10;name: my-skill&#10;..." />
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
      {props.policyNote && <Caption1>{props.policyNote}</Caption1>}

      <Body1>Activity</Body1>
      {props.audit.length === 0 ? (
        <Caption1>Applies, rejects, connection tests, and skill runs show up here.</Caption1>
      ) : (
        <ul className="op-audit">
          {props.audit.slice(0, 40).map((e, i) => (
            <li key={`${e.ts}-${i}`}>
              <Caption1>
                {new Date(e.ts).toLocaleString()} · {e.action}
              </Caption1>
              <span>{e.summary}</span>
            </li>
          ))}
        </ul>
      )}
      <Caption1>
        Keys stay on this machine. The add-in talks to your endpoint directly — nothing is proxied
        through us.
      </Caption1>
    </div>
  );
}

function modelIdFromInput(models: ModelInfo[], raw: string): string {
  const exact = models.find((m) => m.id === raw);
  if (exact) return exact.id;
  const byName = models.filter((m) => (m.name ?? m.id) === raw);
  return byName.length === 1 ? byName[0]!.id : raw;
}

function filterModels(models: ModelInfo[], query: string, selected: string): ModelInfo[] {
  const q = query.trim().toLowerCase();
  if (!q || q === selected.trim().toLowerCase()) return models;
  return models.filter(
    (m) => m.id.toLowerCase().includes(q) || (m.name?.toLowerCase().includes(q) ?? false)
  );
}

export function formatError(err: unknown, companionUp: boolean): string {
  if (err instanceof LlmError && err.code === "cors") {
    return companionUp
      ? "The endpoint still blocked the request. Check the URL."
      : "CORS blocked this endpoint. Enable CORS on the server, or run npm run companion and retry.";
  }
  if (err instanceof LlmError && err.code === "auth") return "The endpoint rejected the API key.";
  if (err instanceof Error) return err.message;
  return String(err);
}
