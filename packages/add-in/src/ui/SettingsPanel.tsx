import {
  Button,
  Caption1,
  Checkbox,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Textarea
} from "@fluentui/react-components";
import { CheckmarkFilled, DismissCircleFilled } from "@fluentui/react-icons";
import { useEffect, useState } from "react";
import { ModelPicker } from "./ModelPicker";
import {
  chatCompletions,
  LlmError,
  type ModelInfo,
  type ProviderConfig,
  type SkillCatalogEntry
} from "@openplugin/core";
import type { AuditEntry } from "../audit";
import type { CompanionStatus } from "../companion";
import { clearDebug, getDebug, setDebugPersist, type DebugEvent } from "../debug";
import { matchPreset, PRESETS } from "../presets";
import type { SearchSettings } from "../settings";

export function SettingsPanel(props: {
  provider: ProviderConfig;
  instructions: string;
  companion: CompanionStatus;
  policyNote?: string;
  skills: Array<SkillCatalogEntry & { source: "bundled" | "imported" }>;
  disabledSkills: string[];
  audit: AuditEntry[];
  autoApply: boolean;
  search: SearchSettings;
  debug: DebugEvent[];
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
  onSearch: (next: SearchSettings) => Promise<void>;
  onTestLogged: (ok: boolean, summary: string) => void;
}) {
  const preset = matchPreset(props.provider.baseUrl);
  const [testing, setTesting] = useState(false);
  const [testStatus, setTestStatus] = useState<"idle" | "ok" | "fail">("idle");
  const [testError, setTestError] = useState<string | null>(null);
  const [importUrl, setImportUrl] = useState("");
  const [paste, setPaste] = useState("");
  const [importMsg, setImportMsg] = useState<string | null>(null);

  useEffect(() => {
    setTestStatus("idle");
    setTestError(null);
  }, [props.provider.baseUrl, props.provider.apiKey, props.provider.model]);

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
        <div><div className="op-eyebrow">YOUR WORKSPACE, YOUR WAY</div><h1>Settings</h1></div>
        <Button appearance="subtle" size="small" onClick={props.onClose}>
          Done
        </Button>
      </div>

      <section className="op-settings-section"><h2>Model connection</h2><p className="op-section-description">Choose the model you want to work with.</p><Field label="Provider">
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
<ModelPicker label="Default model" model={props.provider.model} models={props.models} loading={props.modelsLoading} onChange={(model) => patch({ model })} />
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

      {(preset.id === "ollama" || props.companion.state === "connected") && <MessageBar intent={props.companion.state === "connected" ? "success" : "info"}>
        <MessageBarBody>
          {props.companion.state === "connected"
            ? "Companion connected on 127.0.0.1:8788. Local models can skip CORS."
            : "Companion not running. Needed for Ollama unless the server sends CORS headers. npm run companion"}
        </MessageBarBody>
      </MessageBar>}

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

      </section><details className="op-settings-section"><summary>Preferences<span>Instructions &amp; change review</span></summary>
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

      </details><details className="op-settings-section"><summary>Skills<span>Manage shortcuts &amp; import your own</span></summary>
      <ul className="op-skill-list">
        {props.skills.map((s) => (
          <li key={s.name}>
            <Checkbox
              checked={!props.disabledSkills.includes(s.name)}
              label={`/${s.name}`}
              onChange={(_, d) => props.onToggleSkill(s.name, Boolean(d.checked))}
            />
            <Caption1>
              {s.source} · {skillBadge(s)} · {s.description}
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

      </details><details className="op-settings-section"><summary>Web search<span>Sources &amp; search preferences</span></summary>
      <Checkbox
        checked={props.search.defaultEnabled}
        label="Enable web search by default in new chats"
        onChange={(_, d) => void props.onSearch({ ...props.search, defaultEnabled: Boolean(d.checked) })}
      />
      <Field label="Search backend">
        <Dropdown
          value={backendLabel(props.search.backend)}
          onOptionSelect={(_, data) => {
            const backend = data.optionValue as SearchSettings["backend"];
            if (!backend) return;
            void props.onSearch({ ...props.search, backend });
          }}
        >
          <Option value="auto">Auto (native → free → Exa)</Option>
          <Option value="native">Provider native</Option>
          <Option value="duckduckgo">DuckDuckGo (free)</Option>
          <Option value="exa">Exa</Option>
          <Option value="custom">Custom URL</Option>
        </Dropdown>
      </Field>
      <Field label="Exa API key">
        <Input
          type="password"
          value={props.search.exaApiKey ?? ""}
          onChange={(_, d) => void props.onSearch({ ...props.search, exaApiKey: d.value })}
        />
      </Field>
      {props.search.backend === "custom" && (
        <Field label="Custom search URL">
          <Input
            value={props.search.custom?.url ?? ""}
            placeholder="https://… POST { query, maxResults }"
            onChange={(_, d) =>
              void props.onSearch({ ...props.search, custom: { ...props.search.custom, url: d.value } })
            }
          />
        </Field>
      )}
      <Caption1>
        Auto uses the model provider’s search when the endpoint supports it (OpenRouter, OpenAI). Otherwise it
        uses DuckDuckGo. If the free engine is rate-limited, add an Exa key.
      </Caption1>

      </details><details className="op-settings-section"><summary>Activity<span>Recent changes &amp; connection tests</span></summary>
      {props.audit.length === 0 ? (
        <Caption1>Applies, rejects, connection tests, and skill runs show up here.</Caption1>
      ) : (
        <ul className="op-audit">
          {props.audit.slice(0, 40).map((e, i) => (
            <li key={`${e.ts}-${i}`}>
              <Caption1>
                {new Date(e.ts).toLocaleString()} · {e.program ?? e.host} · {e.documentTitle ?? "document"} ·{" "}
                {e.action}
              </Caption1>
              <span>{e.summary}</span>
            </li>
          ))}
        </ul>
      )}

      </details><details className="op-settings-section"><summary>Diagnostics<span>Debug logs &amp; troubleshooting</span></summary>
      <Checkbox
        label="Keep debug log for this session"
        onChange={(_, d) => setDebugPersist(Boolean(d.checked))}
      />
      <div className="op-debug-actions">
        <Button
          size="small"
          onClick={() => {
            void navigator.clipboard.writeText(JSON.stringify(getDebug(), null, 2));
          }}
        >
          Copy
        </Button>
        <Button
          size="small"
          onClick={() => {
            const blob = new Blob([getDebug().map((e) => JSON.stringify(e)).join("\n")], {
              type: "application/jsonl"
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "openplugin-debug.jsonl";
            a.click();
            URL.revokeObjectURL(url);
          }}
        >
          Download
        </Button>
        <Button size="small" onClick={() => clearDebug()}>
          Clear
        </Button>
      </div>
      {props.debug.length === 0 ? (
        <Caption1>Agent, apply, and search events show up here.</Caption1>
      ) : (
        <ul className="op-debug">
          {props.debug.slice(0, 80).map((e, i) => (
            <li key={`${e.ts}-${i}`}>
              <Caption1>
                {e.level} · {e.source} · {new Date(e.ts).toLocaleTimeString()}
              </Caption1>
              <code>{e.message}</code>
            </li>
          ))}
        </ul>
      )}
      </details><Caption1>
        Keys stay on this machine. The add-in talks to your endpoint directly — nothing is proxied
        through us.
      </Caption1>
    </div>
  );
}

function skillBadge(s: SkillCatalogEntry): string {
  if (s.inject === "always") return "Always injected";
  if (s.disableModelInvocation) return "Slash only";
  if (s.userInvocable === false) return "Model loads";
  return "Slash";
}

function backendLabel(backend: SearchSettings["backend"]): string {
  if (backend === "native") return "Provider native";
  if (backend === "duckduckgo") return "DuckDuckGo (free)";
  if (backend === "exa") return "Exa";
  if (backend === "custom") return "Custom URL";
  return "Auto (native → free → Exa)";
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
