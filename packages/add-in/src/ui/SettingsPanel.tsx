import {
  Body1,
  Button,
  Caption1,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Textarea
} from "@fluentui/react-components";
import { useState } from "react";
import { chatCompletions, LlmError, type ProviderConfig } from "@openplugin/core";
import type { CompanionStatus } from "../companion";
import { matchPreset, PRESETS } from "../presets";

export function SettingsPanel(props: {
  provider: ProviderConfig;
  instructions: string;
  companion: CompanionStatus;
  policyNote?: string;
  onClose: () => void;
  onChange: (next: ProviderConfig) => Promise<void>;
  onInstructions: (text: string) => Promise<void>;
  onImportUrl: (url: string) => Promise<void>;
}) {
  const preset = matchPreset(props.provider.baseUrl);
  const [testing, setTesting] = useState(false);
  const [testMsg, setTestMsg] = useState<string | null>(null);
  const [importUrl, setImportUrl] = useState("");
  const [importMsg, setImportMsg] = useState<string | null>(null);

  function patch(partial: Partial<ProviderConfig>) {
    void props.onChange({ ...props.provider, ...partial });
  }

  async function testConnection() {
    setTesting(true);
    setTestMsg(null);
    try {
      const result = await chatCompletions({
        config: { ...props.provider, maxOutputTokens: 8 },
        messages: [{ role: "user", content: "Reply with ok" }]
      });
      const text = result.message.role === "assistant" ? result.message.content : "";
      setTestMsg(`Connected. ${text || "Empty reply."}`);
    } catch (err) {
      setTestMsg(formatError(err, props.companion.state === "connected"));
    } finally {
      setTesting(false);
    }
  }

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
      <Field label="Model">
        <Input value={props.provider.model} onChange={(_, d) => patch({ model: d.value })} />
      </Field>
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

      <Button disabled={testing || !props.provider.baseUrl} onClick={() => void testConnection()}>
        {testing ? "Testing…" : "Test connection"}
      </Button>
      {testMsg && <Caption1>{testMsg}</Caption1>}

      <Field label="Instructions for this host">
        <Textarea
          value={props.instructions}
          placeholder="Always use thousand separators. Keep formulas."
          onChange={(_, d) => void props.onInstructions(d.value)}
        />
      </Field>

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
      {importMsg && <Caption1>{importMsg}</Caption1>}
      {props.policyNote && <Caption1>{props.policyNote}</Caption1>}
      <Caption1>
        Keys stay on this machine. The add-in talks to your endpoint directly — nothing is proxied
        through us.
      </Caption1>
    </div>
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
