import {
  chatCompletions,
  LlmError,
  runAgent,
  type AgentEvent,
  type AgentResult,
  type ChatMessage,
  type HostKind,
  type ProviderConfig
} from "@openplugin/core";
import {
  Body1,
  Button,
  Dropdown,
  Field,
  Input,
  Option,
  Spinner,
  Tab,
  TabList,
  Textarea,
  Title3
} from "@fluentui/react-components";
import { useMemo, useState } from "react";
import { bundledSkills } from "./bundled-skills";
import { PRESETS } from "./presets";
import { createHost } from "./runtime-host";
import { loadProvider, saveProvider } from "./settings";

type ChatLine =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string }
  | { kind: "tool"; text: string }
  | { kind: "error"; text: string };

export function App(props: { hostKind: HostKind; inOffice: boolean }) {
  const host = useMemo(() => createHost(props.hostKind, props.inOffice), [props.hostKind, props.inOffice]);
  const skills = useMemo(() => bundledSkills(), []);
  const [tab, setTab] = useState<"chat" | "settings">("chat");
  const [provider, setProvider] = useState<ProviderConfig>(loadProvider);
  const [input, setInput] = useState("");
  const [lines, setLines] = useState<ChatLine[]>([]);
  const [history, setHistory] = useState<ChatMessage[]>([]);
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState<string | null>(null);
  const [pending, setPending] = useState<AgentResult | null>(null);
  const [status, setStatus] = useState<string | null>(null);

  async function send() {
    const text = input.trim();
    if (!text || busy) return;
    setInput("");
    setLines((l) => [...l, { kind: "user", text }]);
    setBusy(true);
    setStatus(null);
    const controller = new AbortController();
    try {
      const result = await runAgent({
        config: provider,
        host,
        skills,
        userMessage: text,
        history,
        signal: controller.signal,
        onEvent: (event: AgentEvent) => {
          if (event.type === "text") {
            setLines((l) => {
              const last = l[l.length - 1];
              if (last?.kind === "assistant") {
                return [...l.slice(0, -1), { kind: "assistant", text: last.text + event.delta }];
              }
              return [...l, { kind: "assistant", text: event.delta }];
            });
          } else if (event.type === "tool") {
            setLines((l) => [...l, { kind: "tool", text: event.name }]);
          } else if (event.type === "skill") {
            setLines((l) => [...l, { kind: "tool", text: `skill:${event.name}` }]);
          } else if (event.type === "error") {
            setLines((l) => [...l, { kind: "error", text: event.message }]);
          }
        }
      });
      setHistory(
        result.messages.filter(
          (m): m is ChatMessage => m.role === "user" || m.role === "assistant" || m.role === "tool"
        )
      );
      if (result.finalText) {
        setLines((l) => {
          const last = l[l.length - 1];
          if (last?.kind === "assistant") return l;
          return [...l, { kind: "assistant", text: result.finalText }];
        });
      }
      if (!result.changeset.isEmpty()) {
        setPreview(result.changeset.preview());
        setPending(result);
      } else {
        setPreview(null);
        setPending(null);
      }
    } catch (err) {
      setLines((l) => [...l, { kind: "error", text: formatError(err) }]);
    } finally {
      setBusy(false);
    }
  }

  async function apply() {
    const result = pending;
    if (!result) return;
    await host.apply(result.changeset);
    result.changeset.clear();
    setPreview(null);
    setPending(null);
    setStatus("Applied.");
  }

  function reject() {
    setPreview(null);
    setPending(null);
    setStatus("Discarded.");
  }

  return (
    <div className="shell">
      <header className="top">
        <Title3>OpenPlugin</Title3>
        <Body1 className="host-badge">{props.hostKind}</Body1>
      </header>
      <TabList
        selectedValue={tab}
        onTabSelect={(_, data) => setTab(data.value as "chat" | "settings")}
      >
        <Tab value="chat">Chat</Tab>
        <Tab value="settings">Settings</Tab>
      </TabList>
      {tab === "settings" ? (
        <SettingsForm
          provider={provider}
          onChange={async (next) => {
            setProvider(next);
            await saveProvider(next);
          }}
        />
      ) : (
        <>
          <div className="thread">
            {lines.length === 0 && (
              <Body1>
                Point Settings at any OpenAI-compatible endpoint, then describe an edit. Mutations
                wait for Apply.
              </Body1>
            )}
            {lines.map((line, i) => (
              <div key={i} className={`line ${line.kind}`}>
                {line.text}
              </div>
            ))}
            {busy && <Spinner size="tiny" label="Working" />}
          </div>
          {preview && (
            <div className="changeset">
              <pre>{preview}</pre>
              <div className="row">
                <Button appearance="primary" onClick={() => void apply()}>
                  Apply
                </Button>
                <Button onClick={reject}>Reject</Button>
              </div>
            </div>
          )}
          {status && <Body1>{status}</Body1>}
          <div className="composer">
            <Textarea
              value={input}
              onChange={(_, d) => setInput(d.value)}
              placeholder="Ask OpenPlugin to edit the document…"
              disabled={busy}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  void send();
                }
              }}
            />
            <Button appearance="primary" disabled={busy || !input.trim()} onClick={() => void send()}>
              Send
            </Button>
          </div>
        </>
      )}
    </div>
  );
}

function SettingsForm(props: {
  provider: ProviderConfig;
  onChange: (next: ProviderConfig) => Promise<void>;
}) {
  const { provider, onChange } = props;
  const [testing, setTesting] = useState(false);
  const [testMsg, setTestMsg] = useState<string | null>(null);

  function patch(partial: Partial<ProviderConfig>) {
    void onChange({ ...provider, ...partial });
  }

  async function testConnection() {
    setTesting(true);
    setTestMsg(null);
    try {
      const result = await chatCompletions({
        config: { ...provider, maxOutputTokens: 8 },
        messages: [{ role: "user", content: "Reply with ok" }]
      });
      const text = result.message.role === "assistant" ? result.message.content : "";
      setTestMsg(`Connected. Model said: ${text || "(empty)"}`);
    } catch (err) {
      setTestMsg(formatError(err));
    } finally {
      setTesting(false);
    }
  }

  return (
    <div className="settings">
      <Field label="Preset">
        <Dropdown
          value={PRESETS.find((p) => p.baseUrl === provider.baseUrl)?.name ?? "Custom"}
          onOptionSelect={(_, data) => {
            const preset = PRESETS.find((p) => p.id === data.optionValue);
            if (!preset) return;
            void onChange({ ...provider, baseUrl: preset.baseUrl, model: preset.model || provider.model });
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
        <Input value={provider.baseUrl} onChange={(_, d) => patch({ baseUrl: d.value })} />
      </Field>
      <Field label="Model">
        <Input value={provider.model} onChange={(_, d) => patch({ model: d.value })} />
      </Field>
      <Field label="API key (stored in this Office profile, never in the document)">
        <Input
          type="password"
          value={provider.apiKey ?? ""}
          onChange={(_, d) => patch({ apiKey: d.value })}
        />
      </Field>
      <Button disabled={testing} onClick={() => void testConnection()}>
        {testing ? "Testing…" : "Test connection"}
      </Button>
      {testMsg && <Body1>{testMsg}</Body1>}
      <Body1>
        Keys stay on this machine. The add-in talks to your endpoint directly. If the browser blocks
        the call (CORS), enable CORS on the server or wait for the optional local companion.
      </Body1>
    </div>
  );
}

function formatError(err: unknown): string {
  if (err instanceof LlmError && err.code === "cors") {
    return "This endpoint blocked the request (CORS). Enable CORS on the server, or use a local companion in a later release.";
  }
  if (err instanceof LlmError && err.code === "auth") {
    return "The endpoint rejected the API key.";
  }
  if (err instanceof Error) return err.message;
  return String(err);
}
