import {
  assertPolicy,
  compileSnapshot,
  describeTool,
  LlmError,
  mergePolicy,
  OPEN_POLICY,
  runAgent,
  type AgentResult,
  type ChatMessage,
  type DocumentSnapshot,
  type HostKind,
  type Policy,
  type ProviderConfig,
  type Skill
} from "@openplugin/core";
import { useEffect, useMemo, useRef, useState } from "react";
import { bundledSkills } from "./bundled-skills";
import { fetchPolicy, postAudit, probeCompanion, type CompanionStatus } from "./companion";
import { clearHistory, loadHistory, saveHistory, type StoredLine } from "./history";
import { COMPANION_ORIGIN, isOllamaUrl, isOpenRouterUrl } from "./presets";
import { createHost } from "./runtime-host";
import { loadInstructions, loadProvider, loadUserPolicy, saveInstructions, saveProvider } from "./settings";
import { importCatalog, importSkillFromUrl, listImportedSkills } from "./skill-store";
import { Composer } from "./ui/Composer";
import { EmptyState } from "./ui/EmptyState";
import { formatError, SettingsPanel } from "./ui/SettingsPanel";
import { Header } from "./ui/Header";
import { ReviewCard } from "./ui/ReviewCard";
import { Thread } from "./ui/Thread";

export function App(props: { hostKind: HostKind; inOffice: boolean }) {
  const host = useMemo(() => createHost(props.hostKind, props.inOffice), [props.hostKind, props.inOffice]);
  const [imported, setImported] = useState<Skill[]>([]);
  const skills = useMemo(() => {
    const reg = bundledSkills();
    for (const s of imported) reg.add(s);
    return reg;
  }, [imported]);

  const [provider, setProvider] = useState<ProviderConfig>(loadProvider);
  const [instructions, setInstructions] = useState(loadInstructions);
  const [companion, setCompanion] = useState<CompanionStatus>({ state: "unknown" });
  const [policy, setPolicy] = useState<Policy>(OPEN_POLICY);
  const [tab, setTab] = useState<"chat" | "settings">("chat");
  const [input, setInput] = useState("");
  const [lines, setLines] = useState<StoredLine[]>(() => loadHistory(props.hostKind).lines);
  const [history, setHistory] = useState<ChatMessage[]>(() => loadHistory(props.hostKind).messages);
  const [busy, setBusy] = useState(false);
  const [pending, setPending] = useState<AgentResult | null>(null);
  const [alwaysApply, setAlwaysApply] = useState(false);
  const [contextLabel, setContextLabel] = useState(props.hostKind);
  const [snapshot, setSnapshot] = useState<DocumentSnapshot | undefined>();
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    void (async () => {
      setImported(await listImportedSkills().catch(() => []));
      const status = await probeCompanion();
      setCompanion(status);
      if (status.state === "connected") {
        const tenant = (await fetchPolicy(status.token)) as Policy | null;
        if (tenant) {
          const merged = mergePolicy(tenant, loadUserPolicy() ?? undefined);
          setPolicy(merged);
          if (merged.catalogUrl) {
            const extra = await importCatalog(merged.catalogUrl).catch(() => []);
            setImported((s) => [...s, ...extra]);
          }
        }
      }
    })();
  }, []);

  useEffect(() => {
    void refreshContext();
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
        void refreshContext();
      });
    } catch {
      /* browser */
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [host]);

  async function refreshContext() {
    try {
      const snap = await compileSnapshot(host, { tokenBudget: 800, previous: snapshot });
      setSnapshot(snap);
      setContextLabel(labelFromSnapshot(snap));
    } catch {
      setContextLabel(props.hostKind);
    }
  }

  async function persist(nextLines: StoredLine[], nextMessages: ChatMessage[]) {
    setLines(nextLines);
    setHistory(nextMessages);
    await saveHistory(props.hostKind, nextLines, nextMessages);
  }

  function effectiveConfig(base: ProviderConfig, viaCompanion: boolean): ProviderConfig {
    const headers = { ...(base.headers ?? {}) };
    if (isOpenRouterUrl(base.baseUrl)) {
      headers["HTTP-Referer"] ??= "https://openplugin.local";
      headers["X-Title"] ??= "OpenPlugin";
    }
    if (viaCompanion && companion.state === "connected") {
      return {
        ...base,
        baseUrl: `${COMPANION_ORIGIN}/v1`,
        headers: {
          ...headers,
          "x-openplugin-token": companion.token,
          "x-openplugin-target": base.baseUrl
        }
      };
    }
    return { ...base, headers };
  }

  async function send(text = input) {
    const prompt = text.trim();
    if (!prompt || busy) return;
    if (!provider.baseUrl || !provider.model) {
      setTab("settings");
      return;
    }
    try {
      assertPolicy(policy, { baseUrl: provider.baseUrl, model: provider.model });
    } catch (err) {
      setLines((l) => [...l, { kind: "error", text: err instanceof Error ? err.message : String(err) }]);
      return;
    }
    setInput("");
    const userLine: StoredLine = { kind: "user", text: prompt };
    const nextLines = [...lines, userLine];
    setLines(nextLines);
    setBusy(true);
    const controller = new AbortController();
    abortRef.current = controller;

    const useCompanionFirst = isOllamaUrl(provider.baseUrl) && companion.state === "connected";
    const message = instructions.trim()
      ? `${prompt}\n\nStanding instructions:\n${instructions.trim()}`
      : prompt;

    try {
      const result = await runTurn(message, effectiveConfig(provider, useCompanionFirst), controller.signal);
      await finishTurn(nextLines, result);
    } catch (err) {
      if (err instanceof LlmError && err.code === "cors") {
        let status = companion;
        if (status.state !== "connected") status = await probeCompanion();
        setCompanion(status);
        if (status.state === "connected" && !useCompanionFirst) {
          try {
            const result = await runTurn(message, effectiveConfig(provider, true), controller.signal);
            await finishTurn(nextLines, result);
            setBusy(false);
            return;
          } catch (retryErr) {
            setLines([...nextLines, { kind: "error", text: formatError(retryErr, true) }]);
            setBusy(false);
            return;
          }
        }
      }
      setLines([...nextLines, { kind: "error", text: formatError(err, companion.state === "connected") }]);
    } finally {
      setBusy(false);
      abortRef.current = null;
    }
  }

  async function runTurn(userMessage: string, config: ProviderConfig, signal: AbortSignal) {
    return runAgent({
      config,
      host,
      skills,
      userMessage,
      history,
      previousSnapshot: snapshot,
      policy,
      signal,
      onEvent: (event) => {
        if (event.type === "text") {
          setLines((l) => {
            const last = l[l.length - 1];
            if (last?.kind === "assistant") {
              return [...l.slice(0, -1), { kind: "assistant", text: last.text + event.delta }];
            }
            return [...l, { kind: "assistant", text: event.delta }];
          });
        } else if (event.type === "tool") {
          setLines((l) => [...l, { kind: "tool", text: describeTool(event.name, (event.args ?? {}) as Record<string, unknown>) }]);
        } else if (event.type === "skill") {
          setLines((l) => [...l, { kind: "tool", text: `Using skill ${event.name}` }]);
        } else if (event.type === "error") {
          setLines((l) => [...l, { kind: "error", text: event.message }]);
        }
      }
    });
  }

  async function finishTurn(baseLines: StoredLine[], result: AgentResult) {
    const withFinal = result.finalText
      ? appendAssistant(baseLines, result.finalText)
      : linesRefTail(baseLines);
    const msgs = result.messages.filter(
      (m): m is ChatMessage => m.role === "user" || m.role === "assistant" || m.role === "tool"
    );
    await persist(withFinal, msgs);
    setSnapshot(result.snapshot);
    if (companion.state === "connected" && policy.audit?.enabled) {
      void postAudit(companion.token, {
        host: props.hostKind,
        model: provider.model,
        endpoint: provider.baseUrl,
        skill: result.loadedSkills[0],
        toolNames: result.changeset.changes.map((c) => c.op),
        tokenEstimate: result.snapshot.tokenEstimate,
        user: "local"
      });
    }
    if (!result.changeset.isEmpty()) {
      if (alwaysApply) {
        await host.apply(result.changeset);
        result.changeset.clear();
        setPending(null);
      } else {
        setPending(result);
      }
    } else {
      setPending(null);
    }
  }

  async function apply() {
    if (!pending) return;
    await host.apply(pending.changeset);
    pending.changeset.clear();
    setPending(null);
  }

  return (
    <div className="op-shell">
      <Header
        hostKind={props.hostKind}
        contextLabel={contextLabel}
        onNewChat={() => {
          abortRef.current?.abort();
          setPending(null);
          void persist([], []);
          void clearHistory(props.hostKind);
        }}
        onOpenSettings={() => setTab("settings")}
      />
      {tab === "settings" ? (
        <SettingsPanel
          provider={provider}
          instructions={instructions}
          companion={companion}
          policyNote={policy.catalogUrl ? `Tenant catalog: ${policy.catalogUrl}` : undefined}
          onClose={() => setTab("chat")}
          onChange={async (next) => {
            setProvider(next);
            await saveProvider(next);
          }}
          onInstructions={async (text) => {
            setInstructions(text);
            await saveInstructions(text);
          }}
          onImportUrl={async (url) => {
            const skill = await importSkillFromUrl(url);
            setImported((s) => [...s.filter((x) => x.name !== skill.name), skill]);
          }}
        />
      ) : (
        <>
          {lines.length === 0 ? (
            <EmptyState hostKind={props.hostKind} onPick={(p) => void send(p)} />
          ) : (
            <Thread hostKind={props.hostKind} lines={lines} busy={busy} />
          )}
          {pending && !pending.changeset.isEmpty() && (
            <ReviewCard
              items={pending.changeset.previewItems()}
              onApply={() => void apply()}
              onReject={() => setPending(null)}
              onAlways={() => {
                setAlwaysApply(true);
                void apply();
              }}
            />
          )}
          <Composer
            value={input}
            busy={busy}
            disabled={!provider.baseUrl}
            skills={skills.list(props.hostKind)}
            model={provider.model}
            onChange={setInput}
            onSend={() => void send()}
            onStop={() => abortRef.current?.abort()}
            onInsertSkill={(name) => setInput((v) => `${v}${v ? " " : ""}Use the ${name} skill. `)}
          />
        </>
      )}
    </div>
  );
}

function appendAssistant(lines: StoredLine[], text: string): StoredLine[] {
  const last = lines[lines.length - 1];
  if (last?.kind === "assistant") {
    if (last.text.includes(text)) return lines;
    return [...lines.slice(0, -1), { kind: "assistant", text: last.text || text }];
  }
  return [...lines, { kind: "assistant", text }];
}

function linesRefTail(lines: StoredLine[]): StoredLine[] {
  return lines;
}

function labelFromSnapshot(snap: DocumentSnapshot): string {
  if (snap.host === "excel") {
    const sel = snap.selection as { sheet?: string; address?: string };
    if (sel?.sheet && sel.address) return `${sel.sheet}!${sel.address}`;
  }
  if (snap.host === "powerpoint") {
    const sel = snap.selection as { slideIndex?: number };
    if (typeof sel?.slideIndex === "number") return `Slide ${sel.slideIndex + 1}`;
  }
  if (snap.host === "word") {
    const sel = snap.selection as { text?: string };
    if (sel?.text) return `${Math.min(sel.text.length, 40)} chars selected`;
  }
  return snap.title || snap.host;
}
