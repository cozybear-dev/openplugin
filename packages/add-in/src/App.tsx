import {
  assertPolicy,
  Changeset,
  compileSnapshot,
  describeTool,
  listModels,
  LlmError,
  mergePolicy,
  OPEN_POLICY,
  parseSlash,
  runAgent,
  type AgentResult,
  type ChatMessage,
  type DocumentSnapshot,
  type HostKind,
  type ModelInfo,
  type Policy,
  type ProviderConfig,
  type Skill
} from "@openplugin/core";
import { useEffect, useMemo, useRef, useState } from "react";
import { appendAuditLocal, loadAudit, type AuditEntry } from "./audit";
import { bundledSkills } from "./bundled-skills";
import { fetchPolicy, postAudit, probeCompanion, type CompanionStatus } from "./companion";
import { clearHistory, loadHistory, saveHistory, type StoredLine } from "./history";
import { COMPANION_ORIGIN, isOllamaUrl, isOpenRouterUrl } from "./presets";
import { loadRevisions, pushRevision, saveRevisions, type AppliedRevision } from "./revisions";
import { createHost } from "./runtime-host";
import { loadInstructions, loadProvider, loadUserPolicy, saveInstructions, saveProvider } from "./settings";
import {
  importCatalog,
  importSkillFromMarkdown,
  importSkillFromUrl,
  listImportedSkills,
  loadDisabledSkills,
  removeImportedSkill,
  saveDisabledSkills
} from "./skill-store";
import { Composer } from "./ui/Composer";
import { EmptyState } from "./ui/EmptyState";
import { formatError, SettingsPanel } from "./ui/SettingsPanel";
import { Header } from "./ui/Header";
import { AppliedBar, ReviewPanel } from "./ui/ReviewCard";
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
  const [chatModel, setChatModel] = useState(() => loadProvider().model);
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [modelsError, setModelsError] = useState<string | null>(null);
  const [modelsLoading, setModelsLoading] = useState(false);
  const [modelsTick, setModelsTick] = useState(0);
  const [instructions, setInstructions] = useState(loadInstructions);
  const [companion, setCompanion] = useState<CompanionStatus>({ state: "unknown" });
  const [webSearch, setWebSearch] = useState(false);
  const [policy, setPolicy] = useState<Policy>(OPEN_POLICY);
  const [tab, setTab] = useState<"chat" | "settings">("chat");
  const [input, setInput] = useState("");
  const [lines, setLines] = useState<StoredLine[]>(() => loadHistory(props.hostKind).lines);
  const [history, setHistory] = useState<ChatMessage[]>(() => loadHistory(props.hostKind).messages);
  const [busy, setBusy] = useState(false);
  const [pending, setPending] = useState<AgentResult | null>(null);
  const [selectedHunks, setSelectedHunks] = useState<Set<string>>(new Set());
  const [alwaysApply, setAlwaysApply] = useState(false);
  const [revisions, setRevisions] = useState<AppliedRevision[]>(() => loadRevisions());
  const [lastApplied, setLastApplied] = useState<AppliedRevision | null>(null);
  const [disabledSkills, setDisabledSkills] = useState<string[]>(() => loadDisabledSkills());
  const [audit, setAudit] = useState<AuditEntry[]>(() => loadAudit());
  const [contextLabel, setContextLabel] = useState(props.hostKind);
  const [snapshot, setSnapshot] = useState<DocumentSnapshot | undefined>();
  const abortRef = useRef<AbortController | null>(null);

  const enabledSkills = useMemo(
    () => skills.list(props.hostKind).filter((s) => !disabledSkills.includes(s.name)),
    [skills, props.hostKind, disabledSkills]
  );
  const activeModel = chatModel || provider.model;
  const modelOptions = useMemo(
    () => unionModels(models, provider.model, chatModel),
    [models, provider.model, chatModel]
  );

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

  function log(action: AuditEntry["action"], summary: string, extra?: Partial<AuditEntry>) {
    const list = appendAuditLocal({
      action,
      host: props.hostKind,
      model: activeModel,
      summary,
      ...extra
    });
    setAudit(list);
    if (companion.state === "connected") {
      void postAudit(companion.token, list[0]!);
    }
  }

  function effectiveConfig(
    base: ProviderConfig,
    viaCompanion: boolean,
    status: CompanionStatus = companion
  ): ProviderConfig {
    const headers = { ...(base.headers ?? {}) };
    if (isOpenRouterUrl(base.baseUrl)) {
      headers["HTTP-Referer"] ??= "https://openplugin.local";
      headers["X-Title"] ??= "OpenPlugin";
    }
    if (viaCompanion && status.state === "connected") {
      return {
        ...base,
        baseUrl: `${COMPANION_ORIGIN}/v1`,
        headers: {
          ...headers,
          "x-openplugin-token": status.token,
          "x-openplugin-target": base.baseUrl
        }
      };
    }
    return { ...base, headers };
  }

  useEffect(() => {
    const ac = new AbortController();
    void (async () => {
      if (!provider.baseUrl) {
        setModels([]);
        setModelsError(null);
        setModelsLoading(false);
        return;
      }
      setModelsLoading(true);
      setModelsError(null);
      try {
        const fetched = await listModels({
          config: effectiveConfig(provider, false),
          signal: ac.signal
        });
        if (ac.signal.aborted) return;
        setModels(fetched);
      } catch (err) {
        if (ac.signal.aborted) return;
        if (err instanceof LlmError && err.code === "abort") return;
        if (err instanceof LlmError && err.code === "cors") {
          let status = companion;
          if (status.state !== "connected") status = await probeCompanion();
          setCompanion(status);
          if (status.state === "connected") {
            try {
              const fetched = await listModels({
                config: effectiveConfig(provider, true, status),
                signal: ac.signal
              });
              if (ac.signal.aborted) return;
              setModels(fetched);
              setModelsLoading(false);
              return;
            } catch (retryErr) {
              if (ac.signal.aborted) return;
              if (retryErr instanceof LlmError && retryErr.code === "abort") return;
              setModelsError(formatError(retryErr, true));
              setModels([]);
              setModelsLoading(false);
              return;
            }
          }
        }
        setModelsError(formatError(err, companion.state === "connected"));
        setModels([]);
      } finally {
        if (!ac.signal.aborted) setModelsLoading(false);
      }
    })();
    return () => ac.abort();
    // Fetch when the endpoint, key, or companion path changes — not on model id edits.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [provider.baseUrl, provider.apiKey, companion.state, modelsTick]);

  async function send(text = input) {
    const prompt = text.trim();
    if (!prompt || busy) return;

    const slash = parseSlash(prompt, enabledSkills.map((s) => s.name));
    if (slash.kind === "list") {
      setInput("");
      setLines((l) => [
        ...l,
        { kind: "user", text: prompt },
        {
          kind: "assistant",
          text: enabledSkills.length
            ? enabledSkills.map((s) => `/${s.name} — ${s.description}`).join("\n")
            : "No skills enabled for this host."
        }
      ]);
      return;
    }
    if (slash.kind === "unknown") {
      setLines((l) => [...l, { kind: "error", text: `Unknown skill /${slash.name}. Type / to list skills.` }]);
      return;
    }

    if (!provider.baseUrl || !activeModel) {
      setTab("settings");
      return;
    }
    try {
      assertPolicy(policy, { baseUrl: provider.baseUrl, model: activeModel });
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
    const skillPrefix =
      slash.kind === "skill"
        ? `Use the ${slash.name} skill.${slash.rest ? `\n\n${slash.rest}` : ""}`
        : prompt;
    if (slash.kind === "skill") log("skill", `/${slash.name}`, { skill: slash.name });
    const message = instructions.trim()
      ? `${skillPrefix}\n\nStanding instructions:\n${instructions.trim()}`
      : skillPrefix;

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
            const result = await runTurn(message, effectiveConfig(provider, true, status), controller.signal);
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
      config: { ...config, model: chatModel || provider.model },
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
        model: chatModel || provider.model,
        endpoint: provider.baseUrl,
        skill: result.loadedSkills[0],
        toolNames: result.changeset.changes.map((c) => c.op),
        tokenEstimate: result.snapshot.tokenEstimate,
        user: "local"
      });
    }
    if (!result.changeset.isEmpty()) {
      const hunks = result.changeset.diff();
      setSelectedHunks(new Set(hunks.map((h) => h.id)));
      if (alwaysApply) {
        await applyChangeset(result.changeset);
        setPending(null);
      } else {
        setPending(result);
      }
    } else {
      setPending(null);
    }
  }

  async function applyChangeset(cs: Changeset) {
    await host.apply(cs);
    const inverse = cs.inverse();
    const summary = cs.previewItems().map((i) => i.title).join(", ");
    const rev: AppliedRevision = {
      id: `${Date.now()}`,
      at: new Date().toISOString(),
      host: props.hostKind,
      summary,
      inverse: inverse.changes
    };
    const next = pushRevision(revisions, rev);
    setRevisions(next);
    setLastApplied(rev);
    await saveRevisions(next);
    log("apply", summary);
  }

  async function applySelected() {
    if (!pending) return;
    const filtered = pending.changeset.filter(selectedHunks);
    if (filtered.isEmpty()) return;
    await applyChangeset(filtered);
    pending.changeset.clear();
    setPending(null);
  }

  async function revertLast() {
    if (!lastApplied?.inverse.length) return;
    const cs = new Changeset();
    for (const c of lastApplied.inverse) cs.add(c);
    await host.apply(cs);
    log("revert", lastApplied.summary);
    const next = revisions.filter((r) => r.id !== lastApplied.id);
    setRevisions(next);
    setLastApplied(next[0] ?? null);
    await saveRevisions(next);
  }

  return (
    <div className="op-shell">
      <Header
        hostKind={props.hostKind}
        contextLabel={contextLabel}
        onNewChat={() => {
          abortRef.current?.abort();
          setPending(null);
          setChatModel(provider.model);
          setWebSearch(false);
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
          skills={[
            ...bundledSkills()
              .list(props.hostKind)
              .map((s) => ({ ...s, source: "bundled" as const })),
            ...imported
              .filter((s) => s.hosts.includes(props.hostKind))
              .map((s) => ({
                name: s.name,
                description: s.description,
                hosts: s.hosts,
                tools: s.tools,
                source: "imported" as const
              }))
          ]}
          disabledSkills={disabledSkills}
          audit={audit}
          autoApply={alwaysApply}
          models={modelOptions}
          modelsError={modelsError}
          modelsLoading={modelsLoading}
          onClose={() => setTab("chat")}
          onChange={async (next) => {
            const previousDefault = provider.model;
            setProvider(next);
            await saveProvider(next);
            if (chatModel === previousDefault) setChatModel(next.model);
          }}
          onRefreshModels={() => setModelsTick((n) => n + 1)}
          onInstructions={async (text) => {
            setInstructions(text);
            await saveInstructions(text);
          }}
          onImportUrl={async (url) => {
            const skill = await importSkillFromUrl(url);
            setImported((s) => [...s.filter((x) => x.name !== skill.name), skill]);
          }}
          onImportMarkdown={async (md) => {
            const nameMatch = md.match(/^name:\s*([a-z0-9-]+)/m);
            const folder = nameMatch?.[1] ?? "imported-skill";
            const skill = await importSkillFromMarkdown(md, folder);
            setImported((s) => [...s.filter((x) => x.name !== skill.name), skill]);
          }}
          onRemoveSkill={async (name) => {
            await removeImportedSkill(name);
            setImported((s) => s.filter((x) => x.name !== name));
          }}
          onToggleSkill={(name, enabled) => {
            const next = enabled
              ? disabledSkills.filter((n) => n !== name)
              : [...disabledSkills, name];
            setDisabledSkills(next);
            saveDisabledSkills(next);
          }}
          onAutoApply={setAlwaysApply}
          onTestLogged={(ok, summary) => log(ok ? "test" : "error", summary)}
        />
      ) : (
        <>
          {lines.length === 0 ? (
            <EmptyState hostKind={props.hostKind} onPick={(p) => void send(p)} />
          ) : (
            <Thread hostKind={props.hostKind} lines={lines} busy={busy} />
          )}
          {pending && !pending.changeset.isEmpty() && (
            <ReviewPanel
              hostKind={props.hostKind}
              hunks={pending.changeset.diff()}
              selected={selectedHunks}
              onToggle={(id, on) => {
                setSelectedHunks((prev) => {
                  const next = new Set(prev);
                  if (on) next.add(id);
                  else next.delete(id);
                  return next;
                });
              }}
              onApply={() => void applySelected()}
              onReject={() => {
                log("reject", pending.changeset.previewItems().map((i) => i.title).join(", "));
                setPending(null);
              }}
            />
          )}
          {!pending && lastApplied && (
            <AppliedBar
              summary={lastApplied.summary}
              canRevert={lastApplied.inverse.length > 0}
              onRevert={() => void revertLast()}
            />
          )}
          <Composer
            value={input}
            busy={busy}
            disabled={!provider.baseUrl}
            skills={enabledSkills}
            model={chatModel}
            models={modelOptions}
            webSearch={webSearch}
            onChange={setInput}
            onSend={() => void send()}
            onStop={() => abortRef.current?.abort()}
            onInsertSkill={(name) => setInput(`/${name} `)}
            onToggleWebSearch={setWebSearch}
            onModelChange={setChatModel}
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

function unionModels(list: ModelInfo[], ...ids: string[]): ModelInfo[] {
  const byId = new Map<string, ModelInfo>();
  for (const model of list) {
    if (model.id) byId.set(model.id, model);
  }
  for (const id of ids) {
    if (id && !byId.has(id)) byId.set(id, { id });
  }
  return [...byId.values()];
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
