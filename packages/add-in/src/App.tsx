import {
  applyChangeset as applyHostChangeset,
  applyFailureMessage,
  assertPolicy,
  Changeset,
  chatCompletions,
  compileSnapshot,
  extractSkillMarkdown,
  listModels,
  LlmError,
  mergePolicy,
  OPEN_POLICY,
  parseSkillMarkdown,
  parseSlash,
  planRestore,
  resolveRestorePoint,
  applyRestoreUndos,
  runAgent,
  serializeSkill,
  skillifyMessages,
  skillsForAgent,
  stubSkillMarkdown,
  uniqueSkillName,
  type AgentResult,
  type ChatMessage,
  type DocumentSnapshot,
  type HostKind,
  type ModelInfo,
  type Policy,
  type ProviderConfig,
  type RestorePoint,
  type SearchResult
} from "@openplugin/core";
import { useEffect, useMemo, useRef, useState } from "react";
import { appendAuditLocal, documentIdentity, loadAudit, programName, type AuditEntry } from "./audit";
import { bundledSkillList } from "./bundled-skills";
import { fetchPolicy, postAudit, postSearch, probeCompanion, type CompanionStatus } from "./companion";
import { debugLog, subscribeDebug, type DebugEvent } from "./debug";
import {
  applyAutoTitle,
  conversationHasContent,
  deleteConversation,
  emptyConversation,
  listConversations,
  loadActiveId,
  migrateLegacyHistory,
  saveActiveId,
  saveConversation,
  visibleConversations,
  type ConversationRecord
} from "./conversation-store";
import { loadHistory, type StoredLine } from "./history";
import { COMPANION_ORIGIN, isOllamaUrl, isOpenRouterUrl } from "./presets";
import { loadRevisions, pushRevision, saveRevisions, type AppliedRevision } from "./revisions";
import { createHost } from "./runtime-host";
import {
  loadInstructions,
  loadProvider,
  loadSearchSettings,
  loadUserPolicy,
  saveInstructions,
  saveProvider,
  saveSearchSettings
} from "./settings";
import {
  downloadSkillMarkdown,
  importCatalog,
  importSkillFromMarkdown,
  importSkillFromUrl,
  listImportedSkills,
  loadDisabledSkills,
  mergeSkillCatalog,
  mergeStoredRecords,
  removeImportedSkill,
  saveDisabledSkills,
  saveImportedSkill,
  toStoredRecord,
  type StoredSkillRecord
} from "./skill-store";
import { installLiveBridge, type LiveBridgeApi } from "./live-bridge";
import { maxStepsLine, shouldShowTimeout, timeoutLine } from "./turn-status";
import { Composer } from "./ui/Composer";
import { EmptyState } from "./ui/EmptyState";
import { formatError, SettingsPanel } from "./ui/SettingsPanel";
import { Header } from "./ui/Header";
import { HistoryPanel } from "./ui/HistoryPanel";
import { SkillsPanel, type SkillEditorState } from "./ui/SkillsPanel";
import { RestoreDialog } from "./ui/RestoreDialog";
import { AppliedBar, ReviewPanel } from "./ui/ReviewCard";
import { Thread } from "./ui/Thread";

export function App(props: { hostKind: HostKind; inOffice: boolean }) {
  const host = useMemo(() => createHost(props.hostKind, props.inOffice), [props.hostKind, props.inOffice]);
  const [imported, setImported] = useState<StoredSkillRecord[]>([]);
  const managedSkills = useMemo(() => mergeSkillCatalog(bundledSkillList(), imported), [imported]);
  const [disabledSkills, setDisabledSkills] = useState<string[]>(() => loadDisabledSkills());
  const [policy, setPolicy] = useState<Policy>(OPEN_POLICY);
  const skills = useMemo(
    () => skillsForAgent(managedSkills, { disabled: disabledSkills, allowedSkills: policy.allowedSkills }),
    [managedSkills, disabledSkills, policy.allowedSkills]
  );

  const [provider, setProvider] = useState<ProviderConfig>(loadProvider);
  const [chatModel, setChatModel] = useState(() => loadProvider().model);
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [modelsError, setModelsError] = useState<string | null>(null);
  const [modelsLoading, setModelsLoading] = useState(false);
  const [modelsTick, setModelsTick] = useState(0);
  const [instructions, setInstructions] = useState(loadInstructions);
  const [companion, setCompanion] = useState<CompanionStatus>({ state: "unknown" });
  const [tab, setTab] = useState<"chat" | "settings" | "skills" | "history">("chat");
  const [skillEditor, setSkillEditor] = useState<SkillEditorState | null>(null);
  const [input, setInput] = useState("");
  const loaded = useMemo(() => loadHistory(props.hostKind), [props.hostKind]);
  const [lines, setLines] = useState<StoredLine[]>(() => loaded.lines);
  const [history, setHistory] = useState<ChatMessage[]>(() => loaded.messages);
  const [restorePoints, setRestorePoints] = useState<RestorePoint[]>(() => loaded.restorePoints);
  const [conversations, setConversations] = useState<ConversationRecord[]>([]);
  const [activeId, setActiveId] = useState("");
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [searchSettings, setSearchSettings] = useState(loadSearchSettings);
  const [webSearch, setWebSearch] = useState(() => loadSearchSettings().defaultEnabled);
  const [debug, setDebug] = useState<DebugEvent[]>([]);
  const [busy, setBusy] = useState(false);
  const [pending, setPending] = useState<AgentResult | null>(null);
  const [selectedHunks, setSelectedHunks] = useState<Set<string>>(new Set());
  const [alwaysApply, setAlwaysApply] = useState(false);
  const [revisions, setRevisions] = useState<AppliedRevision[]>(() => loadRevisions());
  const [lastApplied, setLastApplied] = useState<AppliedRevision | null>(null);
  const [audit, setAudit] = useState<AuditEntry[]>(() => loadAudit());
  const [contextLabel, setContextLabel] = useState<string>(props.hostKind);
  const [snapshot, setSnapshot] = useState<DocumentSnapshot | undefined>();
  const abortRef = useRef<AbortController | null>(null);
  const operationRef = useRef(0);
  const restoreRef = useRef<RestorePoint[]>(restorePoints);
  restoreRef.current = restorePoints;
  const recordRef = useRef<ConversationRecord | null>(null);
  const [pendingRestore, setPendingRestore] = useState<number | null>(null);

  const slashSkills = useMemo(() => skills.listForSlash(props.hostKind), [skills, props.hostKind]);
  const activeModel = chatModel || provider.model;
  const modelOptions = useMemo(
    () => unionModels(models, provider.model, chatModel),
    [models, provider.model, chatModel]
  );

  useEffect(() => subscribeDebug(setDebug), []);

  useEffect(() => {
    void (async () => {
      try {
        await migrateLegacyHistory(props.hostKind);
        let list = await listConversations(props.hostKind);
        const storedId = loadActiveId(props.hostKind);
        const found = storedId ? list.find((c) => c.id === storedId) : undefined;
        const newest = [...list].sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt))[0];
        const current = found ?? newest ?? emptyConversation(props.hostKind);
        if (!found) {
          if (!newest) {
            await saveConversation(current);
            list = await listConversations(props.hostKind);
            if (!list.some((c) => c.id === current.id)) list = [current, ...list];
          }
          saveActiveId(props.hostKind, current.id);
        }
        if (recordRef.current) {
          setConversations(list);
          setHistoryError(null);
          return;
        }
        recordRef.current = current;
        setActiveId(current.id);
        setConversations(list);
        setLines(current.lines);
        setHistory(current.messages);
        setRestorePoints(current.restorePoints);
        restoreRef.current = current.restorePoints;
        setHistoryError(null);
      } catch {
        setHistoryError("Chats could not be saved.");
      }
    })();
  }, [props.hostKind]);

  useEffect(() => {
    void (async () => {
      const stored = await listImportedSkills().catch(() => []);
      setImported(stored);
      const status = await probeCompanion();
      setCompanion(status);
      if (status.state === "connected") {
        const tenant = (await fetchPolicy(status.token)) as Policy | null;
        if (tenant) {
          const merged = mergePolicy(tenant, loadUserPolicy() ?? undefined);
          setPolicy(merged);
          if (merged.catalogUrl) {
            const extra = await importCatalog(merged.catalogUrl, stored).catch(() => []);
            setImported((s) => mergeStoredRecords(s, extra));
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

  function recordFromState(
    nextLines: StoredLine[],
    nextMessages: ChatMessage[],
    nextPoints: RestorePoint[]
  ): ConversationRecord {
    const identity = documentIdentity();
    const base = recordRef.current ?? emptyConversation(props.hostKind);
    return applyAutoTitle(
      {
        ...base,
        lines: nextLines,
        messages: nextMessages,
        restorePoints: nextPoints,
        updatedAt: new Date().toISOString(),
        documentTitle: identity.documentTitle ?? base.documentTitle,
        documentUrl: identity.documentUrl ?? base.documentUrl
      },
      nextLines
    );
  }

  async function persist(
    nextLines: StoredLine[],
    nextMessages: ChatMessage[],
    nextPoints = restoreRef.current
  ) {
    restoreRef.current = nextPoints;
    setLines(nextLines);
    setHistory(nextMessages);
    setRestorePoints(nextPoints);
    const next = recordFromState(nextLines, nextMessages, nextPoints);
    recordRef.current = next;
    setActiveId(next.id);
    setConversations((list) => [next, ...list.filter((c) => c.id !== next.id)]);
    try {
      await saveConversation(next);
      saveActiveId(props.hostKind, next.id);
      setHistoryError(null);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
  }

  function stopTurn() {
    operationRef.current += 1;
    abortRef.current?.abort();
    setBusy(false);
    abortRef.current = null;
    setPending(null);
  }

  async function startNewChat() {
    stopTurn();
    setChatModel(provider.model);
    setWebSearch(searchSettings.defaultEnabled);
    if (!conversationHasContent(lines)) {
      setTab("chat");
      return;
    }
    try {
      await persist(lines, history, restoreRef.current);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
    const created = emptyConversation(props.hostKind);
    recordRef.current = created;
    saveActiveId(props.hostKind, created.id);
    setActiveId(created.id);
    setLines([]);
    setHistory([]);
    setRestorePoints([]);
    restoreRef.current = [];
    setConversations((list) => [created, ...list.filter((c) => c.id !== created.id)]);
    setTab("chat");
    try {
      await saveConversation(created);
      setHistoryError(null);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
  }

  async function selectConversation(id: string) {
    if (id === activeId) {
      setTab("chat");
      return;
    }
    stopTurn();
    try {
      await persist(lines, history, restoreRef.current);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
    let list = conversations;
    try {
      list = await listConversations(props.hostKind);
      setConversations(list);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
    const next = list.find((c) => c.id === id);
    if (!next) return;
    recordRef.current = next;
    saveActiveId(props.hostKind, next.id);
    setActiveId(next.id);
    setLines(next.lines);
    setHistory(next.messages);
    setRestorePoints(next.restorePoints);
    restoreRef.current = next.restorePoints;
    setTab("chat");
  }

  async function renameConversation(id: string, title: string) {
    const apply = (record: ConversationRecord): ConversationRecord =>
      record.id === id ? { ...record, title, titleSource: "user" as const } : record;
    if (recordRef.current) recordRef.current = apply(recordRef.current);
    setConversations((list) => list.map(apply));
    const target = (recordRef.current?.id === id ? recordRef.current : conversations.find((c) => c.id === id));
    if (!target) return;
    try {
      await saveConversation(apply(target));
      setHistoryError(null);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
  }

  async function removeConversation(id: string) {
    stopTurn();
    try {
      await deleteConversation(id);
      setHistoryError(null);
    } catch {
      setHistoryError("Chats could not be saved.");
    }
    const remaining = conversations.filter((c) => c.id !== id);
    if (id !== activeId) {
      setConversations(remaining);
      return;
    }
    const next =
      remaining.find((c) => conversationHasContent(c.lines)) ?? emptyConversation(props.hostKind);
    recordRef.current = next;
    saveActiveId(props.hostKind, next.id);
    setActiveId(next.id);
    setLines(next.lines);
    setHistory(next.messages);
    setRestorePoints(next.restorePoints);
    restoreRef.current = next.restorePoints;
    setConversations(remaining.some((c) => c.id === next.id) ? remaining : [next, ...remaining]);
    if (!remaining.some((c) => c.id === next.id)) {
      try {
        await saveConversation(next);
      } catch {
        setHistoryError("Chats could not be saved.");
      }
    }
    setTab("chat");
  }

  function log(action: AuditEntry["action"], summary: string, extra?: Partial<AuditEntry>) {
    const list = appendAuditLocal({
      action,
      host: props.hostKind,
      program: programName(props.hostKind),
      model: activeModel,
      endpoint: provider.baseUrl,
      summary,
      ...documentIdentity(),
      documentTitle: snapshot?.title ?? documentIdentity().documentTitle,
      ...extra
    });
    setAudit(list);
    debugLog("ui", `${action}: ${summary}`, { level: action === "error" ? "error" : "info" });
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

    const slash = parseSlash(prompt, slashSkills.map((s) => s.name));
    if (slash.kind === "list") {
      setInput("");
      setLines((l) => [
        ...l,
        { kind: "user", text: prompt },
        {
          kind: "assistant",
          text: slashSkills.length
            ? slashSkills.map((s) => `/${s.name} — ${s.description}`).join("\n")
            : "No slash skills enabled for this host."
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
    const checkpoint = {
      id: `${Date.now()}`,
      at: new Date().toISOString(),
      revisionIds: revisions.map((r) => r.id)
    };
    const userLine: StoredLine = { kind: "user", text: prompt, checkpoint };
    const nextLines = [...lines, userLine];
    const point: RestorePoint = {
      ...checkpoint,
      userLineIndex: nextLines.length - 1,
      messageCount: history.length + 1
    };
    restoreRef.current = [...restoreRef.current, point].slice(-40);
    setRestorePoints(restoreRef.current);
    setLines(nextLines);
    setBusy(true);
    const operation = ++operationRef.current;
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
      await finishTurn(nextLines, result, operation);
    } catch (err) {
      if (operation !== operationRef.current) return;
      if (err instanceof LlmError && err.code === "abort") {
        if (shouldShowTimeout(err, controller.signal.aborted)) {
          setLines((l) => [...l, { kind: "error", text: timeoutLine() }]);
        }
        return;
      }
      if (err instanceof LlmError && err.code === "cors") {
        let status = companion;
        if (status.state !== "connected") status = await probeCompanion();
        setCompanion(status);
        if (status.state === "connected" && !useCompanionFirst) {
          try {
            const result = await runTurn(message, effectiveConfig(provider, true), controller.signal);
            await finishTurn(nextLines, result, operation);
            setBusy(false);
            return;
          } catch (retryErr) {
            if (operation !== operationRef.current) return;
            if (retryErr instanceof LlmError && retryErr.code === "abort") {
              if (shouldShowTimeout(retryErr, controller.signal.aborted)) {
                setLines((l) => [...l, { kind: "error", text: timeoutLine() }]);
              }
              return;
            }
            setLines([...nextLines, { kind: "error", text: formatError(retryErr, true) }]);
            setBusy(false);
            return;
          }
        }
      }
      setLines([...nextLines, { kind: "error", text: formatError(err, companion.state === "connected") }]);
    } finally {
      if (operation === operationRef.current) {
        setBusy(false);
        abortRef.current = null;
      }
    }
  }

  async function runTurn(userMessage: string, config: ProviderConfig, signal: AbortSignal) {
    return runAgent({
      config: { ...config, model: chatModel || provider.model, maxOutputTokens: 8192 },
      host,
      skills,
      userMessage,
      history,
      previousSnapshot: snapshot,
      policy,
      signal,
      webSearch: {
        enabled: webSearch,
        backend: searchSettings.backend,
        exaApiKey: searchSettings.exaApiKey,
        customUrl: searchSettings.custom?.url,
        proxy:
          companion.state === "connected"
            ? async (req) =>
                (await postSearch(companion.token, {
                  ...req,
                  apiKey: searchSettings.exaApiKey,
                  customUrl: searchSettings.custom?.url
                })) as SearchResult
            : undefined
      },
      onEvent: (event) => {
        if (event.type === "text") {
          setLines((l) => {
            const last = l[l.length - 1];
            if (last?.kind === "assistant") {
              return [...l.slice(0, -1), { kind: "assistant", text: last.text + event.delta }];
            }
            return [...l, { kind: "assistant", text: event.delta }];
          });
        } else if (event.type === "activity") {
          debugLog("agent", event.label, { data: { name: event.name, phase: event.phase } });
          setLines((l) => {
            const last = l[l.length - 1];
            if (
              last?.kind === "activity" &&
              last.name === event.name &&
              last.phase === "start" &&
              event.phase !== "start"
            ) {
              return [...l.slice(0, -1), { ...event, kind: "activity" as const }];
            }
            if (event.phase === "start" || event.phase === "error") {
              return [...l, { ...event, kind: "activity" as const }];
            }
            return l;
          });
        } else if (event.type === "error") {
          debugLog("agent", event.message, { level: "error" });
          setLines((l) => [...l, { kind: "error", text: event.message }]);
        }
      }
    });
  }

  async function finishTurn(baseLines: StoredLine[], result: AgentResult, operation = operationRef.current) {
    if (operation !== operationRef.current) return;
    const withFinal = result.finalText
      ? appendAssistant(baseLines, result.finalText)
      : linesRefTail(baseLines);
    const toPersist =
      result.stopReason === "max_steps"
        ? [...withFinal, { kind: "error" as const, text: maxStepsLine(result.steps) }]
        : withFinal;
    const msgs = result.messages.filter(
      (m): m is ChatMessage => m.role === "user" || m.role === "assistant" || m.role === "tool"
    );
    await persist(toPersist, msgs);
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
        const { failed } = await applyChangeset(result.changeset);
        if (failed.length) {
          const remaining = new Changeset();
          for (const r of failed) remaining.add(r.change);
          setPending({ ...result, changeset: remaining });
          setSelectedHunks(new Set(remaining.diff().map((h) => h.id)));
        } else {
          setPending(null);
        }
      } else {
        setPending(result);
      }
    } else {
      setPending(null);
    }
  }

  async function applyChangeset(cs: Changeset) {
    const results = await applyHostChangeset(host, cs);
    const succeeded = results.filter((r) => r.ok);
    const failed = results.filter((r): r is Extract<(typeof results)[number], { ok: false }> => !r.ok);
    if (succeeded.length) {
      const applied = new Changeset();
      for (const r of succeeded) applied.add(r.change);
      const inverse = applied.inverse();
      const summary = applied.previewItems().map((i) => i.title).join(", ");
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
    if (failed.length) {
      const msg = applyFailureMessage(succeeded, failed);
      setLines((l) => [...l, { kind: "error", text: msg.text }]);
      log(msg.kind === "warning" ? "apply" : "error", msg.text);
    }
    return { succeeded, failed };
  }

  async function applySelected() {
    if (!pending) return;
    const filtered = pending.changeset.filter(selectedHunks);
    if (filtered.isEmpty()) return;
    try {
      const { failed } = await applyChangeset(filtered);
      if (!failed.length) {
        pending.changeset.clear();
        setPending(null);
        return;
      }
      const failedChanges = new Set(failed.map((r) => r.change));
      const selectedChanges = new Set(filtered.changes);
      const remaining = new Changeset();
      for (const change of pending.changeset.changes) {
        if (!selectedChanges.has(change) || failedChanges.has(change)) remaining.add(change);
      }
      setPending({ ...pending, changeset: remaining });
      setSelectedHunks(new Set(remaining.diff().map((h) => h.id)));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLines((l) => [...l, { kind: "error", text: message }]);
      log("error", message);
    }
  }

  async function restoreTo(userLineIndex: number) {
    const line = lines[userLineIndex];
    const point = resolveRestorePoint({
      lines,
      messages: history,
      restorePoints,
      userLineIndex,
      lineCheckpoint: line?.kind === "user" ? line.checkpoint : undefined
    });
    abortRef.current?.abort();
    const operation = ++operationRef.current;
    setBusy(false);
    setPending(null);
    const plan = planRestore({
      lines,
      messages: history,
      restorePoint: point,
      revisions
    });
    const keptPoints = restorePoints.filter((p) => p.userLineIndex <= userLineIndex);
    try {
      await persist(plan.lines, plan.messages, keptPoints);
      const undoOutcome = await applyRestoreUndos({
        undos: plan.undos,
        remainingRevisions: plan.remainingRevisions,
        apply: async (rev) => {
          if (operation !== operationRef.current) return [];
          const cs = new Changeset();
          for (const c of rev.inverse) cs.add(c);
          return applyHostChangeset(host, cs);
        }
      });
      if (operation !== operationRef.current) return;
      const restoreErrors = [...undoOutcome.errors];
      if (plan.irreversible.length) {
        restoreErrors.push(
          `Could not auto-undo: ${plan.irreversible.map((r) => r.id).join(", ")}. Use Office Undo if needed.`
        );
      }
      setRevisions(undoOutcome.remainingRevisions);
      setLastApplied(undoOutcome.remainingRevisions[0] ?? null);
      await saveRevisions(undoOutcome.remainingRevisions);
      if (restoreErrors.length) {
        const restoredLines = [
          ...plan.lines,
          ...restoreErrors.map((text) => ({ kind: "error" as const, text }))
        ];
        await persist(restoredLines, plan.messages, keptPoints);
        log("error", restoreErrors.join("; "));
      } else {
        log("restore", `Restored to earlier message`);
      }
    } catch (err) {
      if (operation !== operationRef.current) return;
      const message = err instanceof Error ? err.message : String(err);
      setLines((current) => [...current, { kind: "error", text: message }]);
      log("error", message);
    }
  }

  async function revertLast() {
    if (!lastApplied?.inverse.length) return;
    const cs = new Changeset();
    for (const c of lastApplied.inverse) cs.add(c);
    try {
      const results = await applyHostChangeset(host, cs);
      const failed = results.filter((r) => !r.ok);
      if (failed.length) {
        const detail = failed.map((r) => r.error).join("; ");
        setLines((l) => [...l, { kind: "error", text: `Could not revert: ${detail}` }]);
        log("error", `Revert failed: ${detail}`);
        return;
      }
      log("revert", lastApplied.summary);
      const next = revisions.filter((r) => r.id !== lastApplied.id);
      setRevisions(next);
      setLastApplied(next[0] ?? null);
      await saveRevisions(next);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLines((l) => [...l, { kind: "error", text: message }]);
      log("error", message);
    }
  }

  async function completeText(messages: ChatMessage[]) {
    const useCompanionFirst = isOllamaUrl(provider.baseUrl) && companion.state === "connected";
    try {
      return await chatCompletions({
        config: {
          ...effectiveConfig(provider, useCompanionFirst),
          model: activeModel,
          maxOutputTokens: 2048
        },
        messages,
        stream: false
      });
    } catch (err) {
      if (err instanceof LlmError && err.code === "cors") {
        let status = companion;
        if (status.state !== "connected") status = await probeCompanion();
        setCompanion(status);
        if (status.state === "connected" && !useCompanionFirst) {
          return await chatCompletions({
            config: {
              ...effectiveConfig(provider, true, status),
              model: activeModel,
              maxOutputTokens: 2048
            },
            messages,
            stream: false
          });
        }
      }
      throw err;
    }
  }

  async function persistSkill(markdown: string, previousName?: string) {
    const folder = markdown.match(/^name:\s*([a-z0-9-]+)/m)?.[1] ?? "imported-skill";
    const skill = parseSkillMarkdown(markdown, folder);
    if (previousName && previousName !== skill.name) await removeImportedSkill(previousName);
    const record = toStoredRecord(skill, { markdown, origin: "imported" });
    await saveImportedSkill(record);
    setImported((s) => {
      const without = s.filter((x) => x.name !== skill.name && x.name !== previousName);
      return mergeStoredRecords(without, [record]);
    });
    setSkillEditor(null);
  }

  function newSkill() {
    const name = uniqueSkillName(
      "new-skill",
      managedSkills.map((s) => s.name)
    );
    setSkillEditor({
      markdown: stubSkillMarkdown({ host: props.hostKind, name }),
      mode: "create",
      generating: false,
      error: null
    });
  }

  async function skillifyChat() {
    setTab("skills");
    const taken = managedSkills.map((s) => s.name);
    const name = uniqueSkillName("from-chat", taken);
    const transcript = lines
      .filter((l): l is Extract<StoredLine, { kind: "user" | "assistant" }> => l.kind === "user" || l.kind === "assistant")
      .map((l) => ({ role: l.kind, text: l.text }));
    const notes = transcript
      .slice(-6)
      .map((t) => `${t.role}: ${t.text}`)
      .join("\n");
    setSkillEditor({
      markdown: stubSkillMarkdown({ host: props.hostKind, name, notes }),
      mode: "skillify",
      generating: true,
      error: null
    });
    if (!provider.baseUrl || !activeModel) {
      setSkillEditor((e) =>
        e
          ? { ...e, generating: false, error: "Connect a model to draft from chat, or write the skill yourself." }
          : e
      );
      return;
    }
    try {
      const messages = skillifyMessages({ host: props.hostKind, transcript, existingNames: taken });
      const result = await completeText(messages);
      const content = typeof result.message.content === "string" ? result.message.content : "";
      let md = extractSkillMarkdown(content);
      const parsedName = md.match(/^name:\s*([a-z0-9-]+)/m)?.[1] ?? name;
      const parsed = parseSkillMarkdown(md, parsedName);
      const unique = uniqueSkillName(parsed.name, taken);
      if (unique !== parsed.name) md = serializeSkill({ ...parsed, name: unique, rootPath: unique });
      setSkillEditor({ markdown: md, mode: "skillify", generating: false, error: null });
    } catch (err) {
      setSkillEditor((e) =>
        e ? { ...e, generating: false, error: formatError(err, companion.state === "connected") } : e
      );
    }
  }

  const liveApi = useRef<LiveBridgeApi | null>(null);
  liveApi.current = {
    getState: () => {
      const last = lines[lines.length - 1];
      return {
        hostKind: props.hostKind,
        inOffice: props.inOffice,
        tab,
        busy,
        input,
        provider: { baseUrl: provider.baseUrl, model: activeModel },
        pending: pending ? pending.changeset.diff().map((h) => ({ id: h.id, title: h.title })) : [],
        lastLine: last && "text" in last ? last.text : undefined
      };
    },
    configure: async (next) => {
      const config = { ...provider, ...next };
      setProvider(config);
      setChatModel(next.model || config.model);
      await saveProvider(config);
    },
    send: (text) => send(text),
    apply: () => applySelected(),
    reject: () => setPending(null),
    setTab: (next) => setTab(next)
  };

  useEffect(() => {
    return installLiveBridge({
      getState: () => liveApi.current!.getState(),
      configure: (next) => liveApi.current!.configure(next),
      send: (text) => liveApi.current!.send(text),
      apply: () => liveApi.current!.apply(),
      reject: () => liveApi.current!.reject(),
      setTab: (next) => liveApi.current!.setTab(next)
    });
  }, []);

  return (
    <div
      className="op-shell"
      data-testid="openplugin-app"
      data-host={props.hostKind}
      data-in-office={props.inOffice ? "1" : "0"}
    >
      <Header
        hostKind={props.hostKind}
        contextLabel={contextLabel}
        onNewChat={() => void startNewChat()}
        onOpenHistory={() => setTab("history")}
        onOpenSkills={() => {
          setSkillEditor(null);
          setTab("skills");
        }}
        onOpenSettings={() => setTab("settings")}
      />
      {tab === "history" ? (
        <HistoryPanel
          conversations={visibleConversations(conversations, activeId)}
          activeId={activeId}
          error={historyError}
          onClose={() => setTab("chat")}
          onSelect={(id) => void selectConversation(id)}
          onRename={(id, title) => void renameConversation(id, title)}
          onDelete={(id) => void removeConversation(id)}
        />
      ) : tab === "settings" ? (
        <SettingsPanel
          provider={provider}
          instructions={instructions}
          companion={companion}
          policyNote={policy.catalogUrl ? `Tenant catalog: ${policy.catalogUrl}` : undefined}
          audit={audit}
          autoApply={alwaysApply}
          search={searchSettings}
          debug={debug}
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
          onOpenSkills={() => {
            setSkillEditor(null);
            setTab("skills");
          }}
          onAutoApply={setAlwaysApply}
          onSearch={async (next) => {
            setSearchSettings(next);
            await saveSearchSettings(next);
          }}
          onTestLogged={(ok, summary) => log(ok ? "test" : "error", summary)}
        />
      ) : tab === "skills" ? (
        <SkillsPanel
          skills={managedSkills}
          disabledSkills={disabledSkills}
          editor={skillEditor}
          onClose={() => {
            setSkillEditor(null);
            setTab("chat");
          }}
          onToggle={(name, enabled) => {
            const next = enabled ? disabledSkills.filter((n) => n !== name) : [...disabledSkills, name];
            setDisabledSkills(next);
            saveDisabledSkills(next);
          }}
          onRemove={async (name) => {
            await removeImportedSkill(name);
            setImported((s) => s.filter((x) => x.name !== name));
          }}
          onExport={(skill) => downloadSkillMarkdown(skill.name, skill.markdown)}
          onOpenEditor={setSkillEditor}
          onChangeEditor={(markdown) => setSkillEditor((e) => (e ? { ...e, markdown } : e))}
          onSaveEditor={() => {
            if (!skillEditor) return;
            void persistSkill(skillEditor.markdown, skillEditor.previousName).catch((err) => {
              setSkillEditor((e) =>
                e ? { ...e, error: err instanceof Error ? err.message : String(err) } : e
              );
            });
          }}
          onCloseEditor={() => setSkillEditor(null)}
          onImportUrl={async (url) => {
            const skill = await importSkillFromUrl(url);
            setImported((s) => mergeStoredRecords(s, [skill]));
          }}
          onImportMarkdown={async (md) => {
            const folder = md.match(/^name:\s*([a-z0-9-]+)/m)?.[1] ?? "imported-skill";
            const skill = await importSkillFromMarkdown(md, folder);
            setImported((s) => mergeStoredRecords(s, [skill]));
          }}
          onNew={newSkill}
        />
      ) : (
        <>
          <div className="op-document-bar"><span className="op-document-label">Selection</span><span title={contextLabel}>{contextLabel}</span><span className="op-review-mode">{alwaysApply ? "Auto-apply on" : "Review before apply"}</span></div>
          {lines.length === 0 ? (
            <EmptyState hostKind={props.hostKind} onPick={(p) => {
              setInput(p);
              requestAnimationFrame(() => document.querySelector<HTMLTextAreaElement>(".op-composer-input")?.focus());
            }} />
          ) : (
            <Thread
              key={activeId}
              hostKind={props.hostKind}
              lines={lines}
              busy={busy}
              onRestore={setPendingRestore}
            />
          )}
          <RestoreDialog
            open={pendingRestore !== null}
            onCancel={() => setPendingRestore(null)}
            onConfirm={() => {
              const index = pendingRestore;
              setPendingRestore(null);
              if (index !== null) void restoreTo(index);
            }}
          />
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
          {(!provider.baseUrl || !activeModel) && <button className="op-connect" onClick={() => setTab("settings")}><span><strong>Connect your model</strong><span>Choose a provider to start working together.</span></span><span aria-hidden>→</span></button>}
          <Composer
            value={input}
            busy={busy}
            disabled={false}
            skills={slashSkills}
            model={chatModel}
            models={modelOptions}
            webSearch={webSearch}
            onChange={setInput}
            onSend={() => void send()}
            onStop={() => abortRef.current?.abort()}
            onInsertSkill={(name) => setInput(`/${name} `)}
            onToggleWebSearch={setWebSearch}
            onModelChange={setChatModel}
            onManageSkills={() => {
              setSkillEditor(null);
              setTab("skills");
            }}
            onSkillify={() => void skillifyChat()}
            canSkillify={lines.some((l) => l.kind === "user" || l.kind === "assistant")}
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
