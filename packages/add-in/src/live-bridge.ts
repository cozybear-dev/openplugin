import type { ProviderConfig } from "@openplugin/core";

export type LiveTab = "chat" | "settings" | "skills" | "history";

export type LiveStateDump = {
  hostKind: string;
  inOffice: boolean;
  tab: LiveTab;
  busy: boolean;
  input: string;
  provider: { baseUrl: string; model: string };
  pending: Array<{ id: string; title: string }>;
  lastLine?: string;
};

export type LiveBridgeApi = {
  getState: () => LiveStateDump;
  configure: (provider: Pick<ProviderConfig, "baseUrl" | "model" | "apiKey">) => Promise<void>;
  send: (text: string) => Promise<void>;
  apply: () => Promise<void>;
  reject: () => void;
  setTab: (tab: LiveTab) => void;
};

declare global {
  interface Window {
    __openpluginLive?: LiveBridgeApi;
  }
}

export function installLiveBridge(api: LiveBridgeApi): () => void {
  if (!import.meta.env.DEV) return () => undefined;
  window.__openpluginLive = api;
  return () => {
    if (window.__openpluginLive === api) delete window.__openpluginLive;
  };
}
