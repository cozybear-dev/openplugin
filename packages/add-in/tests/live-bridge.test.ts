import { describe, expect, it } from "vitest";
import { installLiveBridge, type LiveBridgeApi } from "../src/live-bridge";

const stub: LiveBridgeApi = {
  getState: () => ({
    hostKind: "excel",
    inOffice: true,
    tab: "chat",
    busy: false,
    input: "",
    provider: { baseUrl: "http://127.0.0.1:8790/v1", model: "openplugin-live" },
    pending: []
  }),
  configure: async () => undefined,
  send: async () => undefined,
  apply: async () => undefined,
  reject: () => undefined,
  setTab: () => undefined
};

describe("installLiveBridge", () => {
  it("exposes the API on window in DEV", () => {
    const uninstall = installLiveBridge(stub);
    expect(window.__openpluginLive?.getState().hostKind).toBe("excel");
    uninstall();
    expect(window.__openpluginLive).toBeUndefined();
  });
});
