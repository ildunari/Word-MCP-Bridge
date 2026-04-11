// @vitest-environment happy-dom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const startOfficeBridge = vi.fn();
const bindOfficeDocumentHandler = vi.fn(() => () => undefined);
const attachWordLiveContextBridge = vi.fn(() => () => undefined);

vi.mock("@word-mcp-bridge/bridge/client", () => ({
  startOfficeBridge,
}));

vi.mock("../src/lib/components/office-document-events", () => ({
  bindOfficeDocumentHandler,
}));

vi.mock("../src/lib/live-context", () => ({
  attachWordLiveContextBridge,
  WORD_TRACKING_MODE_CHANGED_EVENT: "word-tracking-mode-maybe-changed",
}));

function makeSnapshot() {
  return {
    sessionId: "word:instance-1",
    instanceId: "instance-1",
    app: "word",
    documentId: "doc-1",
    documentMetadata: { title: "Draft.docx" },
    host: {},
    tools: [],
    connectedAt: 1,
    updatedAt: 2,
    gateway: {
      liveContext: {
        selection: { hasSelection: false },
        updatedAt: 2,
      },
    },
  };
}

function makeStatus() {
  return {
    enabled: true,
    serverUrl: "wss://localhost:4017/ws",
    phase: "connected",
    isConnected: true,
    hasConnected: true,
    lastError: null,
    sessionHealth: "live",
    details: {
      lastWebSocketOpenAt: null,
      lastHelloSnapshotCapturedAt: null,
      lastHelloSentAt: null,
      lastSessionUpdatedAt: null,
      lastToolCompletedAt: null,
      lastTimedOutInvoke: null,
      activeInvokeCount: 0,
      currentSessionId: "word:instance-1",
      currentDocumentId: "doc-1",
      lastDocumentSwitchAt: null,
      lastReconnectAt: null,
    },
    diagnostics: [],
  };
}

describe("shared runtime lifecycle", () => {
  let visibilityHandler: ((args: { visibilityMode?: string }) => void) | null = null;
  let stop = vi.fn();
  let refresh = vi.fn(async () => makeSnapshot());

  beforeEach(() => {
    vi.resetModules();
    stop = vi.fn();
    refresh = vi.fn(async () => makeSnapshot());
    visibilityHandler = null;
    startOfficeBridge.mockReset();
    bindOfficeDocumentHandler.mockClear();
    attachWordLiveContextBridge.mockClear();
    vi.stubGlobal("fetch", vi.fn(async () => ({
      ok: true,
      json: async () => ({
        sessions: [
          {
            pendingCount: 0,
            health: "live",
            snapshot: { ...makeSnapshot(), app: "word" },
          },
        ],
      }),
    })));
    (globalThis as any).Office = {
      context: {
        document: {
          addHandlerAsync: vi.fn(),
          removeHandlerAsync: vi.fn(),
          url: "",
        },
      },
      EventType: {
        DocumentSelectionChanged: "DocumentSelectionChanged",
      },
      StartupBehavior: {
        load: "load",
      },
      addin: {
        hide: vi.fn(async () => undefined),
        showAsTaskpane: vi.fn(async () => undefined),
        setStartupBehavior: vi.fn(async () => undefined),
        onVisibilityModeChanged: vi.fn(async (handler: typeof visibilityHandler) => {
          visibilityHandler = handler;
          return async () => {
            visibilityHandler = null;
          };
        }),
      },
    };
    startOfficeBridge.mockReturnValue({
      enabled: true,
      instanceId: "instance-1",
      refresh,
      reconnect: vi.fn(async () => makeSnapshot()),
      emitEvent: vi.fn(),
      getStatus: () => makeStatus(),
      subscribe: vi.fn(() => () => undefined),
      stop,
    });
  });

  afterEach(async () => {
    const mod = await import("../src/lib/shared-runtime");
    mod.__resetWordTaskpaneRuntimeForTests();
    vi.unstubAllGlobals();
  });

  it("starts the bridge runtime only once", async () => {
    const mod = await import("../src/lib/shared-runtime");

    await mod.ensureWordTaskpaneRuntime();
    await mod.ensureWordTaskpaneRuntime();

    expect(startOfficeBridge).toHaveBeenCalledTimes(1);
    expect(refresh).toHaveBeenCalledTimes(1);
  });

  it("keeps the bridge alive when UI subscribers detach", async () => {
    const mod = await import("../src/lib/shared-runtime");

    await mod.ensureWordTaskpaneRuntime();
    const unsubscribe = mod.subscribeWordTaskpaneRuntime(() => undefined);
    unsubscribe();

    expect(stop).not.toHaveBeenCalled();
  });

  it("tracks hidden visibility without resetting the active session", async () => {
    const mod = await import("../src/lib/shared-runtime");

    await mod.ensureWordTaskpaneRuntime();
    visibilityHandler?.({ visibilityMode: "Hidden" });

    const snapshot = mod.getWordTaskpaneRuntimeSnapshot();
    expect(snapshot.paneVisibility).toBe("hidden");
    expect(snapshot.snapshot?.sessionId).toBe("word:instance-1");
  });
});
