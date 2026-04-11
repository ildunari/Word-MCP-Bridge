// @vitest-environment happy-dom

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  WORD_TRACKING_MODE_CHANGED_EVENT,
  buildWordLiveContext,
} from "../src/lib/live-context";
import {
  createWordBridgeAdapter,
  deriveWordTaskpaneRuntimeState,
  isBridgeForcedEnabled,
  resolveBridgeSessionsUrl,
  resolveConfiguredBridgeUrl,
} from "../src/lib/bridge-adapter";
import { resolveWordDocumentTitle } from "../src/lib/document-title";
import { resolveTableDimensions } from "../src/lib/tools";

function makeBridgeStatus(overrides: Record<string, unknown> = {}) {
  return {
    enabled: true,
    serverUrl: "wss://localhost:4017/ws",
    phase: "connecting",
    isConnected: false,
    hasConnected: false,
    lastError: null,
    sessionHealth: "registration_pending",
    details: {
      lastWebSocketOpenAt: null,
      lastHelloSnapshotCapturedAt: null,
      lastHelloSentAt: null,
      lastSessionUpdatedAt: null,
      lastToolCompletedAt: null,
      lastTimedOutInvoke: null,
      activeInvokeCount: 0,
      currentSessionId: null,
      currentDocumentId: null,
      lastDocumentSwitchAt: null,
      lastReconnectAt: null,
    },
    diagnostics: [],
    ...overrides,
  };
}

describe("Word add-in helpers", () => {
  beforeEach(() => {
    window.history.replaceState({}, "", "/taskpane.html");
    const storage = new Map<string, string>();
    vi.stubGlobal("localStorage", {
      getItem: (key: string) => storage.get(key) ?? null,
      setItem: (key: string, value: string) => {
        storage.set(key, value);
      },
      removeItem: (key: string) => {
        storage.delete(key);
      },
    });
  });

  it("builds and truncates live selection context", () => {
    const context = buildWordLiveContext({
      selectedText: "  abcdefghijklmnopqrstuvwxyz  ",
      selectedStyle: "Heading 1",
      trackingMode: "TrackMineOnly",
      focusTarget: "document",
      maxSelectionLength: 10,
      updatedAt: 42,
    });

    expect(context).toEqual({
      selection: {
        hasSelection: true,
        selectedText: "abcdefghij...",
        selectedStyle: "Heading 1",
      },
      trackingMode: "TrackMineOnly",
      focusTarget: "document",
      updatedAt: 42,
    });
  });

  it("resolves bridge url and enable state from query first", () => {
    localStorage.setItem("office-agents-bridge-url", "wss://localhost:9999/ws");
    localStorage.setItem("office-agents-bridge-enabled", "false");
    window.history.replaceState(
      {},
      "",
      "/taskpane.html?office_bridge=true&office_bridge_url=wss://localhost:4018/ws",
    );

    expect(resolveConfiguredBridgeUrl()).toBe("wss://localhost:4018/ws");
    expect(isBridgeForcedEnabled()).toBe(true);
  });

  it("derives the local sessions url from the configured websocket bridge url", () => {
    expect(resolveBridgeSessionsUrl("wss://localhost:4018/ws")).toBe(
      "https://localhost:4018/sessions",
    );
    expect(resolveBridgeSessionsUrl("ws://127.0.0.1:9999/custom")).toBe(
      "http://127.0.0.1:9999/sessions",
    );
  });

  it("exports the tracking mode event constant", () => {
    expect(WORD_TRACKING_MODE_CHANGED_EVENT).toBe(
      "word-tracking-mode-maybe-changed",
    );
  });

  it("ignores the taskpane page title when resolving the Word document title", () => {
    expect(
      resolveWordDocumentTitle({
        documentUrl: null,
        pageTitle: "Word MCP Bridge",
      }),
    ).toBeNull();

    expect(
      resolveWordDocumentTitle({
        documentUrl: "file:///Users/kosta/Documents/Draft.docx",
        pageTitle: "Word MCP Bridge",
      }),
    ).toBe("Draft.docx");
  });

  it("decodes encoded filenames when deriving a Word document title from the document url", () => {
    expect(
      resolveWordDocumentTitle({
        documentUrl: "file:///Users/kosta/Documents/Grant%20Draft%20v2.docx",
        pageTitle: "Word MCP Bridge",
      }),
    ).toBe("Grant Draft v2.docx");
  });

  it("builds a lightweight runtime state from bridge connection status", () => {
    const runtimeState = deriveWordTaskpaneRuntimeState(makeBridgeStatus());

    expect(runtimeState.mode).toBe("connecting");
    expect(runtimeState.taskPhase).toBe("connecting");
    expect(runtimeState.waitingState).toBe("bridge_connection");
    expect(runtimeState.degradedGuardrails).toEqual(["bridge_connection"]);
  });

  it("passes the runtime-state provider through the Word bridge adapter", () => {
    const adapter = createWordBridgeAdapter({
      getRuntimeState: () =>
        deriveWordTaskpaneRuntimeState(
          makeBridgeStatus({
            phase: "connected",
            isConnected: true,
            hasConnected: true,
            sessionHealth: "live",
          }),
        ),
    });

    expect(adapter.getRuntimeState?.()?.mode).toBe("ready");
    expect(adapter.getRuntimeState?.()?.taskPhase).toBe("connected");
  });

  it("falls back to first-row cell count when Word reports zero columns", async () => {
    const sync = vi.fn(async () => undefined);
    const firstRow = {
      cells: {
        items: [{}, {}, {}],
        load: vi.fn(),
      },
    };
    const table = {
      rowCount: 2,
      columnCount: 0,
      load: vi.fn(),
      rows: {
        items: [firstRow, {}],
        load: vi.fn(),
      },
    };

    const dimensions = await resolveTableDimensions({ sync }, table);

    expect(dimensions).toEqual({
      rowCount: 2,
      columnCount: 3,
    });
    expect(sync).toHaveBeenCalledTimes(2);
  });
});
