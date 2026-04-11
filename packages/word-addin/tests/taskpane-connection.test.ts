import { describe, expect, it } from "vitest";
import {
  deriveTaskpaneConnectionView,
  deriveTaskpaneDashboardView,
} from "../src/lib/taskpane-connection";

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

describe("deriveTaskpaneDashboardView", () => {
  it("shows a connected document dashboard for a ready single document", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Grant Draft.docx",
          url: "file:///Users/kosta/Documents/Grant%20Draft.docx",
          trackingMode: "TrackAll",
          wordCount: 1432,
          paragraphCount: 18,
          selectionLength: 24,
        },
        host: {},
        gateway: {
          liveContext: {
            selection: {
              hasSelection: true,
              selectedText: "This paragraph needs stronger evidence.",
            },
            trackingMode: "TrackAll",
            focusTarget: "document",
            updatedAt: 2,
          },
        },
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        sessionHealth: "live",
      }),
      serverSessionRegistered: true,
      connectedSessionCount: 1,
    });

    expect(view.statusLabel).toBe("Connected");
    expect(view.documentCard.title).toBe("Grant Draft.docx");
    expect(view.documentCard.summary).toContain("Saved Word document");
    expect(view.selectionCard.title).toBe("Selection ready");
    expect(view.readinessCard.items[0]).toContain("Current document text");
    expect(view.readinessCard.warning).toBeUndefined();
  });

  it("shows connecting guidance without leaking low-level bridge wording", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: null,
      bridgeStatus: makeBridgeStatus({
        lastError: { message: "Constructor blocked", at: 4 },
      }),
    });

    expect(view.statusLabel).toBe("Connecting");
    expect(view.subtitle).toContain("attaching to the local Word system");
    expect(view.subtitle).not.toContain("websocket");
    expect(view.recoveryText).toContain("open the helper app");
  });

  it("does not present a stale server-registered session as fully connected", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Draft.docx",
          trackingMode: "Off",
        },
        host: {},
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        hasConnected: true,
      }),
      serverSessionRegistered: true,
      connectedSessionCount: 1,
    });

    expect(view.statusLabel).toBe("Connecting");
    expect(view.headline).toContain("opening");
  });

  it("shows reconnect guidance when a document is available locally but the bridge dropped", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "word-local:demo",
        documentMetadata: {
          title: "Word MCP Bridge",
          trackingMode: "Off",
          wordCount: 0,
          paragraphCount: 1,
        },
        host: {},
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "reconnecting",
        hasConnected: true,
        lastError: { message: "Could not connect to the bridge server.", at: 3 },
        sessionHealth: "reconnecting",
      }),
    });

    expect(view.statusLabel).toBe("Reconnect needed");
    expect(view.documentCard.title).toBe("Untitled Word document");
    expect(view.recoveryTitle).toBe("Reconnect this document");
    expect(view.showHelperHint).toBe(true);
  });

  it("shows a multiple-documents warning without exposing raw selectors", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Paper Draft.docx",
          trackingMode: "TrackMineOnly",
        },
        host: {},
        gateway: {
          liveContext: {
            selection: { hasSelection: false },
            trackingMode: "TrackMineOnly",
            updatedAt: 2,
          },
        },
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        sessionHealth: "live",
      }),
      serverSessionRegistered: true,
      connectedSessionCount: 2,
    });

    expect(view.statusLabel).toBe("Multiple documents connected");
    expect(view.readinessCard.warning).toContain("helper app");
    expect(view.subtitle).not.toContain("session");
  });

  it("stays in connecting when the bridge is up but this pane has not matched a live server session yet", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Draft A.docx",
          trackingMode: "Off",
        },
        host: {},
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        sessionHealth: "live",
      }),
      serverSessionRegistered: false,
      connectedSessionCount: 2,
    });

    expect(view.statusLabel).toBe("Connecting");
    expect(view.headline).toContain("opening");
    expect(view.subtitle).toContain("matching itself");
  });

  it("does not present stale document state as ready when registration sync is in error", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Draft A.docx",
          trackingMode: "Off",
        },
        host: {},
        tools: [],
        connectedAt: 1,
        updatedAt: 5,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        sessionHealth: "live",
      }),
      serverSessionRegistered: false,
      connectedSessionCount: 0,
      lastRegistrationSyncAt: 10,
      lastRegistrationSyncState: "error",
    });

    expect(view.statusLabel).toBe("Connecting");
    expect(view.readinessCard.title).toContain("Waiting");
    expect(view.connectionCard.items).toContain("Registration state: Error");
  });

  it("surfaces precise reconnect diagnostics for taskpane races", () => {
    const view = deriveTaskpaneDashboardView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        documentMetadata: {
          title: "Draft A.docx",
          trackingMode: "Off",
        },
        host: {},
        tools: [],
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: makeBridgeStatus({
        phase: "reconnecting",
        hasConnected: true,
        lastError: { message: "Could not connect to the bridge server.", at: 9 },
        sessionHealth: "reconnecting",
        diagnostics: [
          { level: "info", message: "Helper poll tick completed.", at: 1 },
          { level: "info", message: "Taskpane focus refresh started.", at: 2 },
          { level: "warn", message: "Taskpane reported stale bundle generation 7; reloading.", at: 3 },
          { level: "warn", message: "WebSocket closed for wss://localhost:4017/ws code=1006", at: 4 },
        ],
      }),
    });

    expect(view.recoveryText).toContain("stale taskpane bundle was reloaded");
    expect(view.recoveryText).toContain("still trying to reattach");
  });
});

describe("deriveTaskpaneConnectionView", () => {
  it("keeps the compact connection view aligned with the dashboard view", () => {
    const view = deriveTaskpaneConnectionView({
      snapshot: null,
      bridgeStatus: makeBridgeStatus({
        phase: "disconnected",
        hasConnected: true,
        sessionHealth: "reconnecting",
      }),
    });

    expect(view.statusLabel).toBe("Reconnect needed");
    expect(view.subtitle).toContain("lost contact");
  });
});
