import { describe, expect, it } from "vitest";
import { deriveTaskpaneConnectionView } from "../src/lib/taskpane-connection";

describe("deriveTaskpaneConnectionView", () => {
  it("shows reconnecting when Word data exists but the bridge socket is down", () => {
    const view = deriveTaskpaneConnectionView({
      snapshot: {
        sessionId: "word:demo",
        instanceId: "instance-1",
        app: "word",
        documentId: "doc-1",
        tools: [],
        host: {},
        connectedAt: 1,
        updatedAt: 2,
      },
      bridgeStatus: {
        enabled: true,
        serverUrl: "wss://localhost:4017/ws",
        phase: "reconnecting",
        isConnected: false,
        hasConnected: true,
        lastError: { message: "Could not connect to the bridge server.", at: 3 },
      },
    });

    expect(view.headline).toBe("Word is ready");
    expect(view.statusLabel).toBe("Reconnecting");
    expect(view.subtitle).toContain("trying to reconnect");
  });

  it("shows waiting for Word when the bridge is up but no document snapshot exists", () => {
    const view = deriveTaskpaneConnectionView({
      snapshot: null,
      bridgeStatus: {
        enabled: true,
        serverUrl: "wss://localhost:4017/ws",
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        lastError: null,
      },
    });

    expect(view.headline).toBe("Waiting for Word");
    expect(view.statusLabel).toBe("Bridge connected");
    expect(view.subtitle).toContain("Open the Word MCP Bridge taskpane");
  });
});
