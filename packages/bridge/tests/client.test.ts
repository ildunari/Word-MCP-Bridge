// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest";
import { startOfficeBridge } from "../src/client";

const adapter = {
  tools: [],
  getDocumentId: async () => "doc-1",
};

class FakeWebSocket {
  static OPEN = 1;
  static instances: FakeWebSocket[] = [];
  readyState = FakeWebSocket.OPEN;
  url: string;
  listeners = new Map<string, Set<(event?: unknown) => void>>();
  sent: string[] = [];

  constructor(url: string) {
    this.url = url;
    FakeWebSocket.instances.push(this);
    queueMicrotask(() => this.emit("open"));
  }

  addEventListener(type: string, listener: (event?: unknown) => void) {
    const listeners = this.listeners.get(type) ?? new Set();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  send(message: string) {
    this.sent.push(message);
  }

  close() {
    this.emit("close");
  }

  emit(type: string, event?: unknown) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener(event);
    }
  }
}

describe("bridge client", () => {
  afterEach(() => {
    FakeWebSocket.instances = [];
    delete (window as typeof window & { __OFFICE_BRIDGE__?: unknown })
      .__OFFICE_BRIDGE__;
    vi.unstubAllGlobals();
  });

  it("removes the exposed global controller on stop", () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
    });

    expect(
      (window as typeof window & { __OFFICE_BRIDGE__?: unknown })
        .__OFFICE_BRIDGE__,
    ).toBe(controller);

    controller.stop();

    expect(
      (window as typeof window & { __OFFICE_BRIDGE__?: unknown })
        .__OFFICE_BRIDGE__,
    ).toBeUndefined();
  });

  it("includes gateway capabilities and live context in the hello snapshot", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: {
        ...adapter,
        getCapabilities: () => ["observe", "tool_call", "unsafe_office_js"],
        getLiveContext: () => ({
          selection: {
            hasSelection: true,
            selectedText: "Selected clause",
            selectedStyle: "Heading 1",
          },
          trackingMode: "TrackMineOnly",
          focusTarget: "document",
          updatedAt: 123,
        }),
      },
      enabled: true,
    });

    await vi.waitFor(() => {
      expect(FakeWebSocket.instances[0]?.sent.length ?? 0).toBeGreaterThanOrEqual(
        1,
      );
    });

    const socket = (window as typeof window & { __OFFICE_BRIDGE__?: unknown })
      .__OFFICE_BRIDGE__;
    expect(socket).toBe(controller);

    const messages =
      FakeWebSocket.instances[0]?.sent.map((message) => JSON.parse(message)) ?? [];
    expect(messages.map((message) => message.type)).toContain("hello");
    const hello = messages.find((message) => message.type === "hello");

    expect(hello?.snapshot.gateway?.capabilities).toEqual([
      "observe",
      "tool_call",
      "unsafe_office_js",
    ]);
    expect(hello?.snapshot.gateway?.liveContext).toEqual({
      selection: {
        hasSelection: true,
        selectedText: "Selected clause",
        selectedStyle: "Heading 1",
      },
      trackingMode: "TrackMineOnly",
      focusTarget: "document",
      updatedAt: 123,
    });
    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("Hello sent for"),
      ),
    ).toBe(true);

    controller.stop();
  });

  it("rejects tool execution when the required capability is missing", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const execute = vi.fn();
    const controller = startOfficeBridge({
      app: "word",
      adapter: {
        tools: [
          {
            name: "mutate_document",
            description: "Mutate the active document",
            requiredCapability: "document_edit",
            execute,
          },
        ],
        getDocumentId: async () => "doc-1",
        getCapabilities: () => ["observe", "tool_call"],
      },
      enabled: true,
    });

    await vi.waitFor(() => {
      expect(FakeWebSocket.instances[0]?.sent.length ?? 0).toBeGreaterThanOrEqual(
        1,
      );
    });

    FakeWebSocket.instances[0]?.emit("message", {
      data: JSON.stringify({
        type: "invoke",
        requestId: "req-1",
        method: "execute_tool",
        params: { toolName: "mutate_document", args: {} },
      }),
    });

    await vi.waitFor(() => {
      const response = FakeWebSocket.instances[0]?.sent
        .map((message) => JSON.parse(message))
        .find(
          (message) => message.type === "response" && message.requestId === "req-1",
        );
      expect(response?.ok).toBe(false);
    });

    const response = FakeWebSocket.instances[0]?.sent
      .map((message) => JSON.parse(message))
      .find((message) => message.type === "response" && message.requestId === "req-1");

    expect(execute).not.toHaveBeenCalled();
    expect(response?.ok).toBe(false);
    expect(response?.error?.message).toContain("document_edit");

    controller.stop();
  });

  it("rejects unsafe Office.js execution when that capability is missing", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: {
        ...adapter,
        getCapabilities: () => ["observe", "tool_call"],
      },
      enabled: true,
    });

    await vi.waitFor(() => {
      expect(FakeWebSocket.instances[0]?.sent.length ?? 0).toBeGreaterThanOrEqual(
        1,
      );
    });

    FakeWebSocket.instances[0]?.emit("message", {
      data: JSON.stringify({
        type: "invoke",
        requestId: "req-2",
        method: "execute_unsafe_office_js",
        params: { code: "return 1;" },
      }),
    });

    await vi.waitFor(() => {
      const response = FakeWebSocket.instances[0]?.sent
        .map((message) => JSON.parse(message))
        .find(
          (message) => message.type === "response" && message.requestId === "req-2",
        );
      expect(response?.ok).toBe(false);
    });

    const response = FakeWebSocket.instances[0]?.sent
      .map((message) => JSON.parse(message))
      .find((message) => message.type === "response" && message.requestId === "req-2");

    expect(response?.ok).toBe(false);
    expect(response?.error?.message).toContain("unsafe_office_js");

    controller.stop();
  });

  it("reports reconnecting bridge state after a socket error", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
      reconnectBaseMs: 50,
    });

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("connected");
    });

    FakeWebSocket.instances[0]?.emit("error");

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("reconnecting");
    });

    expect(controller.getStatus().lastError?.message).toContain(
      "Could not connect to the bridge server",
    );
    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("WebSocket error while connecting"),
      ),
    ).toBe(true);
    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("WebSocket closed"),
      ),
    ).toBe(true);
    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("Scheduling websocket reconnect"),
      ),
    ).toBe(true);

    controller.stop();
  });

  it("records a diagnostic when the WebSocket constructor throws", async () => {
    class ThrowingWebSocket {
      static OPEN = 1;
      readyState = 0;

      constructor() {
        throw new Error("Constructor blocked");
      }
    }

    vi.stubGlobal("WebSocket", ThrowingWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
      reconnectBaseMs: 50,
    });

    await vi.waitFor(() => {
      expect(controller.getStatus().lastError?.message).toContain(
        "Constructor blocked",
      );
    });

    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("WebSocket constructor failed"),
      ),
    ).toBe(true);

    controller.stop();
  });

  it("retries when the initial bridge snapshot fails after the socket opens", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const getDocumentId = vi
      .fn<() => Promise<string>>()
      .mockRejectedValueOnce(new Error("Word host not ready"))
      .mockResolvedValue("doc-1");
    const controller = startOfficeBridge({
      app: "word",
      adapter: {
        tools: [],
        getDocumentId,
      },
      enabled: true,
      reconnectBaseMs: 10,
      reconnectMaxMs: 10,
    });

    await vi.waitFor(() => {
      expect(FakeWebSocket.instances.length).toBeGreaterThanOrEqual(2);
      expect(controller.getStatus().phase).toBe("connected");
    });

    expect(getDocumentId).toHaveBeenCalledTimes(2);
    expect(controller.getStatus().lastError).toBeNull();
    expect(
      controller.getStatus().diagnostics.some((entry) =>
        entry.message.includes("Connected socket could not build a bridge snapshot"),
      ),
    ).toBe(true);

    controller.stop();
  });

  it("ignores close events from stale sockets after a newer connection wins", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
      reconnectBaseMs: 10,
      reconnectMaxMs: 10,
    });

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("connected");
    });

    const firstSocket = FakeWebSocket.instances[0];
    firstSocket?.emit("error");

    await vi.waitFor(() => {
      expect(FakeWebSocket.instances.length).toBeGreaterThanOrEqual(2);
      expect(controller.getStatus().phase).toBe("connected");
    });

    firstSocket?.emit("close", { code: 1006, reason: "stale close" });

    expect(controller.getStatus().phase).toBe("connected");

    controller.stop();
  });

  it("records diagnostics that distinguish helper polls, taskpane refreshes, stale reloads, and socket drops", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
      reconnectBaseMs: 10,
      reconnectMaxMs: 10,
    });

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("connected");
    });

    controller.emitEvent("bridge_status", {
      status: "helper_poll",
      source: "mac-helper",
    } as never);
    controller.emitEvent("bridge_status", {
      status: "taskpane_refresh",
      source: "focus",
    } as never);
    window.dispatchEvent(new Event("focus"));
    window.dispatchEvent(new Event("beforeunload"));
    FakeWebSocket.instances[0]?.emit("error");

    await vi.waitFor(() => {
      const messages = controller.getStatus().diagnostics.map((entry) => entry.message);
      expect(messages.some((message) => message.includes("Helper poll"))).toBe(true);
      expect(messages.some((message) => message.includes("Taskpane focus refresh"))).toBe(true);
      expect(messages.some((message) => message.includes("stale page reload"))).toBe(true);
      expect(messages.some((message) => message.includes("WebSocket closed"))).toBe(true);
    });

    controller.stop();
  });

  it("notifies subscribers when the bridge connection phase changes", async () => {
    vi.stubGlobal("WebSocket", FakeWebSocket);
    const controller = startOfficeBridge({
      app: "word",
      adapter: { ...adapter },
      enabled: true,
      reconnectBaseMs: 50,
    });
    const observedPhases: string[] = [];

    const unsubscribe = controller.subscribe((status) => {
      observedPhases.push(status.phase);
    });

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("connected");
    });

    FakeWebSocket.instances[0]?.emit("error");

    await vi.waitFor(() => {
      expect(controller.getStatus().phase).toBe("reconnecting");
    });

    unsubscribe();
    controller.stop();

    expect(observedPhases).toContain("connecting");
    expect(observedPhases).toContain("connected");
    expect(observedPhases).toContain("reconnecting");
  });
});
