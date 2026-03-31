// @vitest-environment happy-dom

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  WORD_TRACKING_MODE_CHANGED_EVENT,
  buildWordLiveContext,
} from "../src/lib/live-context";
import {
  isBridgeForcedEnabled,
  resolveConfiguredBridgeUrl,
} from "../src/lib/bridge-adapter";

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

  it("exports the tracking mode event constant", () => {
    expect(WORD_TRACKING_MODE_CHANGED_EVENT).toBe(
      "word-tracking-mode-maybe-changed",
    );
  });
});
