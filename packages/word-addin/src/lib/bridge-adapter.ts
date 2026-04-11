import type { OfficeBridgeConnectionStatus } from "@word-mcp-bridge/bridge/client";
import type { BridgeCapability } from "@word-mcp-bridge/bridge/protocol";
import type { BridgeRuntimeStateSlice } from "@word-mcp-bridge/bridge/protocol";
import { readWordLiveContext } from "./live-context";
import { resolveWordDocumentTitle } from "./document-title";
import { createWordBridgeTools } from "./tools";

declare const Office: any;
declare const Word: any;

const FALLBACK_DOCUMENT_ID_KEY = "word-mcp-bridge:fallback-document-id";
const BRIDGE_ENABLE_STORAGE_KEY = "office-agents-bridge-enabled";
const BRIDGE_URL_STORAGE_KEY = "office-agents-bridge-url";

const ENABLED_CAPABILITIES: BridgeCapability[] = [
  "observe",
  "document_edit",
  "unsafe_office_js",
];

interface CreateWordBridgeAdapterOptions {
  getRuntimeState?: () => BridgeRuntimeStateSlice | null;
}

export function createWordBridgeAdapter(
  options: CreateWordBridgeAdapterOptions = {},
) {
  const tools = createWordBridgeTools();
  return {
    metadataTag: "word_context",
    tools,
    async getDocumentId() {
      return getDocumentId();
    },
    async getDocumentMetadata() {
      return {
        metadata: await readWordDocumentMetadata(),
      };
    },
    async getLiveContext() {
      return readWordLiveContext();
    },
    getCapabilities() {
      return [...ENABLED_CAPABILITIES];
    },
    getRuntimeState() {
      return options.getRuntimeState?.() ?? null;
    },
  };
}

export function resolveConfiguredBridgeUrl(): string {
  const params = new URLSearchParams(window.location.search);
  const queryValue = params.get("office_bridge_url");
  if (queryValue?.trim()) {
    return queryValue.trim();
  }

  try {
    const stored = localStorage.getItem(BRIDGE_URL_STORAGE_KEY);
    if (stored?.trim()) {
      return stored.trim();
    }
  } catch {
    // Ignore storage access failures.
  }

  return "wss://localhost:4017/ws";
}

export function resolveBridgeSessionsUrl(bridgeWebSocketUrl: string): string {
  const parsed = new URL(bridgeWebSocketUrl);
  parsed.protocol = parsed.protocol === "wss:" ? "https:" : "http:";
  parsed.pathname = "/sessions";
  parsed.search = "";
  parsed.hash = "";
  return parsed.toString();
}

export function isBridgeForcedEnabled(): boolean {
  const params = new URLSearchParams(window.location.search);
  const queryValue = params.get("office_bridge");
  if (queryValue === "0" || queryValue === "false") {
    return false;
  }
  if (queryValue === "1" || queryValue === "true") {
    return true;
  }

  try {
    return localStorage.getItem(BRIDGE_ENABLE_STORAGE_KEY) !== "false";
  } catch {
    return true;
  }
}

async function getDocumentId(): Promise<string> {
  const documentUrl = getOfficeDocumentUrl();
  if (documentUrl) {
    return documentUrl;
  }

  try {
    const existing = sessionStorage.getItem(FALLBACK_DOCUMENT_ID_KEY);
    if (existing) return existing;
    const generated = `word-local:${crypto.randomUUID()}`;
    sessionStorage.setItem(FALLBACK_DOCUMENT_ID_KEY, generated);
    return generated;
  } catch {
    return "word-local:unknown-document";
  }
}

async function readWordDocumentMetadata() {
  try {
    return await Word.run(async (context: any) => {
      const body = context.document.body;
      const selection = context.document.getSelection();
      const paragraphs = body.paragraphs;

      body.load("text");
      selection.load("text");
      paragraphs.load("items");

      try {
        context.document.load("changeTrackingMode");
      } catch {
        // Ignore unsupported tracking APIs.
      }

      await context.sync();

      const bodyText = typeof body.text === "string" ? body.text : "";
      const selectedText =
        typeof selection.text === "string" ? selection.text.trim() : "";

      return {
        title: getDocumentTitle(),
        url: getOfficeDocumentUrl(),
        trackingMode: safeReadTrackingMode(context.document),
        wordCount: countWords(bodyText),
        characterCount: bodyText.length,
        paragraphCount: Array.isArray(paragraphs.items) ? paragraphs.items.length : 0,
        selectionLength: selectedText.length,
        updatedAt: Date.now(),
      };
    });
  } catch (error) {
    return {
      title: getDocumentTitle(),
      url: getOfficeDocumentUrl(),
      trackingMode: "Unknown",
      warning:
        error instanceof Error ? error.message : "Could not read Word metadata",
      updatedAt: Date.now(),
    };
  }
}

function safeReadTrackingMode(documentLike: any): string {
  try {
    return String(documentLike.changeTrackingMode ?? "Unknown");
  } catch {
    return "Unknown";
  }
}

function getOfficeDocumentUrl(): string | null {
  const url = Office?.context?.document?.url;
  return typeof url === "string" && url.trim() ? url.trim() : null;
}

function getDocumentTitle(): string | null {
  return resolveWordDocumentTitle({
    documentUrl: getOfficeDocumentUrl(),
    pageTitle: typeof document !== "undefined" ? document.title : null,
  });
}

export function deriveWordTaskpaneRuntimeState(
  status: OfficeBridgeConnectionStatus,
): BridgeRuntimeStateSlice {
  const connected = status.phase === "connected" && status.isConnected;
  const reconnecting = status.phase === "reconnecting";
  const mode = connected
    ? "ready"
    : reconnecting
      ? "reconnecting"
      : "connecting";

  return {
    mode,
    taskPhase: status.phase,
    isStreaming: false,
    permissionMode: "local",
    waitingState: connected ? null : "bridge_connection",
    waitingReason: connected
      ? null
      : (status.lastError?.message ??
        "Waiting for the Word taskpane to finish attaching to the local bridge."),
    handoffSummary: null,
    nextRecommendedAction: connected
      ? "Use Refresh if the document context looks stale."
      : "Keep the helper open and let the Word taskpane finish connecting.",
    activePlanSummary: null,
    activeTaskSummary: null,
    contextBudget: { usagePct: 0, action: "healthy" },
    lastVerification: null,
    latestCompletion: null,
    sessionStats: {
      inputTokens: 0,
      outputTokens: 0,
      totalCost: 0,
      messageCount: 0,
    },
    error: status.lastError?.message ?? null,
    threadCount: 0,
    activeThreadId: null,
    degradedGuardrails: connected ? [] : ["bridge_connection"],
    promptProvenance: null,
  };
}

function countWords(text: string): number {
  const trimmed = text.trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}
