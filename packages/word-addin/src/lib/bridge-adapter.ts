import type { BridgeCapability } from "@word-mcp-bridge/bridge/protocol";
import { readWordLiveContext } from "./live-context";

declare const Office: any;
declare const Word: any;

const FALLBACK_DOCUMENT_ID_KEY = "word-mcp-bridge:fallback-document-id";
const BRIDGE_ENABLE_STORAGE_KEY = "office-agents-bridge-enabled";
const BRIDGE_URL_STORAGE_KEY = "office-agents-bridge-url";

const ENABLED_CAPABILITIES: BridgeCapability[] = ["observe", "unsafe_office_js"];

export function createWordBridgeAdapter() {
  return {
    metadataTag: "word_context",
    tools: [],
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
  const documentUrl = getOfficeDocumentUrl();
  if (documentUrl) {
    const basename = documentUrl.split("/").pop()?.trim();
    if (basename) return basename;
  }

  if (typeof document !== "undefined" && document.title.trim()) {
    return document.title.trim();
  }

  return null;
}

function countWords(text: string): number {
  const trimmed = text.trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}
