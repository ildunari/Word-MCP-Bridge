import type { BridgeSessionSnapshot } from "./protocol.js";

const GENERIC_WORD_TITLES = new Set([
  "microsoft word",
  "word",
  "word mcp bridge",
]);

function normalizeText(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function looksLikeFilename(value: string): boolean {
  return /[./\\]/.test(value);
}

function basename(value: string): string {
  return value.split(/[\\/]/).pop() ?? value;
}

function documentTitle(snapshot: BridgeSessionSnapshot): string | null {
  const metadata =
    snapshot.documentMetadata &&
    typeof snapshot.documentMetadata === "object" &&
    !Array.isArray(snapshot.documentMetadata)
      ? (snapshot.documentMetadata as Record<string, unknown>)
      : null;

  const title = normalizeText(metadata?.title);
  if (!title) return null;

  if (snapshot.app === "word" && GENERIC_WORD_TITLES.has(title.toLowerCase())) {
    return null;
  }

  return title;
}

export function isUnsavedWordDocument(
  snapshot: Pick<BridgeSessionSnapshot, "app" | "documentId">,
): boolean {
  return snapshot.app === "word" && snapshot.documentId.startsWith("word-local:");
}

export function getSessionDocumentLabel(
  snapshot: BridgeSessionSnapshot,
): string {
  const title = documentTitle(snapshot);
  if (title) return title;

  if (isUnsavedWordDocument(snapshot)) {
    return "Untitled Word document";
  }

  const documentId = normalizeText(snapshot.documentId);
  if (documentId && looksLikeFilename(documentId)) {
    return basename(documentId);
  }

  const appLabel = normalizeText(snapshot.appName) ?? snapshot.app;
  return `${appLabel} document`;
}

export function getSessionDocumentSummary(
  snapshot: BridgeSessionSnapshot,
): string {
  if (isUnsavedWordDocument(snapshot)) {
    return "Unsaved local document";
  }

  const documentId = normalizeText(snapshot.documentId);
  if (documentId && looksLikeFilename(documentId)) {
    return "Saved document";
  }

  if (documentTitle(snapshot)) {
    return documentId ? `Document ID ${documentId}` : "Saved document";
  }

  return documentId ? `Document ID ${documentId}` : "Live document";
}

export function describeSessionChoice(
  snapshot: BridgeSessionSnapshot,
): string {
  const label = getSessionDocumentLabel(snapshot);
  const summary = getSessionDocumentSummary(snapshot);
  return `${label} [${snapshot.sessionId}; ${summary}]`;
}
