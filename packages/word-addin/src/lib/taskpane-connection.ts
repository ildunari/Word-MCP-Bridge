import type {
  OfficeBridgeConnectionStatus,
  OfficeBridgeSessionHealth,
} from "@word-mcp-bridge/bridge/client";
import type {
  BridgeLiveContext,
  BridgeSessionSnapshot,
} from "@word-mcp-bridge/bridge/protocol";

export interface TaskpaneConnectionView {
  headline: string;
  subtitle: string;
  statusLabel: string;
}

export interface TaskpaneDashboardCard {
  eyebrow: string;
  title: string;
  summary: string;
  detail?: string;
}

export interface TaskpaneDashboardView {
  statusLabel:
    | "Connected"
    | "Connecting"
    | "Needs helper app"
    | "Multiple documents connected"
    | "Reconnect needed";
  headline: string;
  subtitle: string;
  tone: "ready" | "working" | "warning";
  documentCard: TaskpaneDashboardCard;
  selectionCard: TaskpaneDashboardCard;
  reviewCard: TaskpaneDashboardCard;
  readinessCard: TaskpaneDashboardCard & {
    items: string[];
    warning?: string;
  };
  connectionCard: TaskpaneDashboardCard & {
    items: string[];
  };
  diagnosticsCard: TaskpaneDashboardCard & {
    items: string[];
  };
  recoveryTitle?: string;
  recoveryText?: string;
  showReconnectAction: boolean;
  showHelperHint: boolean;
}

interface TaskpaneConnectionViewOptions {
  snapshot: BridgeSessionSnapshot | null;
  bridgeStatus: OfficeBridgeConnectionStatus;
  serverSessionRegistered?: boolean;
  connectedSessionCount?: number;
  matchedBy?: "instanceId" | "documentId" | null;
  matchedServerSessionId?: string | null;
  serverPendingCount?: number;
  serverSessionHealth?: string | null;
  lastRegistrationSyncAt?: number | null;
  lastRegistrationSyncState?: "matched" | "pending" | "error";
}

export function deriveTaskpaneDashboardView(
  options: TaskpaneConnectionViewOptions,
): TaskpaneDashboardView {
  const {
    snapshot,
    bridgeStatus,
    serverSessionRegistered = false,
    connectedSessionCount = 0,
    matchedBy = null,
    matchedServerSessionId = null,
    serverPendingCount = 0,
    serverSessionHealth = null,
    lastRegistrationSyncAt = null,
    lastRegistrationSyncState = "pending",
  } = options;
  const liveContext = snapshot?.gateway?.liveContext ?? null;
  const documentMetadata = normalizeMetadata(snapshot);
  const documentTitle = resolveDocumentTitle(snapshot, documentMetadata);
  const reviewMode = resolveReviewMode(liveContext, documentMetadata);
  const selectionText = resolveSelectionText(liveContext);
  const hasSelection = Boolean(liveContext?.selection?.hasSelection);
  const transportConnected =
    Boolean(snapshot) &&
    bridgeStatus.phase === "connected" &&
    bridgeStatus.isConnected;
  const registrationAvailable = lastRegistrationSyncState !== "error";
  const sessionReady =
    transportConnected &&
    registrationAvailable &&
    (connectedSessionCount === 0 || serverSessionRegistered);
  const multipleDocumentsConnected = sessionReady && connectedSessionCount > 1;
  const canUseSelection = sessionReady && hasSelection;

  let statusLabel: TaskpaneDashboardView["statusLabel"] = "Connecting";
  let headline = sessionReady
    ? `${documentTitle} is ready`
    : "Getting this document ready";
  let subtitle =
    "The panel is collecting the current document and selection so the agent can use them.";
  let tone: TaskpaneDashboardView["tone"] = sessionReady ? "ready" : "working";
  let recoveryTitle: string | undefined;
  let recoveryText: string | undefined;
  let showReconnectAction = false;
  let showHelperHint = false;

  if (!bridgeStatus.enabled) {
    statusLabel = "Needs helper app";
    headline = "The local Word tools are turned off";
    subtitle =
      "Start the helper app, then reopen or refresh this panel to attach the current document.";
    tone = "warning";
    recoveryTitle = "Start the helper app";
    recoveryText =
      "Open Word MCP Bridge Helper.app, wait for it to show green status, then come back here and reconnect.";
    showHelperHint = true;
  } else if (multipleDocumentsConnected) {
    statusLabel = "Multiple documents connected";
    headline = `${documentTitle} is ready`;
    subtitle =
      "More than one Word document is live. This panel reflects the document you are viewing right now.";
    tone = "ready";
    showReconnectAction = false;
  } else if (sessionReady) {
    statusLabel = "Connected";
    headline = `${documentTitle} is ready`;
    subtitle =
      "Your current document and selection are ready for the agent to use.";
    tone = "ready";
    showReconnectAction = false;
  } else if (transportConnected && connectedSessionCount > 0 && !serverSessionRegistered) {
    statusLabel = "Connecting";
    headline = `${documentTitle} is opening`;
    subtitle =
      "This panel is still matching itself to the live Word document list. Give it a moment, then refresh if needed.";
    tone = "working";
    recoveryTitle = "Still connecting?";
    recoveryText =
      "If another Word document is open, use the helper app to confirm which document is connected to this panel.";
    showHelperHint = true;
  } else if (bridgeStatus.phase === "connecting") {
    statusLabel = "Connecting";
    headline = snapshot ? `${documentTitle} is opening` : "Connecting this document";
    subtitle =
      "The panel is still attaching to the local Word system. This usually settles in a moment.";
    tone = "working";
    recoveryTitle = "Still connecting?";
    recoveryText =
      "If this takes longer than a few seconds, refresh the panel. If it still hangs, open the helper app and check that Word is connected.";
    showHelperHint = true;
  } else if (
    bridgeStatus.phase === "disconnected" ||
    bridgeStatus.phase === "reconnecting" ||
    bridgeStatus.lastError
  ) {
    statusLabel = "Reconnect needed";
    headline = snapshot ? `${documentTitle} is available locally` : "Reconnect needed";
    subtitle =
      "The panel lost contact with the local Word bridge. Reconnect to keep this document live for the agent." +
      formatLastError(bridgeStatus);
    tone = "warning";
    recoveryTitle = "Reconnect this document";
    recoveryText = buildReconnectRecoveryText(bridgeStatus);
    showReconnectAction = true;
    showHelperHint = true;
  } else {
    statusLabel = "Connecting";
    headline = snapshot ? `${documentTitle} is opening` : "Connecting this document";
    subtitle =
      "The panel is still attaching to the local Word system. This usually settles in a moment.";
    tone = "working";
    recoveryTitle = "Still connecting?";
    recoveryText =
      "If this takes longer than a few seconds, refresh the panel. If it still hangs, open the helper app and check that Word is connected.";
    showHelperHint = true;
  }

  const documentCard: TaskpaneDashboardCard = {
    eyebrow: "This document",
    title: documentTitle,
    summary: documentMetadata.url
      ? "Saved Word document"
      : "Unsaved document on this Mac",
    detail: buildDocumentDetail(documentMetadata),
  };

  const selectionCard: TaskpaneDashboardCard = {
    eyebrow: "Current selection",
    title: hasSelection ? "Selection ready" : "No selection",
    summary: hasSelection
      ? "The agent can use the text you have selected right now."
      : "Select text in Word if you want the agent to focus on a specific passage.",
    detail: selectionText,
  };

  const reviewCard: TaskpaneDashboardCard = {
    eyebrow: "Review context",
    title: reviewMode,
    summary: describeReviewMode(reviewMode),
    detail: buildReviewDetail(documentMetadata, liveContext),
  };

  const readinessItems = [
    sessionReady
      ? "Current document text and structure are available."
      : "Current document text will appear here once the panel finishes connecting.",
    canUseSelection
      ? "Your current selection is ready for the agent."
      : "No live selection is available yet.",
    reviewMode === "Unknown"
      ? "Review mode is not available yet."
      : `Review mode is ${reviewMode.toLowerCase()}.`,
  ];

  const readinessCard: TaskpaneDashboardView["readinessCard"] = {
    eyebrow: "AI-ready context",
    title: sessionReady
      ? "Ready for this document"
      : "Waiting for the document to finish connecting",
    summary: sessionReady
      ? "This panel is focused on the details the agent can use right now."
      : "The agent will get richer context here as soon as the connection settles.",
    items: readinessItems,
    ...(multipleDocumentsConnected
      ? {
          warning:
            "More than one Word document is connected. The helper app is the best place to inspect the full document list.",
        }
      : {}),
  };

  const connectionCard: TaskpaneDashboardView["connectionCard"] = {
    eyebrow: "Connection details",
    title: humanizeSessionHealth(
      sessionReady ? bridgeStatus.sessionHealth : "registration_pending",
    ),
    summary:
      "This panel tracks both the taskpane instance and the currently matched Word document.",
    detail:
      lastRegistrationSyncAt != null
        ? `Last registration sync ${new Date(lastRegistrationSyncAt).toLocaleTimeString()}.`
        : "Registration details will appear after the first successful sync.",
    items: [
      `Panel instance: ${snapshot?.instanceId ?? "Not available yet"}`,
      `Matched session: ${matchedServerSessionId ?? snapshot?.sessionId ?? "Pending"}`,
      `Document ID: ${snapshot?.documentId ?? "Pending"}`,
      `Match source: ${matchedBy === "instanceId" ? "Instance ID" : matchedBy === "documentId" ? "Document ID fallback" : "Pending"}`,
      `Connected Word docs: ${connectedSessionCount}`,
      `Server session health: ${humanizeSessionHealth(
        serverSessionHealth ?? (serverPendingCount > 0 ? "registration_pending" : serverSessionRegistered ? "live" : "registration_pending"),
      )}`,
      `Server pending requests: ${serverPendingCount}`,
      `Registration state: ${humanizeRegistrationState(lastRegistrationSyncState)}`,
    ],
  };

  const diagnosticsCard: TaskpaneDashboardView["diagnosticsCard"] = {
    eyebrow: "Diagnostics",
    title: "Recent bridge signals",
    summary:
      "These are the latest connection breadcrumbs the taskpane can use to explain refresh and reconnect behavior.",
    items: buildDiagnosticItems(bridgeStatus, lastRegistrationSyncAt),
  };

  return {
    statusLabel,
    headline,
    subtitle,
    tone,
    documentCard,
    selectionCard,
    reviewCard,
    readinessCard,
    connectionCard,
    diagnosticsCard,
    recoveryTitle,
    recoveryText,
    showReconnectAction,
    showHelperHint,
  };
}

export function deriveTaskpaneConnectionView(
  options: TaskpaneConnectionViewOptions,
): TaskpaneConnectionView {
  const dashboard = deriveTaskpaneDashboardView(options);
  return {
    headline: dashboard.headline,
    subtitle: dashboard.subtitle,
    statusLabel: dashboard.statusLabel,
  };
}

function normalizeMetadata(
  snapshot: BridgeSessionSnapshot | null,
): Record<string, unknown> {
  const metadata = snapshot?.documentMetadata;
  if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) {
    return {};
  }
  return metadata as Record<string, unknown>;
}

function resolveDocumentTitle(
  snapshot: BridgeSessionSnapshot | null,
  metadata: Record<string, unknown>,
): string {
  const rawTitle = stringOrNull(metadata.title);
  if (rawTitle && rawTitle !== "Word MCP Bridge") return rawTitle;

  const documentId = snapshot?.documentId ?? "";
  if (documentId.startsWith("word-local:")) {
    return "Untitled Word document";
  }

  return documentId || "Word document";
}

function resolveReviewMode(
  liveContext: BridgeLiveContext | null | undefined,
  metadata: Record<string, unknown>,
): string {
  return (
    liveContext?.trackingMode ??
    stringOrNull(metadata.trackingMode) ??
    "Unknown"
  );
}

function resolveSelectionText(
  liveContext: BridgeLiveContext | null | undefined,
): string {
  const selectedText = liveContext?.selection?.selectedText?.trim();
  if (selectedText) return selectedText;
  return "Nothing is selected right now.";
}

function buildDocumentDetail(metadata: Record<string, unknown>): string {
  const wordCount = numberOrNull(metadata.wordCount);
  const paragraphCount = numberOrNull(metadata.paragraphCount);
  const parts: string[] = [];

  if (wordCount != null) {
    parts.push(`${wordCount} ${wordCount === 1 ? "word" : "words"}`);
  }
  if (paragraphCount != null) {
    parts.push(`${paragraphCount} ${paragraphCount === 1 ? "paragraph" : "paragraphs"}`);
  }

  return parts.length > 0 ? parts.join(" across ") : "Document details will appear when Word finishes reporting them.";
}

function describeReviewMode(reviewMode: string): string {
  if (reviewMode === "Off") {
    return "Track changes is currently off for this document.";
  }
  if (reviewMode === "TrackAll") {
    return "Track changes is on for everyone editing this document.";
  }
  if (reviewMode === "TrackMineOnly") {
    return "Track changes is only tracking your edits right now.";
  }
  return "Review details are still loading from Word.";
}

function buildReviewDetail(
  metadata: Record<string, unknown>,
  liveContext: BridgeLiveContext | null | undefined,
): string {
  const selectionLength = numberOrNull(metadata.selectionLength);
  const focusTarget = liveContext?.focusTarget;
  const details: string[] = [];

  if (focusTarget && focusTarget !== "unknown") {
    details.push(`Focus is on the ${focusTarget}.`);
  }
  if (selectionLength != null && selectionLength > 0) {
    details.push(`${selectionLength} characters are currently selected.`);
  }

  return details.join(" ") || "The panel is following the document and your selection as you move through Word.";
}

function formatLastError(status: OfficeBridgeConnectionStatus): string {
  const message = status.lastError?.message?.trim();
  return message ? ` Last problem: ${message}` : "";
}

function buildReconnectRecoveryText(status: OfficeBridgeConnectionStatus): string {
  const details: string[] = [
    "Refresh this panel first. If that does not help, open the helper app and confirm the local bridge is running.",
  ];
  const messages = status.diagnostics.map((entry) => entry.message.toLowerCase());

  if (messages.some((message) => message.includes("helper poll"))) {
    details.push("The helper poll is still reaching Word, so this looks more like a taskpane-side reconnect than a full helper outage.");
  }
  if (messages.some((message) => message.includes("focus refresh"))) {
    details.push("A taskpane focus refresh ran just before the reconnect warning.");
  }
  if (messages.some((message) => message.includes("stale bundle") || message.includes("stale page reload"))) {
    details.push("A stale taskpane bundle was reloaded.");
  }
  if (messages.some((message) => message.includes("websocket closed"))) {
    details.push("The websocket dropped, and the taskpane is still trying to reattach.");
  }

  return details.join(" ");
}

function buildDiagnosticItems(
  status: OfficeBridgeConnectionStatus,
  lastRegistrationSyncAt: number | null,
): string[] {
  const details = status.details;
  return [
    `Transport health: ${humanizeSessionHealth(status.sessionHealth)}`,
    `Last websocket open: ${formatTimestamp(details.lastWebSocketOpenAt)}`,
    `Last hello snapshot: ${formatTimestamp(details.lastHelloSnapshotCapturedAt)}`,
    `Last hello sent: ${formatTimestamp(details.lastHelloSentAt)}`,
    `Last session update: ${formatTimestamp(details.lastSessionUpdatedAt)}`,
    `Last completed tool: ${formatTimestamp(details.lastToolCompletedAt)}`,
    `Last timed-out request: ${formatTimedOutInvoke(details.lastTimedOutInvoke)}`,
    `Active invokes: ${details.activeInvokeCount}`,
    `Last registration sync: ${formatTimestamp(lastRegistrationSyncAt)}`,
  ];
}

function humanizeSessionHealth(health: OfficeBridgeSessionHealth | string): string {
  switch (health) {
    case "live":
      return "Live";
    case "registration_pending":
      return "Registration pending";
    case "stale":
      return "Stale";
    case "reconnecting":
      return "Reconnecting";
    case "orphaned":
      return "Orphaned";
    default:
      return health;
  }
}

function humanizeRegistrationState(state: "matched" | "pending" | "error"): string {
  if (state === "matched") return "Matched";
  if (state === "error") return "Error";
  return "Pending";
}

function formatTimestamp(value: number | null): string {
  return value == null ? "Not yet" : new Date(value).toLocaleTimeString();
}

function formatTimedOutInvoke(
  value: OfficeBridgeConnectionStatus["details"]["lastTimedOutInvoke"],
): string {
  if (!value) return "None";
  return `${value.toolName ?? value.method} at ${new Date(value.at).toLocaleTimeString()}`;
}

function stringOrNull(value: unknown): string | null {
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

function numberOrNull(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}
