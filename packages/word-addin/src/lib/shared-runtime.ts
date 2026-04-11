import {
  type OfficeBridgeConnectionStatus,
  type OfficeBridgeController,
  startOfficeBridge,
} from "@word-mcp-bridge/bridge/client";
import type {
  BridgeLiveContext,
  BridgeSessionSnapshot,
} from "@word-mcp-bridge/bridge/protocol";
import {
  createWordBridgeAdapter,
  deriveWordTaskpaneRuntimeState,
  isBridgeForcedEnabled,
  resolveBridgeSessionsUrl,
  resolveConfiguredBridgeUrl,
} from "./bridge-adapter";
import { bindOfficeDocumentHandler } from "./components/office-document-events";
import {
  attachWordLiveContextBridge,
  WORD_TRACKING_MODE_CHANGED_EVENT,
} from "./live-context";

declare const Office: any;

type BridgeSessionLookupResponse = {
  ok?: boolean;
  sessions?: {
    pendingCount?: number;
    health?: string;
    snapshot?: (BridgeSessionSnapshot & { app?: string }) | null;
  }[];
};

export type WordTaskpaneVisibilityMode = "visible" | "hidden" | "unknown";
export type WordTaskpaneStartupBehavior =
  | "inactive"
  | "load"
  | "unsupported"
  | "error";

export interface WordTaskpaneRuntimeSnapshot {
  bridgeEnabled: boolean;
  bridgeUrl: string;
  bridgeStatus: OfficeBridgeConnectionStatus;
  snapshot: BridgeSessionSnapshot | null;
  liveContext: BridgeLiveContext | null;
  isRefreshing: boolean;
  lastRefreshLabel: string;
  errorMessage: string;
  serverSessionRegistered: boolean;
  connectedSessionCount: number;
  matchedBy: "instanceId" | "documentId" | null;
  matchedServerSessionId: string | null;
  serverPendingCount: number;
  serverSessionHealth: string | null;
  lastRegistrationSyncAt: number | null;
  lastRegistrationSyncState: "matched" | "pending" | "error";
  paneVisibility: WordTaskpaneVisibilityMode;
  sharedRuntimeAvailable: boolean;
  startupBehavior: WordTaskpaneStartupBehavior;
  instanceId: string | null;
}

type Listener = (snapshot: WordTaskpaneRuntimeSnapshot) => void;

const initialBridgeEnabled = isBridgeForcedEnabled();
const initialBridgeUrl = resolveConfiguredBridgeUrl();

const initialBridgeStatus: OfficeBridgeConnectionStatus = {
  enabled: initialBridgeEnabled,
  serverUrl: initialBridgeUrl,
  phase: initialBridgeEnabled ? "connecting" : "disabled",
  isConnected: false,
  hasConnected: false,
  lastError: null,
  sessionHealth: initialBridgeEnabled ? "registration_pending" : "orphaned",
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
};

const runtimeState: WordTaskpaneRuntimeSnapshot = {
  bridgeEnabled: initialBridgeEnabled,
  bridgeUrl: initialBridgeUrl,
  bridgeStatus: initialBridgeStatus,
  snapshot: null,
  liveContext: null,
  isRefreshing: false,
  lastRefreshLabel: "Never",
  errorMessage: "",
  serverSessionRegistered: false,
  connectedSessionCount: 0,
  matchedBy: null,
  matchedServerSessionId: null,
  serverPendingCount: 0,
  serverSessionHealth: null,
  lastRegistrationSyncAt: null,
  lastRegistrationSyncState: "pending",
  paneVisibility: "unknown",
  sharedRuntimeAvailable: false,
  startupBehavior: "inactive",
  instanceId: null,
};

const listeners = new Set<Listener>();

let controller: OfficeBridgeController | null = null;
let activeRefresh: Promise<void> | null = null;
let deferredRefreshTimer: number | null = null;
let registrationSyncGeneration = 0;
let initialized = false;
let visibilityModeCleanup: (() => Promise<void> | void) | null = null;
let detachBridgeEvents = () => undefined;
let detachSelectionHandler = () => undefined;
let unsubscribeBridgeStatus = () => undefined;
let statusPollTimer: number | null = null;
let runtimeReadyPromise: Promise<void> | null = null;

function cloneSnapshot(): WordTaskpaneRuntimeSnapshot {
  return {
    ...runtimeState,
    bridgeStatus: {
      ...runtimeState.bridgeStatus,
      details: { ...runtimeState.bridgeStatus.details },
      diagnostics: [...runtimeState.bridgeStatus.diagnostics],
    },
  };
}

function publish() {
  const next = cloneSnapshot();
  for (const listener of listeners) {
    listener(next);
  }
}

function updateRuntimeState(partial: Partial<WordTaskpaneRuntimeSnapshot>) {
  Object.assign(runtimeState, partial);
  publish();
}

function clearMatchedServerState() {
  runtimeState.serverSessionRegistered = false;
  runtimeState.connectedSessionCount = 0;
  runtimeState.matchedBy = null;
  runtimeState.matchedServerSessionId = null;
  runtimeState.serverPendingCount = 0;
  runtimeState.serverSessionHealth = null;
}

function shouldAdoptServerSnapshot(
  currentSnapshot: BridgeSessionSnapshot | null,
  nextSnapshot: BridgeSessionSnapshot,
) {
  if (!currentSnapshot) return true;
  if (currentSnapshot.sessionId !== nextSnapshot.sessionId) return true;
  return (nextSnapshot.updatedAt ?? 0) >= (currentSnapshot.updatedAt ?? 0);
}

function syncBridgeStatus() {
  if (!controller) return;
  runtimeState.bridgeStatus = controller.getStatus();
  publish();
}

function resolveMatchedSession(
  sessions: NonNullable<BridgeSessionLookupResponse["sessions"]>,
) {
  const wordSessions = sessions.filter((session) => session.snapshot?.app === "word");
  const byInstanceId = wordSessions.find(
    (session) => session.snapshot?.instanceId === controller?.instanceId,
  );
  if (byInstanceId) {
    return {
      matchedSession: byInstanceId,
      connectedWordSessions: wordSessions.length,
      matchedBy: "instanceId" as const,
    };
  }

  const documentId =
    runtimeState.snapshot?.documentId != null &&
    runtimeState.snapshot.documentId.trim().length > 0
      ? runtimeState.snapshot.documentId
      : null;
  const byDocumentIdMatches = documentId
    ? wordSessions.filter((session) => session.snapshot?.documentId === documentId)
    : [];
  const byDocumentId =
    byDocumentIdMatches.length === 1 ? byDocumentIdMatches[0] : undefined;

  return {
    matchedSession: byDocumentId,
    connectedWordSessions: wordSessions.length,
    matchedBy: byDocumentId ? ("documentId" as const) : null,
  };
}

async function syncSessionRegistration() {
  const generation = ++registrationSyncGeneration;
  try {
    const response = await fetch(
      resolveBridgeSessionsUrl(
        runtimeState.bridgeStatus.serverUrl || runtimeState.bridgeUrl,
      ),
    );
    if (generation !== registrationSyncGeneration) return;
    if (!response.ok) {
      clearMatchedServerState();
      runtimeState.lastRegistrationSyncAt = Date.now();
      runtimeState.lastRegistrationSyncState = "error";
      publish();
      return;
    }
    const payload = (await response.json()) as BridgeSessionLookupResponse;
    if (generation !== registrationSyncGeneration) return;
    const {
      matchedSession,
      connectedWordSessions,
      matchedBy: nextMatchedBy,
    } = resolveMatchedSession(payload.sessions ?? []);
    runtimeState.connectedSessionCount = connectedWordSessions;
    runtimeState.serverSessionRegistered = Boolean(matchedSession?.snapshot);
    runtimeState.matchedBy = nextMatchedBy;
    runtimeState.matchedServerSessionId = matchedSession?.snapshot?.sessionId ?? null;
    runtimeState.serverPendingCount = matchedSession?.pendingCount ?? 0;
    runtimeState.serverSessionHealth = matchedSession?.health ?? null;
    runtimeState.lastRegistrationSyncAt = Date.now();
    runtimeState.lastRegistrationSyncState = matchedSession?.snapshot
      ? "matched"
      : "pending";
    if (
      matchedSession?.snapshot &&
      shouldAdoptServerSnapshot(runtimeState.snapshot, matchedSession.snapshot)
    ) {
      runtimeState.snapshot = matchedSession.snapshot;
      runtimeState.liveContext = matchedSession.snapshot.gateway?.liveContext ?? null;
    }
    publish();
  } catch {
    if (generation !== registrationSyncGeneration) return;
    clearMatchedServerState();
    runtimeState.lastRegistrationSyncAt = Date.now();
    runtimeState.lastRegistrationSyncState = "error";
    publish();
  }
}

function shouldShowRefreshBusyState(reason: string) {
  return (
    reason === "manual refresh" ||
    reason === "reconnect" ||
    reason === "hide panel" ||
    reason === "show panel"
  );
}

async function refreshStatus(reason = "manual refresh") {
  if (!controller) return;
  const showBusyState = shouldShowRefreshBusyState(reason);
  if (activeRefresh) {
    if (showBusyState) {
      await activeRefresh;
    }
    return;
  }

  if (showBusyState) {
    updateRuntimeState({ isRefreshing: true });
  }
  updateRuntimeState({ errorMessage: "" });
  syncBridgeStatus();

  activeRefresh = (async () => {
    try {
      const nextSnapshot = await controller!.refresh();
      runtimeState.snapshot = nextSnapshot;
      runtimeState.liveContext = nextSnapshot?.gateway?.liveContext ?? null;
      runtimeState.lastRefreshLabel = new Date().toLocaleTimeString();
      syncBridgeStatus();
      await syncSessionRegistration();
    } catch (error) {
      runtimeState.errorMessage =
        error instanceof Error
          ? error.message
          : "Could not refresh Word bridge status.";
      syncBridgeStatus();
      await syncSessionRegistration();
    }
    publish();
  })();

  try {
    await activeRefresh;
  } finally {
    activeRefresh = null;
    if (showBusyState) {
      updateRuntimeState({ isRefreshing: false });
    }
  }
}

function scheduleRefresh(reason: string, delayMs = 160) {
  if (deferredRefreshTimer !== null) {
    window.clearTimeout(deferredRefreshTimer);
    deferredRefreshTimer = null;
  }
  deferredRefreshTimer = window.setTimeout(() => {
    deferredRefreshTimer = null;
    void refreshStatus(reason);
  }, delayMs);
}

function normalizeVisibilityMode(
  value: unknown,
): WordTaskpaneVisibilityMode {
  if (typeof value !== "string") return "unknown";
  const normalized = value.toLowerCase();
  if (normalized === "taskpane") return "visible";
  if (
    normalized === "hidden" ||
    normalized === "none" ||
    normalized === "minimized"
  ) {
    return "hidden";
  }
  return "unknown";
}

async function registerVisibilityModeListener() {
  const addin = Office?.addin;
  runtimeState.sharedRuntimeAvailable = Boolean(
    addin?.hide && addin?.showAsTaskpane,
  );
  runtimeState.paneVisibility = runtimeState.sharedRuntimeAvailable
    ? "visible"
    : "unknown";
  publish();

  if (!addin?.onVisibilityModeChanged) {
    return;
  }

  try {
    const removeHandler = await addin.onVisibilityModeChanged((args: any) => {
      runtimeState.paneVisibility = normalizeVisibilityMode(args?.visibilityMode);
      publish();
      void refreshStatus("visibility mode change");
    });
    visibilityModeCleanup = removeHandler;
  } catch {
    visibilityModeCleanup = null;
  }
}

async function setStartupBehaviorLoad() {
  const addin = Office?.addin;
  if (!addin?.setStartupBehavior || !Office?.StartupBehavior?.load) {
    updateRuntimeState({ startupBehavior: "unsupported" });
    return false;
  }

  try {
    await addin.setStartupBehavior(Office.StartupBehavior.load);
    updateRuntimeState({ startupBehavior: "load" });
    return true;
  } catch {
    updateRuntimeState({ startupBehavior: "error" });
    return false;
  }
}

function buildRuntimeMetadata() {
  const connected =
    runtimeState.bridgeStatus.phase === "connected" &&
    runtimeState.bridgeStatus.isConnected;
  return {
    taskpaneVisibility: runtimeState.paneVisibility,
    sharedRuntimeEnabled: runtimeState.sharedRuntimeAvailable,
    startupBehavior: runtimeState.startupBehavior,
    hiddenActive: runtimeState.paneVisibility === "hidden" && connected,
  };
}

function initializeRuntime() {
  if (initialized) return;
  initialized = true;

  const adapter = createWordBridgeAdapter({
    getRuntimeState: () =>
      deriveWordTaskpaneRuntimeState(runtimeState.bridgeStatus, {
        paneVisibility: runtimeState.paneVisibility,
        sharedRuntimeAvailable: runtimeState.sharedRuntimeAvailable,
        startupBehavior: runtimeState.startupBehavior,
      }),
    getDocumentMetadataExtras: buildRuntimeMetadata,
  });
  controller = startOfficeBridge({
    app: "word",
    adapter,
    enabled: runtimeState.bridgeEnabled,
    serverUrl: runtimeState.bridgeUrl,
    forwardConsole: false,
  });
  runtimeState.instanceId = controller.instanceId;
  runtimeState.bridgeStatus = controller.getStatus();
  unsubscribeBridgeStatus = controller.subscribe((status) => {
    runtimeState.bridgeStatus = status;
    publish();
  });

  detachBridgeEvents = attachWordLiveContextBridge(controller);
  const officeDocument =
    typeof Office === "undefined" ? undefined : Office?.context?.document;
  detachSelectionHandler = bindOfficeDocumentHandler(
    officeDocument,
    typeof Office === "undefined"
      ? "DocumentSelectionChanged"
      : (Office?.EventType?.DocumentSelectionChanged ?? "DocumentSelectionChanged"),
    () => {
      scheduleRefresh("selection change");
    },
  );

  const handleWindowFocus = () => {
    scheduleRefresh("window focus");
  };
  const handleVisibilityChange = () => {
    if (document.visibilityState === "hidden") return;
    scheduleRefresh("visibility change");
  };
  const handleTrackingModeChange = () => {
    scheduleRefresh("tracking mode change");
  };

  window.addEventListener("focus", handleWindowFocus);
  document.addEventListener("visibilitychange", handleVisibilityChange);
  window.addEventListener(
    WORD_TRACKING_MODE_CHANGED_EVENT,
    handleTrackingModeChange,
  );

  statusPollTimer = window.setInterval(() => {
    syncBridgeStatus();
    void syncSessionRegistration();
  }, 1_500);

  void registerVisibilityModeListener();
  void refreshStatus("startup");
}

export async function ensureWordTaskpaneRuntime() {
  if (!runtimeReadyPromise) {
    runtimeReadyPromise = Promise.resolve().then(() => {
      initializeRuntime();
    });
  }
  await runtimeReadyPromise;
}

export function subscribeWordTaskpaneRuntime(listener: Listener) {
  listeners.add(listener);
  listener(cloneSnapshot());
  return () => {
    listeners.delete(listener);
  };
}

export function getWordTaskpaneRuntimeSnapshot() {
  return cloneSnapshot();
}

export async function refreshWordTaskpaneRuntime(reason = "manual refresh") {
  await ensureWordTaskpaneRuntime();
  await refreshStatus(reason);
}

async function waitForReconnectReady(timeoutMs = 5_000) {
  if (!controller) return false;
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    syncBridgeStatus();
    const status = controller.getStatus();
    if (status.phase === "connected" && status.isConnected) {
      return true;
    }
    await new Promise((resolve) => window.setTimeout(resolve, 120));
  }
  syncBridgeStatus();
  return false;
}

export async function reconnectWordTaskpaneRuntime() {
  await ensureWordTaskpaneRuntime();
  if (!controller) return;
  await controller.reconnect();
  await waitForReconnectReady();
  await refreshStatus("reconnect");
}

export async function hideWordTaskpaneRuntime() {
  await ensureWordTaskpaneRuntime();
  const addin = Office?.addin;
  if (!runtimeState.sharedRuntimeAvailable || !addin?.hide) {
    updateRuntimeState({
      errorMessage:
        "This Word client does not expose the shared-runtime hide API.",
    });
    return false;
  }
  const startupOkay = await setStartupBehaviorLoad();
  if (!startupOkay && runtimeState.startupBehavior === "error") {
    updateRuntimeState({
      errorMessage:
        "The add-in could not enable startup behavior before hiding the panel.",
    });
    return false;
  }
  try {
    await addin.hide();
    updateRuntimeState({ paneVisibility: "hidden" });
    await refreshStatus("hide panel");
    return true;
  } catch (error) {
    updateRuntimeState({
      errorMessage:
        error instanceof Error
          ? error.message
          : "The add-in could not hide the taskpane.",
    });
    return false;
  }
}

export async function showWordTaskpaneRuntime() {
  await ensureWordTaskpaneRuntime();
  const addin = Office?.addin;
  if (!runtimeState.sharedRuntimeAvailable || !addin?.showAsTaskpane) {
    updateRuntimeState({
      errorMessage:
        "This Word client does not expose the shared-runtime show API.",
    });
    return false;
  }
  try {
    await addin.showAsTaskpane();
    updateRuntimeState({ paneVisibility: "visible" });
    await refreshStatus("show panel");
    return true;
  } catch (error) {
    updateRuntimeState({
      errorMessage:
        error instanceof Error
          ? error.message
          : "The add-in could not show the taskpane.",
    });
    return false;
  }
}

export function __resetWordTaskpaneRuntimeForTests() {
  listeners.clear();
  controller?.stop();
  controller = null;
  unsubscribeBridgeStatus();
  unsubscribeBridgeStatus = () => undefined;
  detachBridgeEvents();
  detachBridgeEvents = () => undefined;
  detachSelectionHandler();
  detachSelectionHandler = () => undefined;
  if (statusPollTimer !== null) {
    window.clearInterval(statusPollTimer);
    statusPollTimer = null;
  }
  if (deferredRefreshTimer !== null) {
    window.clearTimeout(deferredRefreshTimer);
    deferredRefreshTimer = null;
  }
  void visibilityModeCleanup?.();
  visibilityModeCleanup = null;
  activeRefresh = null;
  registrationSyncGeneration = 0;
  initialized = false;
  runtimeReadyPromise = null;
  Object.assign(runtimeState, {
    bridgeEnabled: initialBridgeEnabled,
    bridgeUrl: initialBridgeUrl,
    bridgeStatus: initialBridgeStatus,
    snapshot: null,
    liveContext: null,
    isRefreshing: false,
    lastRefreshLabel: "Never",
    errorMessage: "",
    serverSessionRegistered: false,
    connectedSessionCount: 0,
    matchedBy: null,
    matchedServerSessionId: null,
    serverPendingCount: 0,
    serverSessionHealth: null,
    lastRegistrationSyncAt: null,
    lastRegistrationSyncState: "pending",
    paneVisibility: "unknown",
    sharedRuntimeAvailable: false,
    startupBehavior: "inactive",
    instanceId: null,
  } satisfies WordTaskpaneRuntimeSnapshot);
}
