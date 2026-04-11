<script lang="ts">
  import { onMount } from "svelte";
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
  } from "../../lib/bridge-adapter";
  import { bindOfficeDocumentHandler } from "../../lib/components/office-document-events";
  import {
    attachWordLiveContextBridge,
    WORD_TRACKING_MODE_CHANGED_EVENT,
  } from "../../lib/live-context";
  import { deriveTaskpaneDashboardView } from "../../lib/taskpane-connection";

  declare const Office: any;

  type BridgeSessionLookupResponse = {
    ok?: boolean;
    sessions?: {
      pendingCount?: number;
      health?: string;
      snapshot?: (BridgeSessionSnapshot & { app?: string }) | null;
    }[];
  };

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
      snapshot?.documentId != null && snapshot.documentId.trim().length > 0
        ? snapshot.documentId
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

  let controller: OfficeBridgeController | null = null;
  let snapshot: BridgeSessionSnapshot | null = null;
  let liveContext: BridgeLiveContext | null = null;
  let bridgeEnabled = isBridgeForcedEnabled();
  let bridgeUrl = resolveConfiguredBridgeUrl();
  let bridgeStatus: OfficeBridgeConnectionStatus = {
    enabled: bridgeEnabled,
    serverUrl: bridgeUrl,
    phase: bridgeEnabled ? "connecting" : "disabled",
    isConnected: false,
    hasConnected: false,
    lastError: null,
    sessionHealth: bridgeEnabled ? "registration_pending" : "orphaned",
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
  let isRefreshing = false;
  let lastRefreshLabel = "Never";
  let errorMessage = "";
  let serverSessionRegistered = false;
  let connectedSessionCount = 0;
  let matchedBy: "instanceId" | "documentId" | null = null;
  let matchedServerSessionId: string | null = null;
  let serverPendingCount = 0;
  let serverSessionHealth: string | null = null;
  let lastRegistrationSyncAt: number | null = null;
  let lastRegistrationSyncState: "matched" | "pending" | "error" = "pending";
  let activeRefresh: Promise<void> | null = null;
  $: dashboardView = deriveTaskpaneDashboardView({
    snapshot,
    bridgeStatus,
    serverSessionRegistered,
    connectedSessionCount,
    matchedBy,
    matchedServerSessionId,
    serverPendingCount,
    serverSessionHealth,
    lastRegistrationSyncAt,
    lastRegistrationSyncState,
  });

  let deferredRefreshTimer: number | null = null;
  let registrationSyncGeneration = 0;

  function clearMatchedServerState() {
    serverSessionRegistered = false;
    connectedSessionCount = 0;
    matchedBy = null;
    matchedServerSessionId = null;
    serverPendingCount = 0;
    serverSessionHealth = null;
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
    bridgeStatus = controller.getStatus();
  }

  async function syncSessionRegistration() {
    const generation = ++registrationSyncGeneration;
    try {
      const response = await fetch(
        resolveBridgeSessionsUrl(bridgeStatus.serverUrl || bridgeUrl),
      );
      if (generation !== registrationSyncGeneration) return;
      if (!response.ok) {
        clearMatchedServerState();
        lastRegistrationSyncAt = Date.now();
        lastRegistrationSyncState = "error";
        return;
      }
      const payload = (await response.json()) as BridgeSessionLookupResponse;
      if (generation !== registrationSyncGeneration) return;
      const { matchedSession, connectedWordSessions, matchedBy: nextMatchedBy } = resolveMatchedSession(
        payload.sessions ?? [],
      );
      connectedSessionCount = connectedWordSessions;
      serverSessionRegistered = Boolean(matchedSession?.snapshot);
      matchedBy = nextMatchedBy;
      matchedServerSessionId = matchedSession?.snapshot?.sessionId ?? null;
      serverPendingCount = matchedSession?.pendingCount ?? 0;
      serverSessionHealth = matchedSession?.health ?? null;
      lastRegistrationSyncAt = Date.now();
      lastRegistrationSyncState = matchedSession?.snapshot ? "matched" : "pending";
      if (
        matchedSession?.snapshot &&
        shouldAdoptServerSnapshot(snapshot, matchedSession.snapshot)
      ) {
        snapshot = matchedSession.snapshot;
        liveContext = matchedSession.snapshot.gateway?.liveContext ?? null;
      }
    } catch {
      if (generation !== registrationSyncGeneration) return;
      clearMatchedServerState();
      lastRegistrationSyncAt = Date.now();
      lastRegistrationSyncState = "error";
    }
  }

  function shouldShowRefreshBusyState(reason: string) {
    return reason === "manual refresh" || reason === "reconnect";
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
      isRefreshing = true;
    }
    errorMessage = "";
    syncBridgeStatus();

    activeRefresh = (async () => {
      try {
        const nextSnapshot = await controller.refresh();
        snapshot = nextSnapshot;
        liveContext = nextSnapshot?.gateway?.liveContext ?? null;
        lastRefreshLabel = new Date().toLocaleTimeString();
        syncBridgeStatus();
        await syncSessionRegistration();
      } catch (error) {
        errorMessage =
          error instanceof Error ? error.message : "Could not refresh Word bridge status.";
        syncBridgeStatus();
        await syncSessionRegistration();
      }
    })();

    try {
      await activeRefresh;
    } finally {
      activeRefresh = null;
      if (showBusyState) {
        isRefreshing = false;
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

  onMount(() => {
    let unsubscribeBridgeStatus = () => undefined;
    let detachBridgeEvents = () => undefined;
    let detachSelectionHandler = () => undefined;
    let statusPollTimer: number | null = null;

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

    try {
      const adapter = createWordBridgeAdapter({
        getRuntimeState: () => deriveWordTaskpaneRuntimeState(bridgeStatus),
      });
      controller = startOfficeBridge({
        app: "word",
        adapter,
        enabled: bridgeEnabled,
        serverUrl: bridgeUrl,
        forwardConsole: false,
      });
      bridgeStatus = controller.getStatus();
      unsubscribeBridgeStatus = controller.subscribe((status) => {
        bridgeStatus = status;
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
      void refreshStatus("startup");
    } catch (error) {
      errorMessage =
        error instanceof Error && error.message.trim()
          ? error.message
          : "Taskpane startup failed before the bridge client initialized.";
    }

    return () => {
      detachBridgeEvents();
      unsubscribeBridgeStatus();
      detachSelectionHandler();
      window.removeEventListener("focus", handleWindowFocus);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
      window.removeEventListener(
        WORD_TRACKING_MODE_CHANGED_EVENT,
        handleTrackingModeChange,
      );
      controller?.stop();
      controller = null;
      if (statusPollTimer !== null) {
        window.clearInterval(statusPollTimer);
      }
      if (deferredRefreshTimer !== null) {
        window.clearTimeout(deferredRefreshTimer);
      }
    };
  });
</script>

<div class="shell">
  <section class="hero">
    <div class="hero-top">
      <div>
        <p class="eyebrow">Word MCP Bridge</p>
        <h1>{dashboardView.headline}</h1>
        <p class="lede">
          {dashboardView.subtitle}
        </p>
      </div>
      <div class="hero-actions">
        <button class="primary-action" on:click={() => void refreshStatus()} disabled={!controller || isRefreshing}>
          {isRefreshing ? "Refreshing..." : "Refresh"}
        </button>
        {#if controller}
          <button class="secondary-action" on:click={async () => {
            if (!controller) return;
            await controller.reconnect();
            await waitForReconnectReady();
            await refreshStatus("reconnect");
          }} disabled={!controller || isRefreshing}>
            Reconnect
          </button>
        {/if}
        <span class:ready={dashboardView.tone === "ready"} class:working={dashboardView.tone === "working"} class:warning={dashboardView.tone === "warning"} class="status-pill">
          {dashboardView.statusLabel}
        </span>
      </div>
    </div>

    <div class="hero-footer">
      <span>{dashboardView.documentCard.summary}</span>
      <span>Last refreshed {lastRefreshLabel}</span>
    </div>
  </section>

  {#if errorMessage}
    <section class="panel panel-error">
      <strong>Latest problem</strong>
      <p>{errorMessage}</p>
    </section>
  {/if}

  {#if dashboardView.recoveryTitle}
    <section class="panel panel-recovery">
      <h2>{dashboardView.recoveryTitle}</h2>
      <p class="panel-lede">{dashboardView.recoveryText}</p>
      {#if dashboardView.showHelperHint}
        <p class="helper-hint">If this stays stuck, open Word MCP Bridge Helper.app and confirm Word is connected there.</p>
      {/if}
    </section>
  {/if}

  <div class="grid">
    <section class="panel">
      <span class="card-label">{dashboardView.documentCard.eyebrow}</span>
      <h2>{dashboardView.documentCard.title}</h2>
      <p class="panel-lede">{dashboardView.documentCard.summary}</p>
      {#if dashboardView.documentCard.detail}
        <p class="card-detail">{dashboardView.documentCard.detail}</p>
      {/if}
    </section>

    <section class="panel">
      <span class="card-label">{dashboardView.selectionCard.eyebrow}</span>
      <h2>{dashboardView.selectionCard.title}</h2>
      <p class="panel-lede">{dashboardView.selectionCard.summary}</p>
      {#if dashboardView.selectionCard.detail}
        <div class="selection-preview">{dashboardView.selectionCard.detail}</div>
      {/if}
    </section>

    <section class="panel">
      <span class="card-label">{dashboardView.reviewCard.eyebrow}</span>
      <h2>{dashboardView.reviewCard.title}</h2>
      <p class="panel-lede">{dashboardView.reviewCard.summary}</p>
      {#if dashboardView.reviewCard.detail}
        <p class="card-detail">{dashboardView.reviewCard.detail}</p>
      {/if}
    </section>

    <section class="panel panel-wide">
      <span class="card-label">{dashboardView.readinessCard.eyebrow}</span>
      <h2>{dashboardView.readinessCard.title}</h2>
      <p class="panel-lede">{dashboardView.readinessCard.summary}</p>
      <ul class="readiness-list">
        {#each dashboardView.readinessCard.items as item}
          <li>{item}</li>
        {/each}
      </ul>
      {#if dashboardView.readinessCard.warning}
        <p class="helper-hint">{dashboardView.readinessCard.warning}</p>
      {/if}
    </section>

    <section class="panel">
      <span class="card-label">{dashboardView.connectionCard.eyebrow}</span>
      <h2>{dashboardView.connectionCard.title}</h2>
      <p class="panel-lede">{dashboardView.connectionCard.summary}</p>
      {#if dashboardView.connectionCard.detail}
        <p class="card-detail">{dashboardView.connectionCard.detail}</p>
      {/if}
      <ul class="readiness-list compact-list">
        {#each dashboardView.connectionCard.items as item}
          <li>{item}</li>
        {/each}
      </ul>
    </section>

    <section class="panel">
      <span class="card-label">{dashboardView.diagnosticsCard.eyebrow}</span>
      <h2>{dashboardView.diagnosticsCard.title}</h2>
      <p class="panel-lede">{dashboardView.diagnosticsCard.summary}</p>
      <details class="diagnostics-drawer">
        <summary>Show recent connection breadcrumbs</summary>
        <ul class="readiness-list compact-list">
          {#each dashboardView.diagnosticsCard.items as item}
            <li>{item}</li>
          {/each}
        </ul>
      </details>
    </section>
  </div>
</div>

<style>
  .shell {
    min-height: 100vh;
    padding: 20px;
    box-sizing: border-box;
  }

  .hero {
    background: rgba(255, 255, 255, 0.86);
    border: 1px solid rgba(25, 40, 72, 0.12);
    border-radius: 20px;
    padding: 20px;
    box-shadow: 0 18px 44px rgba(41, 66, 135, 0.12);
    margin-bottom: 18px;
  }

  .hero-top {
    display: flex;
    justify-content: space-between;
    gap: 18px;
    align-items: flex-start;
    flex-wrap: wrap;
  }

  .eyebrow {
    margin: 0 0 8px;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    font-size: 11px;
    font-weight: 700;
    color: #4c67ad;
  }

  h1 {
    margin: 0;
    font-size: 24px;
  }

  .lede {
    margin: 10px 0 0;
    line-height: 1.5;
    color: #42506a;
  }

  .hero-actions {
    display: flex;
    gap: 12px;
    align-items: center;
    flex-wrap: wrap;
  }

  button {
    border: 0;
    border-radius: 999px;
    background: #2348d8;
    color: white;
    padding: 10px 16px;
    cursor: pointer;
    font-weight: 600;
  }

  .secondary-action {
    background: rgba(35, 72, 216, 0.1);
    color: #2348d8;
  }

  button:disabled {
    opacity: 0.6;
    cursor: default;
  }

  .status-pill {
    border-radius: 999px;
    padding: 8px 12px;
    background: rgba(25, 40, 72, 0.08);
    color: #40506d;
    font-size: 13px;
    font-weight: 600;
  }

  .status-pill.ready {
    background: rgba(39, 174, 96, 0.14);
    color: #176c42;
  }

  .status-pill.working {
    background: rgba(35, 72, 216, 0.12);
    color: #2348d8;
  }

  .status-pill.warning {
    background: rgba(154, 103, 0, 0.14);
    color: #9a6700;
  }

  .hero-footer {
    margin-top: 16px;
    display: flex;
    flex-wrap: wrap;
    gap: 12px 20px;
    font-size: 13px;
    color: #61708b;
  }

  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 16px;
  }

  .panel {
    background: rgba(255, 255, 255, 0.84);
    border: 1px solid rgba(25, 40, 72, 0.12);
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 10px 30px rgba(41, 66, 135, 0.1);
  }

  .panel-wide {
    grid-column: 1 / -1;
  }

  h2 {
    margin: 0 0 10px;
    font-size: 20px;
  }

  .panel-error {
    border-color: rgba(220, 38, 38, 0.2);
    background: rgba(255, 237, 237, 0.92);
    margin-bottom: 16px;
  }

  .panel-recovery {
    margin-bottom: 16px;
    border-color: rgba(35, 72, 216, 0.14);
    background: rgba(243, 246, 255, 0.92);
  }

  .panel-lede {
    margin: 0 0 14px;
    color: #61708b;
    font-size: 13px;
    line-height: 1.45;
  }

  .selection-preview {
    padding: 10px 12px;
    border-radius: 12px;
    background: rgba(25, 40, 72, 0.05);
    line-height: 1.45;
  }

  .card-label {
    display: block;
    margin-bottom: 10px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #61708b;
  }

  .card-detail,
  .helper-hint {
    margin: 0;
    color: #52617d;
    font-size: 13px;
    line-height: 1.5;
  }

  .helper-hint {
    margin-top: 12px;
  }

  .readiness-list {
    margin: 0;
    padding-left: 18px;
    display: grid;
    gap: 10px;
    color: #162033;
  }

  .compact-list {
    margin-top: 14px;
    gap: 8px;
    font-size: 13px;
  }

  .diagnostics-drawer summary {
    cursor: pointer;
    color: #2348d8;
    font-weight: 600;
  }

  .diagnostics-drawer[open] summary {
    margin-bottom: 12px;
  }
</style>
