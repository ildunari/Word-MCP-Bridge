<script lang="ts">
  import { onMount } from "svelte";
  import {
    type OfficeBridgeController,
    startOfficeBridge,
  } from "@word-mcp-bridge/bridge/client";
  import type {
    BridgeLiveContext,
    BridgeSessionSnapshot,
  } from "@word-mcp-bridge/bridge/protocol";
  import {
    createWordBridgeAdapter,
    resolveConfiguredBridgeUrl,
  } from "../../lib/bridge-adapter";
  import { bindOfficeDocumentHandler } from "../../lib/components/office-document-events";
  import {
    attachWordLiveContextBridge,
    WORD_TRACKING_MODE_CHANGED_EVENT,
  } from "../../lib/live-context";

  declare const Office: any;

  type ActivityItem = {
    message: string;
    at: string;
  };

  let controller: OfficeBridgeController | null = null;
  let snapshot: BridgeSessionSnapshot | null = null;
  let liveContext: BridgeLiveContext | null = null;
  let bridgeEnabled = true;
  let bridgeUrl = resolveConfiguredBridgeUrl();
  let isRefreshing = false;
  let lastRefreshLabel = "Never";
  let errorMessage = "";
  let activity: ActivityItem[] = [];
  const localSetupCommands = [
    "pnpm setup:word",
    "pnpm bridge:serve",
    "pnpm dev-server:word",
    "pnpm start:word",
  ];
  const hostedModeCommands = [
    "office-bridge serve",
    "office-bridge status",
    "office-bridge mcp-serve",
  ];

  function log(message: string) {
    activity = [
      { message, at: new Date().toLocaleTimeString() },
      ...activity,
    ].slice(0, 8);
  }

  async function refreshStatus(reason = "manual refresh") {
    if (!controller) return;

    isRefreshing = true;
    errorMessage = "";

    try {
      const nextSnapshot = await controller.refresh();
      snapshot = nextSnapshot;
      liveContext = nextSnapshot?.gateway?.liveContext ?? null;
      lastRefreshLabel = new Date().toLocaleTimeString();
      log(`Updated status after ${reason}.`);
    } catch (error) {
      errorMessage =
        error instanceof Error ? error.message : "Could not refresh Word bridge status.";
      log(`Refresh failed after ${reason}.`);
    } finally {
      isRefreshing = false;
    }
  }

  function metadataValue(key: string): string {
    const metadata = snapshot?.documentMetadata;
    if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) {
      return "n/a";
    }

    const value = (metadata as Record<string, unknown>)[key];
    if (value == null || value === "") return "n/a";
    return String(value);
  }

  function selectionPreview(): string {
    if (!liveContext?.selection?.hasSelection) return "No current selection";
    return liveContext.selection.selectedText ?? "Selection exists";
  }

  onMount(() => {
    const adapter = createWordBridgeAdapter();
    controller = startOfficeBridge({
      app: "word",
      adapter,
      enabled: bridgeEnabled,
      serverUrl: bridgeUrl,
      forwardConsole: false,
    });

    const detachBridgeEvents = attachWordLiveContextBridge(controller);
    const officeDocument =
      typeof Office === "undefined" ? undefined : Office?.context?.document;
    const detachSelectionHandler = bindOfficeDocumentHandler(
      officeDocument,
      typeof Office === "undefined"
        ? "DocumentSelectionChanged"
        : (Office?.EventType?.DocumentSelectionChanged ?? "DocumentSelectionChanged"),
      () => {
        log("Word selection changed.");
        void refreshStatus("selection change");
      },
    );

    const handleWindowFocus = () => {
      log("Taskpane regained focus.");
      void refreshStatus("window focus");
    };
    const handleVisibilityChange = () => {
      log("Document visibility changed.");
      void refreshStatus("visibility change");
    };
    const handleTrackingModeChange = () => {
      log("Word tracking mode changed.");
      void refreshStatus("tracking mode change");
    };

    window.addEventListener("focus", handleWindowFocus);
    document.addEventListener("visibilitychange", handleVisibilityChange);
    window.addEventListener(
      WORD_TRACKING_MODE_CHANGED_EVENT,
      handleTrackingModeChange,
    );

    log(
      controller.enabled
        ? `Bridge client started for ${bridgeUrl}.`
        : "Bridge client is disabled by query or local storage.",
    );
    void refreshStatus("startup");

    return () => {
      detachBridgeEvents();
      detachSelectionHandler();
      window.removeEventListener("focus", handleWindowFocus);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
      window.removeEventListener(
        WORD_TRACKING_MODE_CHANGED_EVENT,
        handleTrackingModeChange,
      );
      controller?.stop();
      controller = null;
    };
  });
</script>

<div class="shell">
  <section class="hero">
    <p class="eyebrow">Word MCP Bridge</p>
    <h1>Minimal Word connector</h1>
    <p class="lede">
      This add-in keeps a live bridge session open so local CLI and MCP clients can inspect
      Word context and run privileged Office.js commands.
    </p>
    <div class="hero-actions">
      <button on:click={() => void refreshStatus()} disabled={!controller || isRefreshing}>
        {isRefreshing ? "Refreshing..." : "Refresh"}
      </button>
      <span class:healthy={Boolean(snapshot)} class="status-pill">
        {snapshot ? "Connected" : controller?.enabled ? "Waiting for bridge" : "Disabled"}
      </span>
    </div>
  </section>

  {#if errorMessage}
    <section class="panel panel-error">
      <strong>Refresh error</strong>
      <p>{errorMessage}</p>
    </section>
  {/if}

  <div class="grid">
    {#if !snapshot}
      <section class="panel panel-wide setup-panel">
        <h2>How to connect</h2>
        <p class="caption">
          This taskpane is waiting for a live bridge session. Start the bridge server, launch
          the Word add-in, and then reopen or refresh this pane.
        </p>
        <div class="setup-grid">
          <div>
            <h3>Local developer mode</h3>
            <ol class="steps">
              {#each localSetupCommands as command}
                <li><code>{command}</code></li>
              {/each}
            </ol>
          </div>
          <div>
            <h3>Hosted add-in mode</h3>
            <ol class="steps">
              {#each hostedModeCommands as command}
                <li><code>{command}</code></li>
              {/each}
            </ol>
          </div>
        </div>
        <p class="caption">
          Once Word is connected, MCP hosts can attach through
          <code>office-bridge mcp-serve</code>.
        </p>
      </section>
    {/if}

    <section class="panel">
      <h2>Bridge</h2>
      <dl>
        <div>
          <dt>Enabled</dt>
          <dd>{controller?.enabled ? "yes" : "no"}</dd>
        </div>
        <div>
          <dt>Bridge URL</dt>
          <dd>{bridgeUrl}</dd>
        </div>
        <div>
          <dt>Session ID</dt>
          <dd>{snapshot?.sessionId ?? "pending"}</dd>
        </div>
        <div>
          <dt>Instance ID</dt>
          <dd>{controller?.instanceId ?? "pending"}</dd>
        </div>
        <div>
          <dt>Last refresh</dt>
          <dd>{lastRefreshLabel}</dd>
        </div>
      </dl>
    </section>

    <section class="panel">
      <h2>Document</h2>
      <dl>
        <div>
          <dt>Document ID</dt>
          <dd>{snapshot?.documentId ?? "pending"}</dd>
        </div>
        <div>
          <dt>Title</dt>
          <dd>{metadataValue("title")}</dd>
        </div>
        <div>
          <dt>Tracking mode</dt>
          <dd>{liveContext?.trackingMode ?? metadataValue("trackingMode")}</dd>
        </div>
        <div>
          <dt>Paragraphs</dt>
          <dd>{metadataValue("paragraphCount")}</dd>
        </div>
        <div>
          <dt>Words</dt>
          <dd>{metadataValue("wordCount")}</dd>
        </div>
        <div>
          <dt>Characters</dt>
          <dd>{metadataValue("characterCount")}</dd>
        </div>
      </dl>
    </section>

    <section class="panel">
      <h2>Live context</h2>
      <dl>
        <div>
          <dt>Focus target</dt>
          <dd>{liveContext?.focusTarget ?? "unknown"}</dd>
        </div>
        <div>
          <dt>Selection</dt>
          <dd>{selectionPreview()}</dd>
        </div>
        <div>
          <dt>Selection style</dt>
          <dd>{liveContext?.selection?.selectedStyle ?? "n/a"}</dd>
        </div>
        <div>
          <dt>Updated at</dt>
          <dd>
            {liveContext?.updatedAt ? new Date(liveContext.updatedAt).toLocaleTimeString() : "n/a"}
          </dd>
        </div>
      </dl>
    </section>

    <section class="panel">
      <h2>Capabilities</h2>
      {#if snapshot?.gateway?.capabilities?.length}
        <ul class="chip-list">
          {#each snapshot.gateway.capabilities as capability}
            <li>{capability}</li>
          {/each}
        </ul>
      {:else}
        <p class="muted">No capabilities advertised yet.</p>
      {/if}
      <p class="caption">
        This minimal add-in enables live observation plus privileged raw Office.js execution.
      </p>
    </section>

    <section class="panel panel-wide">
      <h2>Recent activity</h2>
      {#if activity.length}
        <ul class="activity-list">
          {#each activity as item}
            <li>
              <span>{item.message}</span>
              <time>{item.at}</time>
            </li>
          {/each}
        </ul>
      {:else}
        <p class="muted">No activity yet.</p>
      {/if}
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
    margin-top: 16px;
  }

  button {
    border: 0;
    border-radius: 999px;
    background: #2d5bff;
    color: white;
    padding: 10px 16px;
    cursor: pointer;
    font-weight: 600;
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

  .status-pill.healthy {
    background: rgba(39, 174, 96, 0.14);
    color: #176c42;
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

  .setup-panel {
    display: grid;
    gap: 14px;
  }

  .setup-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
  }

  h3 {
    margin: 0 0 10px;
    font-size: 15px;
  }

  .panel-error {
    border-color: rgba(220, 38, 38, 0.2);
    background: rgba(255, 237, 237, 0.92);
    margin-bottom: 16px;
  }

  h2 {
    margin: 0 0 14px;
    font-size: 16px;
  }

  dl {
    display: grid;
    gap: 12px;
    margin: 0;
  }

  dl div {
    display: grid;
    gap: 4px;
  }

  dt {
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #6a7892;
  }

  dd {
    margin: 0;
    word-break: break-word;
    color: #162033;
  }

  .chip-list {
    list-style: none;
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    padding: 0;
    margin: 0;
  }

  .chip-list li {
    border-radius: 999px;
    padding: 7px 11px;
    background: rgba(45, 91, 255, 0.1);
    color: #2447c5;
    font-size: 13px;
    font-weight: 600;
  }

  .activity-list {
    list-style: none;
    padding: 0;
    margin: 0;
    display: grid;
    gap: 10px;
  }

  .activity-list li {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    font-size: 14px;
  }

  .activity-list time,
  .caption,
  .muted {
    color: #61708b;
  }

  .steps {
    margin: 0;
    padding-left: 18px;
    display: grid;
    gap: 8px;
  }

  code {
    font-family:
      ui-monospace, SFMono-Regular, SFMono-Regular, Menlo, Monaco, Consolas,
      "Liberation Mono", "Courier New", monospace;
    font-size: 12px;
  }

  .caption {
    margin: 14px 0 0;
    font-size: 13px;
    line-height: 1.45;
  }

  .muted {
    margin: 0;
  }
</style>
