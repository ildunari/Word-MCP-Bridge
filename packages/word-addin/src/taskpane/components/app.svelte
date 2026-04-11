<script lang="ts">
  import { onMount } from "svelte";
  import { deriveTaskpaneDashboardView } from "../../lib/taskpane-connection";
  import {
    ensureWordTaskpaneRuntime,
    getWordTaskpaneRuntimeSnapshot,
    hideWordTaskpaneRuntime,
    reconnectWordTaskpaneRuntime,
    refreshWordTaskpaneRuntime,
    subscribeWordTaskpaneRuntime,
    type WordTaskpaneRuntimeSnapshot,
  } from "../../lib/shared-runtime";

  let runtime: WordTaskpaneRuntimeSnapshot = getWordTaskpaneRuntimeSnapshot();
  $: dashboardView = deriveTaskpaneDashboardView({
    snapshot: runtime.snapshot,
    bridgeStatus: runtime.bridgeStatus,
    serverSessionRegistered: runtime.serverSessionRegistered,
    connectedSessionCount: runtime.connectedSessionCount,
    matchedBy: runtime.matchedBy,
    matchedServerSessionId: runtime.matchedServerSessionId,
    serverPendingCount: runtime.serverPendingCount,
    serverSessionHealth: runtime.serverSessionHealth,
    lastRegistrationSyncAt: runtime.lastRegistrationSyncAt,
    lastRegistrationSyncState: runtime.lastRegistrationSyncState,
    paneVisibility: runtime.paneVisibility,
    sharedRuntimeAvailable: runtime.sharedRuntimeAvailable,
    startupBehavior: runtime.startupBehavior,
  });

  onMount(() => {
    void ensureWordTaskpaneRuntime();
    const unsubscribe = subscribeWordTaskpaneRuntime((next) => {
      runtime = next;
    });
    return () => {
      unsubscribe();
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
        <button class="primary-action" on:click={() => void refreshWordTaskpaneRuntime()} disabled={runtime.isRefreshing}>
          {runtime.isRefreshing ? "Refreshing..." : "Refresh"}
        </button>
        {#if runtime.instanceId}
          <button class="secondary-action" on:click={() => void reconnectWordTaskpaneRuntime()} disabled={runtime.isRefreshing}>
            Reconnect
          </button>
        {/if}
        {#if runtime.sharedRuntimeAvailable && dashboardView.statusLabel === "Connected" && runtime.paneVisibility !== "hidden"}
          <button class="secondary-action" on:click={() => void hideWordTaskpaneRuntime()} disabled={runtime.isRefreshing}>
            Hide Panel
          </button>
        {/if}
        <span class:ready={dashboardView.tone === "ready"} class:working={dashboardView.tone === "working"} class:warning={dashboardView.tone === "warning"} class="status-pill">
          {dashboardView.statusLabel}
        </span>
      </div>
    </div>

    <div class="hero-footer">
      <span>{dashboardView.documentCard.summary}</span>
      <span>Last refreshed {runtime.lastRefreshLabel}</span>
    </div>
  </section>

  {#if runtime.errorMessage}
    <section class="panel panel-error">
      <strong>Latest problem</strong>
      <p>{runtime.errorMessage}</p>
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
