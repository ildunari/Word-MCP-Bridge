import type { GatewayHostAdapter } from "./gateway-types.js";
import {
  type BridgeCapability,
  BRIDGE_PROTOCOL_VERSION,
  type BridgeEventMessage,
  type BridgeEventName,
  type BridgeEventPayloads,
  type BridgeGatewayState,
  type BridgeHostInfo,
  type BridgeRuntimeStateSlice,
  type BridgeSessionSnapshot,
  type BridgeToolDefinition,
  type BridgeToolExecutionResult,
  type BridgeVfsDeleteParams,
  type BridgeVfsEntry,
  type BridgeVfsListParams,
  type BridgeVfsReadParams,
  type BridgeVfsReadResult,
  type BridgeVfsWriteParams,
  type BridgeWireMessage,
  createBridgeId,
  extractToolError,
  extractToolImages,
  extractToolText,
  isBridgeInvokeMessage,
  normalizeBridgeCapabilities,
  normalizeBridgeUrl,
  serializeForJson,
  toBridgeError,
  uint8ArrayToBase64,
} from "./protocol.js";

declare const Office: any;
declare const Excel: any;
declare const PowerPoint: any;
declare const Word: any;

interface BridgeExecutableTool {
  name: string;
  label?: string;
  description?: string;
  parameters?: unknown;
  requiredCapability?: BridgeCapability;
  execute: (
    toolCallId: string,
    params: unknown,
    signal?: AbortSignal,
  ) => Promise<unknown>;
}

interface BridgeAdapter extends GatewayHostAdapter {
  tools: BridgeExecutableTool[];
  appName?: string;
  appVersion?: string;
  getRuntimeState?: () => BridgeRuntimeStateSlice | null;
}

interface BridgeVfsAdapter {
  snapshot: () => Promise<{ path: string; data: Uint8Array }[]>;
  readFile: (path: string) => Promise<string>;
  readFileBuffer: (path: string) => Promise<Uint8Array>;
  writeFile: (path: string, content: string | Uint8Array) => Promise<void>;
  deleteFile: (path: string) => Promise<void>;
}

export interface OfficeBridgeClientOptions {
  app: string;
  adapter: BridgeAdapter;
  vfs?: BridgeVfsAdapter;
  enabled?: boolean;
  serverUrl?: string;
  reconnectBaseMs?: number;
  reconnectMaxMs?: number;
  forwardConsole?: boolean;
  exposeGlobal?: boolean;
}

export interface OfficeBridgeController {
  readonly enabled: boolean;
  readonly instanceId: string;
  refresh: () => Promise<BridgeSessionSnapshot | null>;
  reconnect: () => Promise<BridgeSessionSnapshot | null>;
  emitEvent: <K extends BridgeEventName>(
    event: K,
    payload: BridgeEventPayloads[K],
  ) => void;
  getStatus: () => OfficeBridgeConnectionStatus;
  subscribe: (
    listener: (status: OfficeBridgeConnectionStatus) => void,
  ) => () => void;
  stop: () => void;
}

export type OfficeBridgeConnectionPhase =
  | "disabled"
  | "connecting"
  | "connected"
  | "reconnecting"
  | "disconnected";

export interface OfficeBridgeStatusError {
  message: string;
  at: number;
}

export type OfficeBridgeSessionHealth =
  | "live"
  | "registration_pending"
  | "stale"
  | "reconnecting"
  | "orphaned";

export interface OfficeBridgeInvokeTimeoutDetails {
  requestId: string;
  method: string;
  toolName?: string;
  at: number;
}

export interface OfficeBridgeConnectionDetails {
  lastWebSocketOpenAt: number | null;
  lastHelloSnapshotCapturedAt: number | null;
  lastHelloSentAt: number | null;
  lastSessionUpdatedAt: number | null;
  lastToolCompletedAt: number | null;
  lastTimedOutInvoke: OfficeBridgeInvokeTimeoutDetails | null;
  activeInvokeCount: number;
  currentSessionId: string | null;
  currentDocumentId: string | null;
  lastDocumentSwitchAt: number | null;
  lastReconnectAt: number | null;
}

export interface OfficeBridgeDiagnosticEntry {
  level: "info" | "warn" | "error";
  message: string;
  at: number;
}

export interface OfficeBridgeConnectionStatus {
  enabled: boolean;
  serverUrl: string;
  phase: OfficeBridgeConnectionPhase;
  isConnected: boolean;
  hasConnected: boolean;
  lastError: OfficeBridgeStatusError | null;
  sessionHealth: OfficeBridgeSessionHealth;
  details: OfficeBridgeConnectionDetails;
  diagnostics: OfficeBridgeDiagnosticEntry[];
}

interface PendingState {
  socket: WebSocket | null;
  stopped: boolean;
  reconnectTimer: number | null;
  reconnectDelayMs: number;
  snapshot: BridgeSessionSnapshot | null;
  status: OfficeBridgeConnectionStatus;
}

const BRIDGE_ENABLE_QUERY_KEY = "office_bridge";
const BRIDGE_URL_QUERY_KEY = "office_bridge_url";
const BRIDGE_ENABLE_STORAGE_KEY = "office-agents-bridge-enabled";
const BRIDGE_URL_STORAGE_KEY = "office-agents-bridge-url";
const BRIDGE_INSTANCE_PREFIX = "office-agents-bridge-instance";
const MAX_BRIDGE_DIAGNOSTICS = 12;

function getStoredInstanceId(app: string): string {
  const key = `${BRIDGE_INSTANCE_PREFIX}:${app}`;
  try {
    const existing = sessionStorage.getItem(key);
    if (existing) return existing;
    const next = createBridgeId(app).replace(/[^a-zA-Z0-9_-]/g, "_");
    sessionStorage.setItem(key, next);
    return next;
  } catch {
    return createBridgeId(app).replace(/[^a-zA-Z0-9_-]/g, "_");
  }
}

function isEnabledByDefault(): boolean {
  const params = new URLSearchParams(window.location.search);
  const query = params.get(BRIDGE_ENABLE_QUERY_KEY);
  if (query === "1" || query === "true") return true;
  if (query === "0" || query === "false") return false;

  try {
    const stored = localStorage.getItem(BRIDGE_ENABLE_STORAGE_KEY);
    if (stored === "true") return true;
    if (stored === "false") return false;
  } catch {
    // Ignore storage failures.
  }

  return window.location.hostname === "localhost";
}

function resolveServerUrl(explicitUrl?: string): string {
  if (explicitUrl) return normalizeBridgeUrl(explicitUrl, "ws");

  const params = new URLSearchParams(window.location.search);
  const query = params.get(BRIDGE_URL_QUERY_KEY);
  if (query) return normalizeBridgeUrl(query, "ws");

  try {
    const stored = localStorage.getItem(BRIDGE_URL_STORAGE_KEY);
    if (stored) return normalizeBridgeUrl(stored, "ws");
  } catch {
    // Ignore storage failures.
  }

  return normalizeBridgeUrl(undefined, "ws");
}

function getToolDefinitions(adapter: BridgeAdapter): BridgeToolDefinition[] {
  return ((adapter.tools ?? []) as BridgeExecutableTool[]).map((tool) => ({
    name: tool.name,
    label: tool.label,
    description: tool.description,
    parameters: serializeForJson(tool.parameters),
    requiredCapability: tool.requiredCapability ?? "tool_call",
  }));
}

async function getAdapterCapabilities(
  adapter: BridgeAdapter,
  vfsEnabled: boolean,
): Promise<BridgeCapability[]> {
  const declaredCapabilities = adapter.getCapabilities
    ? await Promise.resolve(adapter.getCapabilities()).catch(() => [])
    : [];
  return normalizeBridgeCapabilities([
    "observe",
    ...(adapter.tools.length > 0 ? ["tool_call"] : []),
    ...(vfsEnabled ? ["vfs_access"] : []),
    ...declaredCapabilities,
  ]);
}

async function captureSessionSnapshot(
  app: string,
  adapter: BridgeAdapter,
  instanceId: string,
  vfsEnabled: boolean,
  previous?: BridgeSessionSnapshot | null,
): Promise<BridgeSessionSnapshot> {
  const documentId = await adapter.getDocumentId();
  const meta = adapter.getDocumentMetadata
    ? await Promise.resolve(adapter.getDocumentMetadata()).catch(() => null)
    : null;
  const liveContext = adapter.getLiveContext
    ? await Promise.resolve(adapter.getLiveContext()).catch(() => null)
    : null;

  const officeContext =
    typeof Office === "undefined" ? undefined : Office?.context;
  const diagnostics = officeContext?.diagnostics;
  const hostInfo: BridgeHostInfo = {
    host: officeContext?.host ? String(officeContext.host) : undefined,
    platform: officeContext?.platform
      ? String(officeContext.platform)
      : undefined,
    officeVersion: diagnostics?.version,
    userAgent: navigator.userAgent,
    href: window.location.href,
    title: document.title,
  };

  const now = Date.now();
  const runtimeState = adapter.getRuntimeState?.() ?? undefined;
  const capabilities = await getAdapterCapabilities(adapter, vfsEnabled);
  const gateway: BridgeGatewayState = {
    capabilities,
    liveContext: liveContext ?? undefined,
  };

  return {
    sessionId: `${app}:${instanceId}`,
    instanceId,
    app,
    appName: adapter.appName,
    appVersion: adapter.appVersion,
    metadataTag: adapter.metadataTag,
    documentId,
    documentMetadata: meta?.metadata,
    tools: getToolDefinitions(adapter),
    host: hostInfo,
    runtimeState,
    gateway,
    connectedAt: previous?.connectedAt ?? now,
    updatedAt: now,
  };
}

function parseWireMessage(
  event: MessageEvent<string>,
): BridgeWireMessage | null {
  try {
    return JSON.parse(event.data) as BridgeWireMessage;
  } catch {
    return null;
  }
}

function scheduleMicrotask(action: () => void) {
  Promise.resolve()
    .then(action)
    .catch(() => undefined);
}

function bridgeConnectionErrorMessage(): string {
  return "Could not connect to the bridge server.";
}

export function startOfficeBridge(
  options: OfficeBridgeClientOptions,
): OfficeBridgeController {
  const enabled = options.enabled ?? isEnabledByDefault();
  const instanceId = getStoredInstanceId(options.app);

  const state: PendingState = {
    socket: null,
    stopped: !enabled,
    reconnectTimer: null,
    reconnectDelayMs: options.reconnectBaseMs ?? 1_000,
    snapshot: null,
    status: {
      enabled,
      serverUrl: resolveServerUrl(options.serverUrl),
      phase: enabled ? "connecting" : "disabled",
      isConnected: false,
      hasConnected: false,
      lastError: null,
      sessionHealth: enabled ? "registration_pending" : "orphaned",
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
    },
  };

  let queue = Promise.resolve<unknown>(undefined);
  let consoleRestore: (() => void) | null = null;
  const statusListeners = new Set<
    (status: OfficeBridgeConnectionStatus) => void
  >();

  const serverUrl = state.status.serverUrl;
  const reconnectBaseMs = options.reconnectBaseMs ?? 1_000;
  const reconnectMaxMs = options.reconnectMaxMs ?? 10_000;

  const cloneStatus = (
    status: OfficeBridgeConnectionStatus,
  ): OfficeBridgeConnectionStatus => ({
    ...status,
    details: { ...status.details },
    diagnostics: [...status.diagnostics],
  });

  const deriveSessionHealth = (
    status: OfficeBridgeConnectionStatus,
  ): OfficeBridgeSessionHealth => {
    if (!status.enabled) return "orphaned";
    if (status.phase === "reconnecting" || status.phase === "disconnected") {
      return "reconnecting";
    }
    if (!status.details.currentSessionId) {
      return status.hasConnected ? "registration_pending" : "orphaned";
    }
    if (!status.isConnected || status.phase === "connecting") {
      return "registration_pending";
    }
    if (status.details.lastTimedOutInvoke) {
      return "stale";
    }
    return "live";
  };

  const publishStatus = (next: Partial<OfficeBridgeConnectionStatus>) => {
    const merged: OfficeBridgeConnectionStatus = {
      ...state.status,
      ...next,
    };
    merged.sessionHealth = deriveSessionHealth(merged);
    const changed =
      merged.phase !== state.status.phase ||
      merged.isConnected !== state.status.isConnected ||
      merged.hasConnected !== state.status.hasConnected ||
      merged.lastError?.message !== state.status.lastError?.message ||
      merged.lastError?.at !== state.status.lastError?.at ||
      merged.sessionHealth !== state.status.sessionHealth ||
      merged.details !== state.status.details ||
      merged.diagnostics !== state.status.diagnostics;
    state.status = merged;
    if (!changed) return;
    for (const listener of statusListeners) {
      listener(cloneStatus(merged));
    }
  };

  const appendDiagnostic = (
    level: OfficeBridgeDiagnosticEntry["level"],
    message: string,
  ) => {
    publishStatus({
      diagnostics: [
        {
          level,
          message,
          at: Date.now(),
        },
        ...state.status.diagnostics,
      ].slice(0, MAX_BRIDGE_DIAGNOSTICS),
    });
  };

  const updateConnectionDetails = (
    next: Partial<OfficeBridgeConnectionDetails>,
  ) => {
    publishStatus({
      details: {
        ...state.status.details,
        ...next,
      },
    });
  };

  const resetSessionState = (reason: string) => {
    appendDiagnostic("warn", `Resetting bridge session state (${reason}).`);
    state.snapshot = null;
    publishStatus({
      lastError: null,
      details: {
        ...state.status.details,
        currentSessionId: null,
        currentDocumentId: null,
        lastSessionUpdatedAt: null,
        activeInvokeCount: 0,
      },
    });
  };

  const send = (message: BridgeWireMessage) => {
    if (!state.socket || state.socket.readyState !== WebSocket.OPEN) return;
    state.socket.send(JSON.stringify(message));
  };

  const sendEvent = (event: string, payload?: unknown) => {
    send({
      type: "event",
      event,
      ts: Date.now(),
      payload: serializeForJson(payload),
    } satisfies BridgeEventMessage);
  };

  const previousBridgeEventSink = options.adapter.bridgeEventSink;
    const bridgeEventSink = (event: string, payload: Record<string, unknown>) => {
    previousBridgeEventSink?.(event, payload);
    sendEvent(event, payload);
    if (event === "bridge_status") {
      const status = typeof payload.status === "string" ? payload.status : "";
      const source = typeof payload.source === "string" ? payload.source : undefined;
      if (status === "helper_poll") {
        appendDiagnostic(
          "info",
          `Helper poll tick completed${source ? ` from ${source}` : ""}.`,
        );
      }
      if (status === "taskpane_refresh") {
        appendDiagnostic(
          "info",
          `Taskpane focus refresh started${source ? ` from ${source}` : ""}.`,
        );
      }
    }
    if (
      event.startsWith("state:") ||
      event.startsWith("message:") ||
      event.startsWith("tool:")
    ) {
      scheduleMicrotask(() => {
        refresh(`bridge_event:${event}`).catch(() => undefined);
      });
    }
  };

  options.adapter.bridgeEventSink = bridgeEventSink;

  const requireCapability = async (capability: BridgeCapability) => {
    const capabilities = await getAdapterCapabilities(
      options.adapter,
      Boolean(options.vfs),
    );
    if (!capabilities.includes(capability)) {
      throw new Error(
        `Bridge capability '${capability}' is not enabled for this session`,
      );
    }
  };

  const refresh = async (reason = "refresh") => {
    appendDiagnostic("info", `Capturing session snapshot (${reason}).`);
    const previousDocumentId = state.snapshot?.documentId ?? null;
    const snapshot = await captureSessionSnapshot(
      options.app,
      options.adapter,
      instanceId,
      Boolean(options.vfs),
      state.snapshot,
    );
    state.snapshot = snapshot;
    updateConnectionDetails({
      currentSessionId: snapshot.sessionId,
      currentDocumentId: snapshot.documentId,
      lastSessionUpdatedAt: Date.now(),
      lastDocumentSwitchAt:
        previousDocumentId != null && previousDocumentId !== snapshot.documentId
          ? Date.now()
          : state.status.details.lastDocumentSwitchAt,
    });
    sendEvent("session_updated", snapshot);
    appendDiagnostic(
      "info",
      `Session snapshot ready for ${snapshot.sessionId} (${reason}).`,
    );
    return snapshot;
  };

  const executeTool = async (
    toolName: string,
    args: unknown,
  ): Promise<BridgeToolExecutionResult> => {
    const tool = ((options.adapter.tools ?? []) as BridgeExecutableTool[]).find(
      (candidate) => candidate.name === toolName,
    );
    if (!tool) {
      throw new Error(`Tool not found: ${toolName}`);
    }

    await requireCapability(tool.requiredCapability ?? "tool_call");

    const toolCallId = createBridgeId(toolName);
    const result = await tool.execute(toolCallId, args);
    const resultText = extractToolText(result);
    const images = extractToolImages(result);
    const error = extractToolError(result);
    const isError = Boolean(error);

    if (!isError) {
      options.adapter.onToolResult?.(toolCallId, resultText, false);
    }

    const executionResult: BridgeToolExecutionResult = {
      toolCallId,
      toolName,
      isError,
      result,
      resultText,
      images,
      error,
    };

    sendEvent("tool_executed", executionResult);
    updateConnectionDetails({
      lastToolCompletedAt: Date.now(),
      lastTimedOutInvoke: null,
    });
    scheduleMicrotask(() => {
      refresh(`after_tool:${toolName}`).catch((refreshError) => {
        sendEvent("bridge_warning", {
          message: "Failed to refresh session after tool execution",
          error: toBridgeError(refreshError),
        });
      });
    });

    return executionResult;
  };

  const decodeBase64 = (dataBase64: string): Uint8Array => {
    const binary = atob(dataBase64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  };

  const requireVfs = (): BridgeVfsAdapter => {
    if (!options.vfs) {
      throw new Error("Bridge VFS adapter is not configured for this app");
    }
    return options.vfs;
  };

  const listVfs = async (params: BridgeVfsListParams | undefined) => {
    await requireCapability("vfs_access");
    const files = await requireVfs().snapshot();
    const prefix = params?.prefix?.trim();
    const entries: BridgeVfsEntry[] = files
      .filter((file) => !prefix || file.path.startsWith(prefix))
      .map((file) => ({ path: file.path, byteLength: file.data.byteLength }))
      .sort((a, b) => a.path.localeCompare(b.path));
    sendEvent("vfs_listed", {
      prefix: prefix ?? null,
      count: entries.length,
    });
    return entries;
  };

  const readVfs = async (
    params: BridgeVfsReadParams,
  ): Promise<BridgeVfsReadResult> => {
    await requireCapability("vfs_access");
    const vfs = requireVfs();
    if (!params?.path) {
      throw new Error("Missing path for vfs_read");
    }

    const encoding = params.encoding === "text" ? "text" : "base64";
    if (encoding === "text") {
      const text = await vfs.readFile(params.path);
      const bytes = new TextEncoder().encode(text);
      const result: BridgeVfsReadResult = {
        path: params.path,
        encoding,
        byteLength: bytes.byteLength,
        text,
      };
      sendEvent("vfs_read", {
        path: params.path,
        encoding,
        byteLength: result.byteLength,
      });
      return result;
    }

    const data = await vfs.readFileBuffer(params.path);
    const result: BridgeVfsReadResult = {
      path: params.path,
      encoding,
      byteLength: data.byteLength,
      dataBase64: uint8ArrayToBase64(data),
    };
    sendEvent("vfs_read", {
      path: params.path,
      encoding,
      byteLength: result.byteLength,
    });
    return result;
  };

  const writeVfs = async (params: BridgeVfsWriteParams) => {
    await requireCapability("vfs_access");
    const vfs = requireVfs();
    if (!params?.path) {
      throw new Error("Missing path for vfs_write");
    }
    if (
      typeof params.text !== "string" &&
      typeof params.dataBase64 !== "string"
    ) {
      throw new Error("vfs_write requires either text or dataBase64");
    }

    const content =
      typeof params.text === "string"
        ? params.text
        : decodeBase64(params.dataBase64 as string);
    await vfs.writeFile(params.path, content);
    sendEvent("vfs_written", {
      path: params.path,
      byteLength:
        typeof content === "string"
          ? new TextEncoder().encode(content).byteLength
          : content.byteLength,
    });
    scheduleMicrotask(() => {
      refresh("after_vfs_write").catch(() => undefined);
    });
    return { success: true, path: params.path };
  };

  const deleteVfs = async (params: BridgeVfsDeleteParams) => {
    await requireCapability("vfs_access");
    if (!params?.path) {
      throw new Error("Missing path for vfs_delete");
    }
    await requireVfs().deleteFile(params.path);
    sendEvent("vfs_deleted", { path: params.path });
    scheduleMicrotask(() => {
      refresh("after_vfs_delete").catch(() => undefined);
    });
    return { success: true, path: params.path };
  };

  const executeUnsafeOfficeJs = async (params: {
    code?: string;
    explanation?: string;
  }) => {
    await requireCapability("unsafe_office_js");
    const code = params.code?.trim();
    if (!code) {
      throw new Error("Missing code for execute_unsafe_office_js");
    }

    const evaluate = async (context: unknown, appGlobal: unknown) => {
      const scope = {
        context,
        Office,
        app: appGlobal,
        Excel: typeof Excel === "undefined" ? undefined : Excel,
        PowerPoint: typeof PowerPoint === "undefined" ? undefined : PowerPoint,
        Word: typeof Word === "undefined" ? undefined : Word,
        window,
        document,
        console,
        fetch: window.fetch.bind(window),
        localStorage,
        sessionStorage,
        globalThis,
        setTimeout,
        clearTimeout,
        setInterval,
        clearInterval,
      };

      const fn = new Function(
        ...Object.keys(scope),
        `"use strict"; return (async () => {\n${code}\n})();`,
      ) as (...fnArgs: unknown[]) => Promise<unknown>;
      return await fn(...Object.values(scope));
    };

    let result: unknown;
    switch (options.app) {
      case "excel":
        result = await Excel.run(async (context: unknown) => {
          return await evaluate(context, Excel);
        });
        break;
      case "powerpoint":
        result = await PowerPoint.run(async (context: unknown) => {
          return await evaluate(context, PowerPoint);
        });
        break;
      case "word":
        result = await Word.run(async (context: unknown) => {
          return await evaluate(context, Word);
        });
        break;
      default:
        throw new Error(
          `Unsafe Office.js execution is not supported for app ${options.app}`,
        );
    }

    const executionResult = {
      mode: "unsafe" as const,
      app: options.app,
      result,
    };

    sendEvent("unsafe_office_js_executed", {
      explanation: params.explanation,
      result: executionResult,
    });
    scheduleMicrotask(() => {
      refresh("after_unsafe_office_js").catch((refreshError) => {
        sendEvent("bridge_warning", {
          message: "Failed to refresh session after unsafe Office.js execution",
          error: toBridgeError(refreshError),
        });
      });
    });

    return executionResult;
  };

  const QUEUE_TASK_TIMEOUT_MS = 60_000;

  const runQueued = <T>(work: () => Promise<T>): Promise<T> => {
    const timedWork = () =>
      Promise.race([
        work(),
        new Promise<never>((_, reject) =>
          setTimeout(
            () => reject(new Error("Bridge task timed out")),
            QUEUE_TASK_TIMEOUT_MS,
          ),
        ),
      ]);
    const task = queue.then(timedWork, timedWork);
    queue = task.then(
      () => undefined,
      () => undefined,
    );
    return task;
  };

  const handleInvoke = (message: BridgeWireMessage) => {
    if (!isBridgeInvokeMessage(message)) return;

    const describeInvoke = () => {
      if (message.method === "execute_tool") {
        const params =
          message.params && typeof message.params === "object"
            ? (message.params as { toolName?: unknown })
            : {};
        const toolName =
          typeof params.toolName === "string" && params.toolName.trim().length > 0
            ? params.toolName
            : "unknown";
        return `${message.method}:${toolName}`;
      }
      return message.method;
    };

    const startedAt = Date.now();
    appendDiagnostic(
      "info",
      `Received bridge invoke ${describeInvoke()} (${message.requestId}).`,
    );

    runQueued(async () => {
      updateConnectionDetails({
        activeInvokeCount: state.status.details.activeInvokeCount + 1,
      });
      try {
        let result: unknown;
        switch (message.method) {
          case "ping":
            result = {
              pong: true,
              now: Date.now(),
              sessionId:
                state.snapshot?.sessionId ?? `${options.app}:${instanceId}`,
            };
            break;
          case "get_session_snapshot":
            result =
              state.snapshot ?? (await refresh("invoke:get_session_snapshot"));
            break;
          case "refresh_session":
            result = await refresh("invoke:refresh_session");
            break;
          case "execute_tool": {
            const params = (message.params ?? {}) as {
              toolName?: string;
              args?: unknown;
            };
            if (!params.toolName) {
              throw new Error("Missing toolName for execute_tool");
            }
            result = await executeTool(params.toolName, params.args ?? {});
            break;
          }
          case "execute_unsafe_office_js": {
            result = await executeUnsafeOfficeJs(
              (message.params ?? {}) as {
                code?: string;
                explanation?: string;
              },
            );
            break;
          }
          case "vfs_list":
            result = await listVfs(
              (message.params ?? {}) as BridgeVfsListParams,
            );
            break;
          case "vfs_read":
            result = await readVfs(
              (message.params ?? {}) as BridgeVfsReadParams,
            );
            break;
          case "vfs_write":
            result = await writeVfs(
              (message.params ?? {}) as BridgeVfsWriteParams,
            );
            break;
          case "vfs_delete":
            result = await deleteVfs(
              (message.params ?? {}) as BridgeVfsDeleteParams,
            );
            break;
          default:
            throw new Error(`Unsupported bridge method: ${message.method}`);
        }

        send({
          type: "response",
          requestId: message.requestId,
          ok: true,
          result: serializeForJson(result),
        });
        updateConnectionDetails({
          activeInvokeCount: Math.max(0, state.status.details.activeInvokeCount - 1),
          lastTimedOutInvoke: null,
        });
        appendDiagnostic(
          "info",
          `Completed bridge invoke ${describeInvoke()} (${message.requestId}) in ${Date.now() - startedAt}ms.`,
        );
      } catch (error) {
        const errorMessage =
          error instanceof Error && error.message.trim().length > 0
            ? error.message
            : "Bridge invoke failed";
        appendDiagnostic(
          errorMessage.includes("timed out") ? "warn" : "error",
          `Bridge invoke ${describeInvoke()} (${message.requestId}) failed after ${Date.now() - startedAt}ms: ${errorMessage}`,
        );
        updateConnectionDetails({
          activeInvokeCount: Math.max(0, state.status.details.activeInvokeCount - 1),
          ...(errorMessage.includes("timed out")
            ? {
                lastTimedOutInvoke: {
                  requestId: message.requestId,
                  method: message.method,
                  toolName:
                    message.method === "execute_tool" &&
                    message.params &&
                    typeof message.params === "object" &&
                    typeof (message.params as { toolName?: unknown }).toolName === "string"
                      ? ((message.params as { toolName: string }).toolName)
                      : undefined,
                  at: Date.now(),
                },
              }
            : {}),
        });
        send({
          type: "response",
          requestId: message.requestId,
          ok: false,
          error: toBridgeError(error),
        });
      }
    }).catch((error) => {
      sendEvent("bridge_error", {
        message: "Unhandled bridge queue error",
        error: toBridgeError(error),
      });
    });
  };

  const clearReconnectTimer = () => {
    if (state.reconnectTimer !== null) {
      window.clearTimeout(state.reconnectTimer);
      state.reconnectTimer = null;
    }
  };

  const scheduleReconnect = (reason = "unspecified") => {
    clearReconnectTimer();
    appendDiagnostic(
      "warn",
      `Scheduling websocket reconnect in ${state.reconnectDelayMs}ms (${reason}).`,
    );
    state.reconnectTimer = window.setTimeout(() => {
      connect().catch(() => undefined);
    }, state.reconnectDelayMs);
    state.reconnectDelayMs = Math.min(
      state.reconnectDelayMs * 2,
      reconnectMaxMs,
    );
  };

  const connect = async () => {
    if (state.stopped) return;
    publishStatus({
      phase: state.status.hasConnected ? "reconnecting" : "connecting",
      isConnected: false,
    });
    appendDiagnostic("info", `Attempting websocket connection to ${serverUrl}.`);

    let socket: WebSocket;
    try {
      socket = new WebSocket(serverUrl);
    } catch (error) {
      const message =
        error instanceof Error && error.message.trim()
          ? error.message
          : bridgeConnectionErrorMessage();
      publishStatus({
        lastError: {
          message,
          at: Date.now(),
        },
      });
      appendDiagnostic(
        "error",
        `WebSocket constructor failed for ${serverUrl}: ${message}`,
      );
      scheduleReconnect("constructor_failed");
      return;
    }
    state.socket = socket;

    socket.addEventListener("open", () => {
      if (state.socket !== socket || state.stopped) return;
      appendDiagnostic("info", `WebSocket opened for ${serverUrl}.`);
      updateConnectionDetails({
        lastWebSocketOpenAt: Date.now(),
        lastTimedOutInvoke: null,
      });
      clearReconnectTimer();
      state.reconnectDelayMs = reconnectBaseMs;
      publishStatus({
        phase: "connected",
        isConnected: true,
        hasConnected: true,
        lastError: null,
      });
      appendDiagnostic("info", "Capturing hello snapshot for websocket open.");
      refresh("hello")
        .then((snapshot) => {
          if (state.socket !== socket || state.stopped) return;
          updateConnectionDetails({
            lastHelloSnapshotCapturedAt: Date.now(),
          });
          appendDiagnostic(
            "info",
            `Hello snapshot captured for ${snapshot.sessionId}.`,
          );
          send({
            type: "hello",
            role: "office-addin",
            protocolVersion: BRIDGE_PROTOCOL_VERSION,
            snapshot,
          });
          updateConnectionDetails({
            lastHelloSentAt: Date.now(),
            currentSessionId: snapshot.sessionId,
            currentDocumentId: snapshot.documentId,
          });
          appendDiagnostic("info", `Hello sent for ${snapshot.sessionId}.`);
          sendEvent("bridge_status", {
            status: "connected",
            serverUrl,
            sessionId: snapshot.sessionId,
          });
          appendDiagnostic(
            "info",
            `Bridge session connected for ${snapshot.sessionId}.`,
          );
        })
        .catch((error) => {
          if (state.socket !== socket || state.stopped) return;
          const message =
            error instanceof Error && error.message.trim()
              ? error.message
              : bridgeConnectionErrorMessage();
          publishStatus({
            lastError: {
              message,
              at: Date.now(),
            },
          });
          appendDiagnostic(
            "error",
            `Connected socket could not build a bridge snapshot: ${message}`,
          );
          sendEvent("bridge_error", {
            message: "Failed to build bridge snapshot",
            error: toBridgeError(error),
          });
          try {
            socket.close();
          } catch {
            publishStatus({
              phase: "reconnecting",
              isConnected: false,
            });
            scheduleReconnect("hello_snapshot_failed");
          }
        });
    });

    socket.addEventListener("message", (event) => {
      if (state.socket !== socket || state.stopped) return;
      const message = parseWireMessage(event as MessageEvent<string>);
      if (!message) return;
      handleInvoke(message);
    });

    socket.addEventListener("close", (event) => {
      if (state.socket !== socket) return;
      if (state.socket === socket) {
        state.socket = null;
      }
      if (state.stopped) {
        publishStatus({
          phase: "disconnected",
          isConnected: false,
        });
        return;
      }

      const closeCode =
        typeof (event as CloseEvent | undefined)?.code === "number"
          ? ` code=${(event as CloseEvent).code}`
          : "";
      const closeReason =
        typeof (event as CloseEvent | undefined)?.reason === "string" &&
        (event as CloseEvent).reason
          ? ` reason=${(event as CloseEvent).reason}`
          : "";
      appendDiagnostic(
        "warn",
        `WebSocket closed for ${serverUrl}.${closeCode}${closeReason}`,
      );
      publishStatus({
        phase: "reconnecting",
        isConnected: false,
      });
      scheduleReconnect("socket_closed");
    });

    socket.addEventListener("error", () => {
      if (state.socket !== socket || state.stopped) return;
      appendDiagnostic(
        "error",
        `WebSocket error while connecting to ${serverUrl}.`,
      );
      publishStatus({
        lastError: {
          message: bridgeConnectionErrorMessage(),
          at: Date.now(),
        },
      });
      socket.close();
    });
  };

  const setupForwarders = () => {
    const handleWindowError = (event: ErrorEvent) => {
      appendDiagnostic(
        "error",
        `Window error: ${event.message || "Unknown taskpane error"}`,
      );
      sendEvent("window_error", {
        message: event.message,
        filename: event.filename,
        lineno: event.lineno,
        colno: event.colno,
        error: toBridgeError(event.error),
      });
    };

    const handleUnhandledRejection = (event: PromiseRejectionEvent) => {
      const reason =
        event.reason instanceof Error
          ? event.reason.message
          : typeof event.reason === "string"
            ? event.reason
            : JSON.stringify(serializeForJson(event.reason));
      appendDiagnostic("error", `Unhandled rejection: ${reason}`);
      sendEvent("unhandled_rejection", {
        reason: serializeForJson(event.reason),
      });
    };

    const handleFocus = () => {
      appendDiagnostic("info", "Taskpane focus refresh started.");
      refresh("window_focus").catch(() => undefined);
    };

    const handleBeforeUnload = () => {
      appendDiagnostic("warn", "Detected stale page reload; notifying bridge before unload.");
      sendEvent("session:hmr_reload", {
        previousSessionId: state.snapshot?.sessionId,
      });
    };

    window.addEventListener("error", handleWindowError);
    window.addEventListener("unhandledrejection", handleUnhandledRejection);
    window.addEventListener("focus", handleFocus);
    window.addEventListener("beforeunload", handleBeforeUnload);

    if (options.forwardConsole !== false) {
      const methods = ["debug", "info", "log", "warn", "error"] as const;
      const originals = new Map<
        (typeof methods)[number],
        (...args: unknown[]) => void
      >();
      let forwarding = false;

      for (const method of methods) {
        const original = console[method].bind(console);
        originals.set(method, original);
        console[method] = ((...args: unknown[]) => {
          original(...args);
          if (forwarding) return;
          forwarding = true;
          try {
            sendEvent("console", {
              level: method,
              args: args.map((arg) => serializeForJson(arg)),
            });
          } finally {
            forwarding = false;
          }
        }) as (typeof console)[typeof method];
      }

      consoleRestore = () => {
        for (const method of methods) {
          const original = originals.get(method);
          if (original) {
            console[method] = original as (typeof console)[typeof method];
          }
        }
      };
    }

    return () => {
      window.removeEventListener("error", handleWindowError);
      window.removeEventListener(
        "unhandledrejection",
        handleUnhandledRejection,
      );
      window.removeEventListener("focus", handleFocus);
      window.removeEventListener("beforeunload", handleBeforeUnload);
      consoleRestore?.();
      consoleRestore = null;
    };
  };

  let teardown = () => undefined;
  if (enabled) {
    appendDiagnostic("info", `Bridge client enabled for ${serverUrl}.`);
    teardown = setupForwarders();
    connect().catch(() => undefined);
  }

  const controller: OfficeBridgeController = {
    enabled,
    instanceId,
    refresh: async () => {
      if (!enabled || state.stopped) return null;
      return refresh("controller.refresh");
    },
    reconnect: async () => {
      if (!enabled) return null;
      clearReconnectTimer();
      updateConnectionDetails({
        lastReconnectAt: Date.now(),
      });
      resetSessionState("controller.reconnect");
      if (state.socket) {
        const activeSocket = state.socket;
        state.socket = null;
        try {
          activeSocket.close();
        } catch {
          // Ignore close failures during forced reconnect.
        }
      }
      state.stopped = false;
      publishStatus({
        phase: "connecting",
        isConnected: false,
      });
      connect().catch(() => undefined);
      return null;
    },
    emitEvent: <K extends BridgeEventName>(
      event: K,
      payload: BridgeEventPayloads[K],
    ) => {
      if (!enabled || state.stopped) return;
      bridgeEventSink(event, serializeForJson(payload) as Record<string, unknown>);
    },
    getStatus: () => cloneStatus(state.status),
    subscribe: (listener) => {
      statusListeners.add(listener);
      listener(cloneStatus(state.status));
      return () => {
        statusListeners.delete(listener);
      };
    },
    stop: () => {
      state.stopped = true;
      clearReconnectTimer();
      teardown();
      if (options.adapter.bridgeEventSink === bridgeEventSink) {
        delete options.adapter.bridgeEventSink;
      }
      if (state.socket) {
        state.socket.close();
        state.socket = null;
      }
      publishStatus({
        phase: enabled ? "disconnected" : "disabled",
        isConnected: false,
      });
      if (
        (
          window as typeof window & {
            __OFFICE_BRIDGE__?: OfficeBridgeController;
          }
        ).__OFFICE_BRIDGE__ === controller
      ) {
        delete (
          window as typeof window & {
            __OFFICE_BRIDGE__?: OfficeBridgeController;
          }
        ).__OFFICE_BRIDGE__;
      }
    },
  };

  if (enabled && options.exposeGlobal !== false) {
    (
      window as typeof window & { __OFFICE_BRIDGE__?: OfficeBridgeController }
    ).__OFFICE_BRIDGE__ = controller;
  }

  return controller;
}
