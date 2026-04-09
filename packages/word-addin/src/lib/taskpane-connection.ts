import type {
  OfficeBridgeConnectionStatus,
} from "@word-mcp-bridge/bridge/client";
import type { BridgeSessionSnapshot } from "@word-mcp-bridge/bridge/protocol";

export interface TaskpaneConnectionView {
  headline: string;
  subtitle: string;
  statusLabel: string;
}

interface TaskpaneConnectionViewOptions {
  snapshot: BridgeSessionSnapshot | null;
  bridgeStatus: OfficeBridgeConnectionStatus;
}

export function deriveTaskpaneConnectionView(
  options: TaskpaneConnectionViewOptions,
): TaskpaneConnectionView {
  const { snapshot, bridgeStatus } = options;
  const lastErrorDetail = bridgeStatus.lastError?.message
    ? ` Last error: ${bridgeStatus.lastError.message}`
    : "";

  if (!bridgeStatus.enabled) {
    return {
      headline: "Bridge disabled",
      statusLabel: "Disabled",
      subtitle:
        "The Word bridge client is disabled for this taskpane. Re-enable it in query settings or local storage to attach MCP tools.",
    };
  }

  if (snapshot && bridgeStatus.phase === "connected") {
    return {
      headline: "Connected to Word",
      statusLabel: "Connected",
      subtitle:
        "Your current document and live selection are ready for AI tools.",
    };
  }

  if (snapshot && bridgeStatus.phase === "reconnecting") {
    return {
      headline: "Word is ready",
      statusLabel: "Reconnecting",
      subtitle:
        "Word document data is available locally, but the bridge is trying to reconnect before MCP tools can attach." +
        lastErrorDetail,
    };
  }

  if (!snapshot && bridgeStatus.phase === "connected") {
    return {
      headline: "Waiting for Word",
      statusLabel: "Bridge connected",
      subtitle:
        "The bridge server is connected. Open the Word MCP Bridge taskpane in your document to register a live session.",
    };
  }

  if (bridgeStatus.phase === "connecting" || bridgeStatus.phase === "reconnecting") {
    return {
      headline: "Waiting for bridge",
      statusLabel:
        bridgeStatus.phase === "connecting" ? "Connecting" : "Reconnecting",
      subtitle:
        "The taskpane is trying to reach the local bridge server. Start the helper or bridge server if it is not running yet." +
        lastErrorDetail,
    };
  }

  if (bridgeStatus.phase === "disconnected") {
    return {
      headline: snapshot ? "Word is ready" : "Bridge disconnected",
      statusLabel: "Disconnected",
      subtitle:
        "The taskpane bridge client is disconnected. Refresh or reopen the taskpane to reconnect." +
        lastErrorDetail,
    };
  }

  if (bridgeStatus.lastError) {
    return {
      headline: snapshot ? "Word is ready" : "Waiting for bridge",
      statusLabel: "Bridge unavailable",
      subtitle: bridgeStatus.lastError.message,
    };
  }

  return {
    headline: snapshot ? "Word is ready" : "Waiting for Word",
    statusLabel: "Waiting",
    subtitle:
      "Open the helper app, start the bridge, then reopen or refresh this task pane.",
  };
}
