import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { type BridgeToolExecutionResult, serializeForJson } from "./protocol.js";
import { requestJson, type BridgeRequestOptions } from "./http-client.js";
import {
  findMatchingSession,
  type BridgeServerStatus,
  type BridgeSessionRecord,
} from "./server.js";
import {
  describeSessionChoice,
  getSessionDocumentLabel,
  getSessionDocumentSummary,
  isUnsavedWordDocument,
} from "./session-labels.js";
import {
  getFirstClassWordToolContracts,
  type WordToolContract,
} from "./word-tool-contracts.js";
import { z } from "zod";

function toStructuredRecord(value: unknown): Record<string, unknown> {
  const serialized = serializeForJson(value);
  if (serialized && typeof serialized === "object" && !Array.isArray(serialized)) {
    return serialized as Record<string, unknown>;
  }
  return { value: serialized };
}

function buildJsonResult(data: unknown) {
  const structuredContent = toStructuredRecord(data);
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(serializeForJson(data), null, 2),
      },
    ],
    structuredContent,
  };
}

function nestedStructuredContent(result: unknown): Record<string, unknown> | null {
  if (!result || typeof result !== "object" || Array.isArray(result)) {
    return null;
  }

  const candidate = (result as { structuredContent?: unknown }).structuredContent;
  if (candidate && typeof candidate === "object" && !Array.isArray(candidate)) {
    return serializeForJson(candidate) as Record<string, unknown>;
  }

  return null;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim()) {
    return error.message.trim();
  }
  return String(error || "Unknown bridge error");
}

export function describeBridgeConnectionFailure(
  error: unknown,
  baseUrl = "https://localhost:4017",
): string {
  const message = errorMessage(error);
  const code =
    typeof error === "object" && error && "code" in error
      ? String((error as { code?: string }).code ?? "")
      : "";
  const lower = message.toLowerCase();

  if (
    code === "ECONNREFUSED" ||
    lower.includes("econnrefused") ||
    lower.includes("cannot connect") ||
    lower.includes("fetch failed")
  ) {
    return `Word MCP Bridge is not reachable at ${baseUrl}. Start the local bridge server, then open Word and the Word MCP Bridge taskpane before retrying.`;
  }
  if (lower.includes("timed out")) {
    return `Word MCP Bridge at ${baseUrl} did not respond in time. Make sure the bridge server is running locally and try again.`;
  }
  if (lower.includes("unauthorized") || lower.includes("forbidden")) {
    return "Word MCP Bridge rejected the request. Make sure the MCP host is using the same bridge auth token as the local bridge server.";
  }
  return `Word MCP Bridge request failed: ${message}`;
}

export function describeMissingBridgeSession() {
  return "Word MCP Bridge is running, but no live Office sessions are connected. Open Word and the Word MCP Bridge taskpane, then retry.";
}

function shouldRewriteAsConnectionFailure(error: unknown): boolean {
  const message = errorMessage(error).toLowerCase();
  const code =
    typeof error === "object" && error && "code" in error
      ? String((error as { code?: string }).code ?? "")
      : "";
  return (
    code === "ECONNREFUSED" ||
    message.includes("econnrefused") ||
    message.includes("cannot connect") ||
    message.includes("fetch failed") ||
    message.includes("timed out") ||
    message.includes("unauthorized") ||
    message.includes("forbidden")
  );
}

export function bridgeToolExecutionResultToMcpResult(
  result: BridgeToolExecutionResult,
) {
  const content = [];
  if (result.resultText.trim()) {
    content.push({ type: "text" as const, text: result.resultText });
  }
  for (const image of result.images) {
    content.push({
      type: "image" as const,
      data: image.data,
      mimeType: image.mimeType,
    });
  }
  if (content.length === 0) {
    content.push({
      type: "text" as const,
      text: JSON.stringify(serializeForJson(result.result), null, 2),
    });
  }
  return {
    content,
    structuredContent:
      nestedStructuredContent(result.result) ?? toStructuredRecord(result.result),
    isError: result.isError,
  };
}

function wordDocumentEntry(session: BridgeSessionRecord) {
  return {
    sessionId: session.snapshot.sessionId,
    instanceId: session.snapshot.instanceId,
    title: getSessionDocumentLabel(session.snapshot),
    summary: getSessionDocumentSummary(session.snapshot),
    documentId: session.snapshot.documentId,
    isUnsaved: isUnsavedWordDocument(session.snapshot),
    isActive: session.snapshot.gateway?.liveContext?.focusTarget === "document",
    capabilities: session.snapshot.gateway?.capabilities ?? [],
    toolCount: session.snapshot.tools.length,
    connectedAt: session.snapshot.connectedAt,
    updatedAt: session.snapshot.updatedAt,
  };
}

export function buildBridgeStatusSummary(status: BridgeServerStatus) {
  return {
    startedAt: status.startedAt,
    uptimeMs: status.uptimeMs,
    host: status.host,
    port: status.port,
    httpUrl: status.httpUrl,
    wsUrl: status.wsUrl,
    sessionCount: status.sessionCount,
    totals: status.totals,
    sessions: status.sessions.map((session) => ({
      sessionId: session.snapshot.sessionId,
      instanceId: session.snapshot.instanceId,
      app: session.snapshot.app,
      title: getSessionDocumentLabel(session.snapshot),
      summary: getSessionDocumentSummary(session.snapshot),
      documentId: session.snapshot.documentId,
      health: session.health,
      pendingCount: session.pendingCount,
      toolCount: session.snapshot.tools.length,
      connectedAt: session.connectedAt,
      lastSeenAt: session.lastSeenAt,
      updatedAt: session.snapshot.updatedAt,
    })),
  };
}

export async function createOfficeBridgeMcpServer(
  options: BridgeRequestOptions = {},
) {
  const server = new McpServer(
    {
      name: "office-bridge",
      version: "0.0.2",
    },
    {
      capabilities: {
        tools: {},
      },
      instructions:
        "Use this server to inspect live Office bridge sessions, read Word live context, execute bridge tools, and optionally run privileged Office.js actions when the target session allows them.",
    },
  );

  async function bridgeRequest<T>(
    method: string,
    pathname: string,
    body?: unknown,
  ): Promise<T> {
    try {
      return await requestJson<T>(method, pathname, body, options);
    } catch (error) {
      if (shouldRewriteAsConnectionFailure(error)) {
        throw new Error(describeBridgeConnectionFailure(error, options.baseUrl));
      }
      if (error instanceof Error) {
        throw error;
      }
      throw new Error(errorMessage(error));
    }
  }

  async function fetchSessions() {
    const response = await bridgeRequest<{
      ok: true;
      sessions: BridgeSessionRecord[];
    }>("GET", "/sessions");
    return response.sessions;
  }

  function withDisplayMetadata(sessions: BridgeSessionRecord[]) {
    return sessions.map((session) => ({
      ...session,
      display: {
        label: getSessionDocumentLabel(session.snapshot),
        summary: getSessionDocumentSummary(session.snapshot),
        selectorHint: describeSessionChoice(session.snapshot),
      },
    }));
  }

  async function fetchBridgeStatus() {
    const response = await bridgeRequest<{
      ok: true;
      status: unknown;
    }>("GET", "/status");
    return response.status;
  }

  async function resolveSession(selector?: string) {
    const sessions = await fetchSessions();
    if (sessions.length === 0) {
      throw new Error(describeMissingBridgeSession());
    }
    if (!selector) {
      if (sessions.length === 1) return sessions[0];
      throw new Error(
        'Multiple bridge sessions are connected. Pass a session selector, or call "list_sessions" first.',
      );
    }
    const matches = findMatchingSession(sessions, selector);
    if (matches.length === 1) return matches[0];
    if (matches.length === 0) {
      throw new Error(
        `No bridge session matches "${selector}". Call "list_sessions" to inspect the available documents and session IDs first.`,
      );
    }
    throw new Error(
      `Bridge session selector "${selector}" is ambiguous. Matching sessions: ${matches.map((session) => describeSessionChoice(session.snapshot)).join(", ")}`,
    );
  }

  async function resolveWordSession(selector?: string) {
    const sessions = (await fetchSessions()).filter(
      (session) => session.snapshot.app === "word",
    );
    if (sessions.length === 0) {
      throw new Error(describeMissingBridgeSession());
    }
    if (!selector) {
      if (sessions.length === 1) return sessions[0];
      throw new Error(
        'Multiple Word bridge sessions are connected. Pass a session selector, or call "word_list_documents" first.',
      );
    }
    const matches = findMatchingSession(sessions, selector);
    if (matches.length === 1) return matches[0];
    if (matches.length === 0) {
      throw new Error(
        `No Word bridge session matches "${selector}". Call "word_list_documents" to inspect the available documents first.`,
      );
    }
    throw new Error(
      `Word bridge session selector "${selector}" is ambiguous. Matching sessions: ${matches.map((session) => describeSessionChoice(session.snapshot)).join(", ")}`,
    );
  }

  async function executeWordBridgeTool(
    tool: WordToolContract,
    session: string | undefined,
    args: Record<string, unknown>,
  ) {
    const resolved = await resolveWordSession(session);
    const advertisedTool = resolved.snapshot.tools.find(
      (candidate) => candidate.name === tool.name,
    );
    if (!advertisedTool) {
      throw new Error(
        `The selected Word session does not currently advertise ${tool.name}. Refresh the Word taskpane and retry.`,
      );
    }
    const response = await bridgeRequest<{
      ok: true;
      result: BridgeToolExecutionResult;
    }>(
      "POST",
      `/sessions/${encodeURIComponent(resolved.snapshot.sessionId)}/tools/${encodeURIComponent(tool.name)}`,
      { args },
    );
    return bridgeToolExecutionResultToMcpResult(response.result);
  }

  server.registerTool(
    "get_bridge_status",
    {
      description:
        "Check whether the local Word MCP Bridge server is reachable, even when no Office sessions are connected yet.",
    },
    async () => {
      const status = (await fetchBridgeStatus()) as BridgeServerStatus;
      return buildJsonResult({
        status,
        summary: buildBridgeStatusSummary(status),
      });
    },
  );

  server.registerTool(
    "word_list_documents",
    {
      description:
        "List connected live Word documents with session IDs, titles, saved/unsaved state, and session summaries.",
    },
    async () => {
      const sessions = (await fetchSessions()).filter(
        (session) => session.snapshot.app === "word",
      );
      return buildJsonResult({
        documents: sessions.map(wordDocumentEntry),
      });
    },
  );

  server.registerTool(
    "list_sessions",
    {
      description: "List connected Office bridge sessions.",
    },
    async () => buildJsonResult({ sessions: withDisplayMetadata(await fetchSessions()) }),
  );

  server.registerTool(
    "get_session_snapshot",
    {
      description:
        "Get the latest snapshot for a bridge session, including runtime state, gateway capabilities, and live context.",
      inputSchema: z.object({
        session: z.string().optional(),
      }),
    },
    async ({ session }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "refresh_session",
        },
      );
      return buildJsonResult(response.result);
    },
  );

  server.registerTool(
    "get_live_context",
    {
      description:
        "Read the current live context for a bridge session, including selected Word text when available.",
      inputSchema: z.object({
        session: z.string().optional(),
      }),
    },
    async ({ session }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "refresh_session",
        },
      );
      const snapshot = response.result as {
        gateway?: { liveContext?: unknown };
      };
      return buildJsonResult({
        sessionId: resolved.snapshot.sessionId,
        liveContext: snapshot.gateway?.liveContext ?? null,
      });
    },
  );

  server.registerTool(
    "get_recent_events",
    {
      description: "Fetch recent bridge events for a session.",
      inputSchema: z.object({
        session: z.string().optional(),
        limit: z.number().int().min(1).max(200).optional(),
      }),
    },
    async ({ session, limit }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{
        ok: true;
        events: unknown[];
      }>(
        "GET",
        `/sessions/${encodeURIComponent(resolved.snapshot.sessionId)}/events?limit=${limit ?? 20}`,
      );
      return buildJsonResult({
        sessionId: resolved.snapshot.sessionId,
        events: response.events,
      });
    },
  );

  server.registerTool(
    "call_bridge_tool",
    {
      description:
        "Execute a registered bridge tool for a session. Respect the session's declared capability boundaries.",
      inputSchema: z.object({
        session: z.string().optional(),
        toolName: z.string(),
        args: z.record(z.string(), z.unknown()).optional(),
      }),
    },
    async ({ session, toolName, args }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{
        ok: true;
        result: BridgeToolExecutionResult;
      }>(
        "POST",
        `/sessions/${encodeURIComponent(resolved.snapshot.sessionId)}/tools/${encodeURIComponent(toolName)}`,
        { args: args ?? {} },
      );
      return bridgeToolExecutionResultToMcpResult(response.result);
    },
  );

  server.registerTool(
    "run_unsafe_office_js",
    {
      description:
        "Run privileged Office.js code against a session when that session explicitly exposes the unsafe_office_js capability.",
      inputSchema: z.object({
        session: z.string().optional(),
        code: z.string(),
        explanation: z.string().optional(),
      }),
    },
    async ({ session, code, explanation }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "execute_unsafe_office_js",
          params: { code, explanation },
        },
      );
      return buildJsonResult(response.result);
    },
  );

  server.registerTool(
    "vfs_list",
    {
      description: "List files exposed through the bridge virtual filesystem.",
      inputSchema: z.object({
        session: z.string().optional(),
        prefix: z.string().optional(),
      }),
    },
    async ({ session, prefix }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "vfs_list",
          params: { prefix },
        },
      );
      return buildJsonResult(response.result);
    },
  );

  server.registerTool(
    "vfs_read",
    {
      description: "Read a bridge virtual filesystem file.",
      inputSchema: z.object({
        session: z.string().optional(),
        path: z.string(),
        encoding: z.enum(["text", "base64"]).optional(),
      }),
    },
    async ({ session, path, encoding }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "vfs_read",
          params: { path, encoding },
        },
      );
      return buildJsonResult(response.result);
    },
  );

  server.registerTool(
    "vfs_write",
    {
      description: "Write a bridge virtual filesystem file.",
      inputSchema: z.object({
        session: z.string().optional(),
        path: z.string(),
        text: z.string().optional(),
        dataBase64: z.string().optional(),
      }),
    },
    async ({ session, path, text, dataBase64 }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "vfs_write",
          params: { path, text, dataBase64 },
        },
      );
      return buildJsonResult(response.result);
    },
  );

  server.registerTool(
    "vfs_delete",
    {
      description: "Delete a bridge virtual filesystem file.",
      inputSchema: z.object({
        session: z.string().optional(),
        path: z.string(),
      }),
    },
    async ({ session, path }) => {
      const resolved = await resolveSession(session);
      const response = await bridgeRequest<{ ok: true; result: unknown }>(
        "POST",
        "/rpc",
        {
          sessionId: resolved.snapshot.sessionId,
          method: "vfs_delete",
          params: { path },
        },
      );
      return buildJsonResult(response.result);
    },
  );

  for (const tool of getFirstClassWordToolContracts()) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: z
          .object({
            session: z.string().optional(),
          })
          .and(tool.inputSchema),
      },
      async (input) => {
        const { session, ...args } = input as Record<string, unknown> & {
          session?: string;
        };
        return executeWordBridgeTool(tool, session, args);
      },
    );
  }

  return server;
}
