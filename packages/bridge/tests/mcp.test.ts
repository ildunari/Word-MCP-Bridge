import { describe, expect, it } from "vitest";
import {
  buildBridgeStatusSummary,
  bridgeToolExecutionResultToMcpResult,
  describeBridgeConnectionFailure,
  describeMissingBridgeSession,
} from "../src/mcp";
import { describeSessionChoice } from "../src/session-labels";

describe("bridgeToolExecutionResultToMcpResult", () => {
  it("preserves text and images for MCP tool responses", () => {
    const result = bridgeToolExecutionResultToMcpResult({
      toolCallId: "tool-1",
      toolName: "screenshot_document",
      isError: false,
      result: { ok: true },
      resultText: "Captured page 1",
      images: [{ data: "abc123", mimeType: "image/png" }],
    });

    expect(result.isError).toBe(false);
    expect(result.content).toEqual([
      { type: "text", text: "Captured page 1" },
      { type: "image", data: "abc123", mimeType: "image/png" },
    ]);
  });

  it("falls back to serialized structured content when text is empty", () => {
    const result = bridgeToolExecutionResultToMcpResult({
      toolCallId: "tool-2",
      toolName: "get_document_text",
      isError: false,
      result: { body: "hello" },
      resultText: "",
      images: [],
    });

    expect(result.content).toEqual([
      {
        type: "text",
        text: JSON.stringify({ body: "hello" }, null, 2),
      },
    ]);
  });

  it("prefers nested structuredContent returned by Word tools", () => {
    const result = bridgeToolExecutionResultToMcpResult({
      toolCallId: "tool-3",
      toolName: "word_get_document_text",
      isError: false,
      result: {
        content: [{ type: "text", text: "Read 2 paragraphs." }],
        structuredContent: {
          paragraphCount: 2,
          paragraphs: [{ paragraphIndex: 0, text: "Hello" }],
        },
      },
      resultText: "Read 2 paragraphs.",
      images: [],
    });

    expect(result.structuredContent).toEqual({
      paragraphCount: 2,
      paragraphs: [{ paragraphIndex: 0, text: "Hello" }],
    });
  });

  it("falls back to the inner tool payload for error structuredContent", () => {
    const result = bridgeToolExecutionResultToMcpResult({
      toolCallId: "tool-4",
      toolName: "word_replace_selection",
      isError: true,
      result: {
        success: false,
        error: "There is no active text selection to replace.",
      },
      resultText: "",
      images: [],
      error: "There is no active text selection to replace.",
    });

    expect(result.isError).toBe(true);
    expect(result.structuredContent).toEqual({
      success: false,
      error: "There is no active text selection to replace.",
    });
  });

  it("surfaces refreshed scope payloads from deterministic range writes", () => {
    const result = bridgeToolExecutionResultToMcpResult({
      toolCallId: "tool-5",
      toolName: "word_replace_text_range",
      isError: false,
      result: {
        content: [{ type: "text", text: "Replaced 6 characters in paragraph 4." }],
        structuredContent: {
          success: true,
          paragraphIndex: 4,
          appliedRange: { startOffset: 10, endOffset: 16 },
          paragraph: {
            beforeText: "Alpha beta gamma",
            text: "Alpha delta gamma",
          },
          verification: {
            exactMatch: true,
            adjacentParagraphsUnchanged: true,
          },
        },
      },
      resultText: "Replaced 6 characters in paragraph 4.",
      images: [],
    });

    expect(result.structuredContent).toEqual({
      success: true,
      paragraphIndex: 4,
      appliedRange: { startOffset: 10, endOffset: 16 },
      paragraph: {
        beforeText: "Alpha beta gamma",
        text: "Alpha delta gamma",
      },
      verification: {
        exactMatch: true,
        adjacentParagraphsUnchanged: true,
      },
    });
  });

  it("describes bridge connection failures with an actionable message", () => {
    const error = Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:4017"), {
      code: "ECONNREFUSED",
    });

    expect(describeBridgeConnectionFailure(error, "https://localhost:4017")).toContain(
      "Start the local bridge server",
    );
  });

  it("describes the zero-session case with next steps", () => {
    expect(describeMissingBridgeSession()).toContain(
      "Open Word and the Word MCP Bridge taskpane",
    );
  });

  it("builds human-readable ambiguity hints for session choices", () => {
    expect(
      describeSessionChoice({
        sessionId: "word:doc-2",
        instanceId: "inst-2",
        app: "word",
        appName: "Microsoft Word",
        documentId: "doc-2",
        documentMetadata: { title: "Draft.docx" },
        tools: [],
        host: { href: "https://localhost:3014/taskpane.html" },
        connectedAt: 1,
        updatedAt: 1,
      }),
    ).toBe("Draft.docx [word:doc-2; Document ID doc-2]");
  });

  it("builds a compact bridge-status summary for verifier lanes", () => {
    const summary = buildBridgeStatusSummary({
      startedAt: 10,
      uptimeMs: 2000,
      host: "127.0.0.1",
      port: 4017,
      httpUrl: "https://localhost:4017",
      wsUrl: "wss://localhost:4017/ws",
      sessionCount: 1,
      totals: {
        connectedSessionCount: 1,
        disconnectedSessionCount: 0,
        pendingCount: 2,
        toolCalls: 7,
        toolErrors: 1,
        eventsReceived: 12,
        connectionDropCount: 0,
      },
      sessions: [
        {
          snapshot: {
            sessionId: "word:doc-2",
            instanceId: "inst-2",
            app: "word",
            appName: "Microsoft Word",
            documentId: "doc-2",
            documentMetadata: { title: "Draft.docx" },
            tools: [{ name: "word_get_document_text" }],
            host: { href: "https://localhost:3014/taskpane.html" },
            connectedAt: 1,
            updatedAt: 5,
          },
          connectedAt: 1,
          lastSeenAt: 6,
          recentEvents: [],
          pendingCount: 2,
          metrics: {
            toolCalls: 7,
            toolErrors: 1,
            eventsReceived: 12,
            connectionDropCount: 0,
          },
          health: "live",
        },
      ],
    });

    expect(summary).toEqual({
      startedAt: 10,
      uptimeMs: 2000,
      host: "127.0.0.1",
      port: 4017,
      httpUrl: "https://localhost:4017",
      wsUrl: "wss://localhost:4017/ws",
      sessionCount: 1,
      totals: {
        connectedSessionCount: 1,
        disconnectedSessionCount: 0,
        pendingCount: 2,
        toolCalls: 7,
        toolErrors: 1,
        eventsReceived: 12,
        connectionDropCount: 0,
      },
      sessions: [
        {
          sessionId: "word:doc-2",
          instanceId: "inst-2",
          app: "word",
          title: "Draft.docx",
          summary: "Document ID doc-2",
          documentId: "doc-2",
          health: "live",
          pendingCount: 2,
          toolCount: 1,
          connectedAt: 1,
          lastSeenAt: 6,
          updatedAt: 5,
        },
      ],
    });
  });
});
