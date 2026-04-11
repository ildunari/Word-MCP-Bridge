import { describe, expect, it } from "vitest";
import type { BridgeSessionSnapshot } from "../src/protocol";
import {
  describeSessionChoice,
  getSessionDocumentLabel,
  getSessionDocumentSummary,
} from "../src/session-labels";

function makeSnapshot(
  overrides: Partial<BridgeSessionSnapshot> = {},
): BridgeSessionSnapshot {
  const now = Date.now();
  return {
    sessionId: "word:demo",
    instanceId: "instance-1",
    app: "word",
    appName: "Microsoft Word",
    documentId: "doc-1",
    documentMetadata: {},
    tools: [],
    host: { href: "https://localhost:3014/taskpane.html" },
    connectedAt: now,
    updatedAt: now,
    ...overrides,
  };
}

describe("session labels", () => {
  it("prefers a real document title", () => {
    const snapshot = makeSnapshot({
      documentId: "doc-1",
      documentMetadata: { title: "Grant Draft.docx" },
    });

    expect(getSessionDocumentLabel(snapshot)).toBe("Grant Draft.docx");
    expect(getSessionDocumentSummary(snapshot)).toBe("Document ID doc-1");
  });

  it("maps local unsaved Word documents to a human-friendly label", () => {
    const snapshot = makeSnapshot({
      documentId: "word-local:abc123",
      documentMetadata: { title: "Word MCP Bridge" },
    });

    expect(getSessionDocumentLabel(snapshot)).toBe("Untitled Word document");
    expect(getSessionDocumentSummary(snapshot)).toBe("Unsaved local document");
  });

  it("falls back to a file basename when the document id looks like a path", () => {
    const snapshot = makeSnapshot({
      app: "excel",
      appName: "Excel",
      sessionId: "excel:budget",
      documentId: "/Users/kosta/Documents/Budget.xlsx",
      documentMetadata: {},
    });

    expect(getSessionDocumentLabel(snapshot)).toBe("Budget.xlsx");
    expect(getSessionDocumentSummary(snapshot)).toBe("Saved document");
  });

  it("builds an ambiguity-friendly selector hint", () => {
    const snapshot = makeSnapshot({
      sessionId: "word:demo-2",
      documentId: "doc-2",
      documentMetadata: { title: "Paper Draft.docx" },
    });

    expect(describeSessionChoice(snapshot)).toBe(
      "Paper Draft.docx [word:demo-2; Document ID doc-2]",
    );
  });
});
