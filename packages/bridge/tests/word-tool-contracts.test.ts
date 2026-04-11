import { describe, expect, it } from "vitest";
import {
  getFirstClassWordToolContracts,
  getWordToolContract,
  WORD_TOOL_CONTRACTS,
} from "../src/word-tool-contracts";

describe("word tool contracts", () => {
  it("defines unique tool names", () => {
    const names = WORD_TOOL_CONTRACTS.map((tool) => tool.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("marks curated first-class Word wrappers", () => {
    const firstClass = getFirstClassWordToolContracts().map((tool) => tool.name);
    expect(firstClass).toContain("word_get_document_text");
    expect(firstClass).toContain("word_get_text_range");
    expect(firstClass).toContain("word_replace_text_range");
    expect(firstClass).toContain("word_format_text_range");
    expect(firstClass).toContain("word_get_revision_scope");
    expect(firstClass).toContain("word_delete_comment");
    expect(firstClass).toContain("word_reopen_comment");
    expect(firstClass).not.toContain("word_insert_paragraph");
  });

  it("keeps write and review tools on the document_edit capability", () => {
    expect(getWordToolContract("word_insert_text")?.requiredCapability).toBe(
      "document_edit",
    );
    expect(getWordToolContract("word_insert_comment")?.requiredCapability).toBe(
      "document_edit",
    );
    expect(getWordToolContract("word_delete_comment")?.requiredCapability).toBe(
      "document_edit",
    );
  });

  it("classifies search-only tools as read tools", () => {
    expect(getWordToolContract("word_search_text")).toMatchObject({
      group: "read",
      readOnly: true,
      requiredCapability: "tool_call",
    });
  });

  it("validates paragraph-range replace inputs with optional stale checks", () => {
    const schema = getWordToolContract("word_replace_text_range")?.inputSchema;

    expect(
      schema?.safeParse({
        paragraphIndex: 4,
        startOffset: 3,
        endOffset: 8,
        text: "updated",
        expectedText: "prior",
      }).success,
    ).toBe(true);
  });

  it("limits inline formatting mutations to the scoped supported properties", () => {
    const schema = getWordToolContract("word_format_text_range")?.inputSchema;

    expect(schema?.safeParse({ bold: true }).success).toBe(true);
    expect(
      schema?.safeParse({
        target: "paragraphRange",
        paragraphIndex: 1,
        startOffset: 0,
        endOffset: 4,
        italic: true,
      }).success,
    ).toBe(true);
    expect(
      schema?.safeParse({
        target: "paragraphRange",
        italic: true,
      }).success,
    ).toBe(false);
    expect(schema?.safeParse({}).success).toBe(false);
  });

  it("describes targetMatchIndexes as integer array items in search-and-replace metadata", () => {
    const contract = getWordToolContract("word_search_and_replace");
    const targetMatchIndexes = contract?.parameters.properties?.targetMatchIndexes as
      | Record<string, unknown>
      | undefined;

    expect(targetMatchIndexes).toMatchObject({
      type: "array",
      items: {
        type: "integer",
        minimum: 0,
      },
    });
  });
});
