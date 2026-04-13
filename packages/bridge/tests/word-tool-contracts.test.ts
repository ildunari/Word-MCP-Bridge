import { describe, expect, it } from "vitest";
import {
  getFirstClassWordToolContracts,
  getWordToolContract,
  normalizeWordToolArgs,
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
    expect(firstClass).toContain("word_insert_paragraph");
    expect(firstClass).toContain("word_replace_text_range");
    expect(firstClass).toContain("word_format_text_range");
    expect(firstClass).toContain("word_get_revision_scope");
    expect(firstClass).toContain("word_delete_comment");
    expect(firstClass).toContain("word_reopen_comment");
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

  it("accepts the additive cursor insert target while rejecting empty text", () => {
    const schema = getWordToolContract("word_insert_text")?.inputSchema;

    expect(
      schema?.safeParse({
        text: "Hello",
        target: "cursor",
      }).success,
    ).toBe(true);
    expect(
      schema?.safeParse({
        text: "",
      }).success,
    ).toBe(false);
  });

  it("normalizes intuitive compatibility aliases to canonical Word tool inputs", () => {
    expect(
      normalizeWordToolArgs("word_search_text", {
        searchText: "alpha",
      }),
    ).toMatchObject({
      query: "alpha",
    });
    expect(
      normalizeWordToolArgs("word_search_and_replace", {
        searchText: "alpha",
        replaceText: "beta",
      }),
    ).toMatchObject({
      query: "alpha",
      replacement: "beta",
    });
    expect(
      normalizeWordToolArgs("word_replace_text_range", {
        replacement: "delta",
      }),
    ).toMatchObject({
      text: "delta",
    });
    expect(
      normalizeWordToolArgs("word_apply_style", {
        styleName: "Heading 1",
      }),
    ).toMatchObject({
      style: "Heading 1",
    });
    expect(
      normalizeWordToolArgs("word_reply_to_comment", {
        commentId: 1110707823,
      }),
    ).toMatchObject({
      commentId: "1110707823",
    });
    expect(
      normalizeWordToolArgs("word_format_text_range", {
        color: "#336699",
      }),
    ).toMatchObject({
      fontColor: "#336699",
    });
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

  it("rejects unsafe zero line spacing and documents point-based spacing values", () => {
    const contract = getWordToolContract("word_set_paragraph_format");
    const lineSpacing = contract?.parameters.properties?.lineSpacing as
      | Record<string, unknown>
      | undefined;

    expect(
      contract?.inputSchema.safeParse({
        paragraphIndex: 0,
        lineSpacing: 0,
      }).success,
    ).toBe(false);
    expect(
      contract?.inputSchema.safeParse({
        paragraphIndex: 0,
        lineSpacing: 12,
      }).success,
    ).toBe(true);
    expect(lineSpacing).toMatchObject({
      type: "number",
      exclusiveMinimum: 0,
      maximum: 144,
    });
    expect(String(lineSpacing?.description ?? "")).toContain("12");
    expect(String(lineSpacing?.description ?? "")).toContain("points");
  });
});
