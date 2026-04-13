import { describe, expect, it } from "vitest";
import { createWordBridgeTools, getWordBridgeToolNames } from "../src/lib/tools";

describe("Word bridge tool registry", () => {
  it("registers a concrete Word tool surface", () => {
    const names = getWordBridgeToolNames();

    expect(names).toContain("word_get_document_text");
    expect(names).toContain("word_get_text_range");
    expect(names).toContain("word_replace_text_range");
    expect(names).toContain("word_insert_paragraph");
    expect(names).toContain("word_format_text_range");
    expect(names).toContain("word_get_revision_scope");
    expect(names).toContain("word_insert_comment");
    expect(names).toContain("word_delete_comment");
    expect(names).toContain("word_reopen_comment");
    expect(names).not.toHaveLength(0);
  });

  it("exposes executable tool definitions with parameter metadata", () => {
    const tools = createWordBridgeTools();
    const readTool = tools.find((tool) => tool.name === "word_get_text_range");
    const writeTool = tools.find((tool) => tool.name === "word_replace_text_range");
    const formatTool = tools.find((tool) => tool.name === "word_format_text_range");

    expect(readTool?.description).toContain("paragraph");
    expect(readTool?.parameters).toMatchObject({
      type: "object",
    });
    expect(writeTool?.requiredCapability).toBe("document_edit");
    expect(formatTool?.requiredCapability).toBe("document_edit");
  });

  it("rejects invalid enum values before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_insert_text");

    const result = await tool?.execute("tool-1", {
      text: "Hello",
      target: "not-a-real-target",
    });

    expect(result).toMatchObject({
      success: false,
    });
  });

  it("rejects empty text inserts before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_insert_text");

    const result = await tool?.execute("tool-empty", {
      text: "",
    });

    expect(result).toMatchObject({
      success: false,
    });
  });

  it("rejects empty formatting mutations before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_format_text_range");

    const result = await tool?.execute("tool-2", {
      paragraphIndex: 0,
      startOffset: 0,
      endOffset: 4,
    });

    expect(result).toMatchObject({
      success: false,
    });
  });

  it("rejects incomplete paragraph-range formatting inputs before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_format_text_range");

    const result = await tool?.execute("tool-3", {
      target: "paragraphRange",
      bold: true,
    });

    expect(result).toMatchObject({
      success: false,
    });
  });

  it("reports unsupported inline formatting fields before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_format_text_range");

    const result = await tool?.execute("tool-unsupported-format", {
      font: "Aptos",
    });

    expect(result).toMatchObject({
      success: false,
      error: expect.stringContaining("Unsupported inline formatting fields"),
    });
  });

  it("rejects unsafe zero line spacing before Office.js execution", async () => {
    const tools = createWordBridgeTools();
    const tool = tools.find((candidate) => candidate.name === "word_set_paragraph_format");

    const result = await tool?.execute("tool-4", {
      paragraphIndex: 0,
      lineSpacing: 0,
    });

    expect(result).toMatchObject({
      success: false,
    });
  });
});
