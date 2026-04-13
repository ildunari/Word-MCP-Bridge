import { execFile as execFileCallback } from "node:child_process";
import { promisify } from "node:util";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const execFile = promisify(execFileCallback);

function parseArgs(argv) {
  const options = {
    bridgeUrl: "https://localhost:4017",
    session: undefined,
    expectHidden: false,
    launch: false,
    write: false,
    schemaOnly: false,
    json: false,
    mcpCommand: undefined,
    mcpArgs: [],
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--") continue;
    if (arg === "--bridge-url") options.bridgeUrl = argv[++index];
    else if (arg === "--session") options.session = argv[++index];
    else if (arg === "--expect-hidden") options.expectHidden = true;
    else if (arg === "--launch") options.launch = true;
    else if (arg === "--write") options.write = true;
    else if (arg === "--schema-only") options.schemaOnly = true;
    else if (arg === "--json") options.json = true;
    else if (arg === "--mcp-command") options.mcpCommand = argv[++index];
    else if (arg === "--mcp-arg") options.mcpArgs.push(argv[++index]);
    else if (arg === "--help" || arg === "-h") options.help = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }

  return options;
}

async function runCommand(command, args, { cwd } = {}) {
  const result = await execFile(command, args, {
    cwd,
    maxBuffer: 1024 * 1024 * 8,
    env: process.env,
  });
  return (result.stdout ?? "").trim();
}

function createMcpClient({ bridgeUrl, mcpCommand, mcpArgs }) {
  const transport = new StdioClientTransport({
    command: mcpCommand ?? "pnpm",
    args:
      mcpArgs && mcpArgs.length > 0
        ? mcpArgs
        : ["exec", "office-bridge", "mcp-serve", "--url", bridgeUrl],
    stderr: "pipe",
    env: {
      ...process.env,
    },
    cwd: process.cwd(),
  });
  const client = new Client({
    name: "word-bridge-smoke",
    version: "1.0.0",
  });
  let stderrText = "";

  transport.stderr?.on("data", (chunk) => {
    stderrText += chunk.toString("utf8");
  });

  client.onerror = () => {
    // Tool calls return structured failures; stderr is captured for extra context.
  };

  return {
    async initialize() {
      await client.connect(transport);
    },
    async listTools() {
      try {
        return await client.listTools();
      } catch (error) {
        const stderrSuffix = stderrText.trim() ? `\n${stderrText.trim()}` : "";
        throw new Error(
          `${error instanceof Error ? error.message : String(error)}${stderrSuffix}`,
        );
      }
    },
    async callTool(name, argumentsValue = {}) {
      try {
        return await client.callTool({
          name,
          arguments: argumentsValue,
        });
      } catch (error) {
        const stderrSuffix = stderrText.trim() ? `\n${stderrText.trim()}` : "";
        throw new Error(
          `${error instanceof Error ? error.message : String(error)}${stderrSuffix}`,
        );
      }
    },
    async close() {
      await transport.close();
    },
  };
}

function extractTextContent(result) {
  return (result?.content ?? [])
    .filter((item) => item?.type === "text")
    .map((item) => item.text)
    .join("\n");
}

function parseJsonTextResult(result) {
  const text = extractTextContent(result);
  if (!text) {
    throw new Error("MCP tool returned no text payload.");
  }
  return JSON.parse(text);
}

function getSelectedSession(sessions, requestedSession) {
  if (requestedSession) {
    const exact = sessions.find((session) => session.snapshot?.sessionId === requestedSession);
    if (!exact) {
      throw new Error(`No MCP session matched ${requestedSession}.`);
    }
    return exact;
  }
  if (sessions.length !== 1) {
    throw new Error(
      `Expected exactly one session for smoke validation, found ${sessions.length}. Pass --session to disambiguate.`,
    );
  }
  return sessions[0];
}

function findTool(tools, name) {
  const tool = tools.find((candidate) => candidate.name === name);
  if (!tool) {
    throw new Error(`Expected tools/list to include ${name}.`);
  }
  return tool;
}

function assertWordToolSchemas(tools) {
  const searchTool = findTool(tools, "word_search_text");
  const formatTool = findTool(tools, "word_format_text_range");
  const insertParagraphTool = findTool(tools, "word_insert_paragraph");
  const bridgeStatusTool = findTool(tools, "get_bridge_status");

  const searchProperties = searchTool.inputSchema?.properties ?? {};
  const formatProperties = formatTool.inputSchema?.properties ?? {};
  const insertParagraphProperties = insertParagraphTool.inputSchema?.properties ?? {};
  const bridgeStatusProperties = bridgeStatusTool.inputSchema?.properties ?? {};

  for (const [toolName, requiredProperty] of [
    ["word_search_text", "query"],
    ["word_format_text_range", "fontColor"],
    ["word_insert_paragraph", "text"],
    ["get_bridge_status", "verbose"],
  ]) {
    const properties =
      toolName === "word_search_text"
        ? searchProperties
        : toolName === "word_format_text_range"
          ? formatProperties
          : toolName === "word_insert_paragraph"
            ? insertParagraphProperties
            : bridgeStatusProperties;
    if (!properties[requiredProperty]) {
      throw new Error(
        `tools/list omitted ${requiredProperty} from ${toolName}. The direct MCP schema is still incomplete.`,
      );
    }
  }

  if (!Array.isArray(searchTool.inputSchema?.required) || !searchTool.inputSchema.required.includes("query")) {
    throw new Error("tools/list did not mark word_search_text.query as required.");
  }

  return {
    toolCount: tools.length,
    verifiedTools: [
      "word_search_text",
      "word_format_text_range",
      "word_insert_paragraph",
      "get_bridge_status",
    ],
  };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    process.stdout.write(
      "Usage: node scripts/bridge/smoke-bridge.mjs [--bridge-url URL] [--session SESSION_ID] [--expect-hidden] [--launch] [--write] [--schema-only] [--mcp-command CMD] [--mcp-arg ARG] [--json]\n",
    );
    return;
  }

  const cwd = process.cwd();
  const results = {
    bridgeUrl: options.bridgeUrl,
    launcherRecovery: null,
    hiddenSharedRuntime: null,
    directSchemas: null,
    directWordCalls: null,
    listSessions: null,
    tableReadUpdate: null,
    invalidParagraphSpacing: null,
  };

  if (options.launch) {
    await runCommand("bash", ["scripts/bridge/launch-word-taskpane.sh", "--mode", "open", "--timeout", "20"], { cwd });
    results.launcherRecovery = { success: true };
  }

  const client = createMcpClient({
    bridgeUrl: options.bridgeUrl,
    mcpCommand: options.mcpCommand,
    mcpArgs: options.mcpArgs,
  });
  try {
    await client.initialize();
    const tools = (await client.listTools())?.tools;
    if (!tools) {
      throw new Error("Unable to read tools/list from the MCP server.");
    }
    results.directSchemas = {
      success: true,
      ...assertWordToolSchemas(tools),
    };

    if (options.schemaOnly) {
      if (options.json) {
        process.stdout.write(`${JSON.stringify(results, null, 2)}\n`);
        return;
      }
      process.stdout.write("Word bridge schema smoke checks passed.\n");
      process.stdout.write(
        `- direct schemas: verified ${results.directSchemas.verifiedTools.length} tool definition(s)\n`,
      );
      return;
    }

    const listSessionsResult = await client.callTool("list_sessions");
    const listSessionsText = extractTextContent(listSessionsResult);
    const listSessionsParsed = JSON.parse(listSessionsText);
    const sessions = listSessionsParsed.sessions ?? [];
    const selectedSession = getSelectedSession(sessions, options.session);
    const sessionId = selectedSession.snapshot.sessionId;

    const revisionsDirectResult = await client.callTool("word_get_revisions", {
      session: sessionId,
    });
    const revisionsDirectStructured =
      revisionsDirectResult.structuredContent ?? parseJsonTextResult(revisionsDirectResult);
    const revisionsFallbackResult = await client.callTool("call_bridge_tool", {
      session: sessionId,
      toolName: "word_get_revisions",
    });
    const revisionsFallbackStructured =
      revisionsFallbackResult.structuredContent ?? parseJsonTextResult(revisionsFallbackResult);
    const documentTextResult = await client.callTool("word_get_document_text", {
      session: sessionId,
      startParagraph: 0,
      endParagraph: 4,
    });
    const documentTextStructured =
      documentTextResult.structuredContent ?? parseJsonTextResult(documentTextResult);
    const paragraphText = (documentTextStructured.paragraphs ?? [])
      .map((paragraph) => paragraph.text ?? "")
      .find((text) => typeof text === "string" && text.trim().length > 0);
    const derivedQuery = paragraphText
      ? paragraphText
          .split(/\s+/)
          .map((token) => token.replace(/[^\p{L}\p{N}_-]+/gu, ""))
          .find((token) => token.length >= 4)
      : undefined;
    if (!derivedQuery) {
      throw new Error("Could not derive a safe direct search query from the active document.");
    }
    const searchDirectResult = await client.callTool("word_search_text", {
      session: sessionId,
      query: derivedQuery,
      maxMatches: 5,
    });
    const searchDirectStructured =
      searchDirectResult.structuredContent ?? parseJsonTextResult(searchDirectResult);
    const unsupportedFormatResult = await client.callTool("word_format_text_range", {
      session: sessionId,
      font: "Aptos",
    });
    const unsupportedFormatStructured =
      unsupportedFormatResult.structuredContent ?? parseJsonTextResult(unsupportedFormatResult);
    if (unsupportedFormatStructured.success !== false) {
      throw new Error("word_format_text_range did not reject unsupported inline formatting fields.");
    }
    results.directWordCalls = {
      success: true,
      revisionsDirectCount: revisionsDirectStructured.revisions?.length ?? 0,
      revisionsFallbackCount: revisionsFallbackStructured.revisions?.length ?? 0,
      derivedQuery,
      searchMatchCount: searchDirectStructured.totalMatches ?? 0,
      unsupportedFormatError: unsupportedFormatStructured.error ?? null,
    };
    if (
      results.directWordCalls.revisionsDirectCount !==
      results.directWordCalls.revisionsFallbackCount
    ) {
      throw new Error(
        "Direct word_get_revisions and call_bridge_tool returned different revision counts.",
      );
    }

    const hasRecentEvents =
      Object.prototype.hasOwnProperty.call(selectedSession, "recentEvents") ||
      Object.prototype.hasOwnProperty.call(selectedSession.snapshot ?? {}, "tools");
    const hasToolNames = Array.isArray(selectedSession.snapshot?.toolNames);

    if (hasRecentEvents || !hasToolNames) {
      throw new Error(
        "list_sessions returned a verbose session payload. Expected toolNames-only compact output with no recentEvents/tools definitions.",
      );
    }

    results.listSessions = {
      success: true,
      bytes: Buffer.byteLength(listSessionsText),
      sessionCount: sessions.length,
      toolCount: selectedSession.snapshot.toolCount,
      toolNamesCount: selectedSession.snapshot.toolNames.length,
    };

    const bridgeStatusResult = await client.callTool("get_bridge_status");
    const bridgeStatus = parseJsonTextResult(bridgeStatusResult).summary;
    const matchingSummarySession = (bridgeStatus.sessions ?? []).find(
      (session) => session.sessionId === sessionId,
    );
    const hiddenActive = Boolean(
      selectedSession.snapshot?.documentMetadata?.hiddenActive ||
        selectedSession.snapshot?.runtimeState?.paneVisibility === "hidden" ||
        matchingSummarySession?.visibilityMode === "hidden",
    );
    if (options.expectHidden && !hiddenActive) {
      throw new Error("Expected a hidden shared-runtime session, but the current session is visible.");
    }
    results.hiddenSharedRuntime = {
      success: true,
      hiddenActive,
      sessionId,
    };

    const tablesResult = await client.callTool("call_bridge_tool", {
      session: sessionId,
      toolName: "word_get_tables",
      args: { includeCellPreview: true },
    });
    const tablesStructured = tablesResult.structuredContent ?? parseJsonTextResult(tablesResult);
    const tables = tablesStructured.tables ?? [];
    if (tables.some((table) => !table.columnCount || table.columnCount <= 0)) {
      throw new Error("One or more tables reported an invalid zero column count.");
    }

    let tableSummary;
    if (tables.length === 0) {
      if (options.write) {
        throw new Error(
          "Smoke validation requires at least one table in the active document when --write is enabled.",
        );
      }
      tableSummary = {
        success: true,
        tableCount: 0,
        skipped: true,
        reason: "no_tables_in_active_document",
      };
    } else {
      let updateSummary = { success: true, skipped: !options.write };
      if (options.write) {
        const tableToPatch = tables.find(
          (table) =>
            Array.isArray(table.preview) &&
            table.preview.length > 1 &&
            Array.isArray(table.preview[1]) &&
            table.preview[1].length > 1 &&
            typeof table.preview[1][1] === "string",
        );
        if (!tableToPatch) {
          throw new Error(
            "Smoke validation requires a previewable table cell at row 1 column 1 so update_table_cell can be exercised safely.",
          );
        }
        const existingText = tableToPatch.preview[1][1];
        const updateResult = await client.callTool("call_bridge_tool", {
          session: sessionId,
          toolName: "word_update_table_cell",
          args: {
            tableIndex: tableToPatch.tableIndex,
            rowIndex: 1,
            columnIndex: 1,
            text: existingText,
          },
        });
        const updateStructured =
          updateResult.structuredContent ?? parseJsonTextResult(updateResult);
        if (updateStructured.success !== true) {
          throw new Error("word_update_table_cell did not report success.");
        }
        updateSummary = {
          success: true,
          skipped: false,
          tableIndex: tableToPatch.tableIndex,
          rowIndex: 1,
          columnIndex: 1,
        };
      }
      tableSummary = {
        success: true,
        tableCount: tables.length,
        skipped: false,
        update: updateSummary,
      };
    }
    results.tableReadUpdate = tableSummary;

    const invalidSpacingResult = await client.callTool("call_bridge_tool", {
      session: sessionId,
      toolName: "word_set_paragraph_format",
      args: {
        paragraphIndex: 0,
        lineSpacing: 0,
      },
    });
    const invalidSpacingStructured =
      invalidSpacingResult.structuredContent ?? parseJsonTextResult(invalidSpacingResult);
    if (invalidSpacingStructured.success !== false) {
      throw new Error("Unsafe zero line spacing was not rejected.");
    }
    results.invalidParagraphSpacing = {
      success: true,
      error: invalidSpacingStructured.error ?? null,
    };
  } finally {
    await client.close();
  }

  if (options.json) {
    process.stdout.write(`${JSON.stringify(results, null, 2)}\n`);
    return;
  }

  process.stdout.write("Word bridge smoke checks passed.\n");
  process.stdout.write(`- launcher recovery: ${results.launcherRecovery?.success ? "ok" : "skipped"}\n`);
  process.stdout.write(
    `- hidden shared runtime: ${results.hiddenSharedRuntime?.hiddenActive ? "hidden-active" : "visible-but-live"}\n`,
  );
  process.stdout.write(
    `- direct schemas: verified ${results.directSchemas?.verifiedTools?.length ?? 0} tool definition(s)\n`,
  );
  process.stdout.write(
    `- direct Word calls: revisions ${results.directWordCalls?.revisionsDirectCount ?? 0}, search matches ${results.directWordCalls?.searchMatchCount ?? 0}\n`,
  );
  process.stdout.write(
    `- compact list_sessions: ${results.listSessions?.bytes ?? 0} bytes for ${results.listSessions?.sessionCount ?? 0} session(s)\n`,
  );
  process.stdout.write(
    `- table read/update: ${results.tableReadUpdate?.tableCount ?? 0} table(s), write ${
      results.tableReadUpdate?.update?.skipped ? "skipped" : "ok"
    }\n`,
  );
  process.stdout.write(
    `- invalid paragraph spacing: rejected with ${results.invalidParagraphSpacing?.error ?? "an error"}\n`,
  );
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exit(1);
});
