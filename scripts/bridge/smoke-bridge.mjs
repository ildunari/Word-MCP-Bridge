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
    json: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--") continue;
    if (arg === "--bridge-url") options.bridgeUrl = argv[++index];
    else if (arg === "--session") options.session = argv[++index];
    else if (arg === "--expect-hidden") options.expectHidden = true;
    else if (arg === "--launch") options.launch = true;
    else if (arg === "--write") options.write = true;
    else if (arg === "--json") options.json = true;
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

function createMcpClient({ bridgeUrl }) {
  const transport = new StdioClientTransport({
    command: "pnpm",
    args: ["exec", "office-bridge", "mcp-serve", "--url", bridgeUrl],
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

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    process.stdout.write(
      "Usage: node scripts/bridge/smoke-bridge.mjs [--bridge-url URL] [--session SESSION_ID] [--expect-hidden] [--launch] [--write] [--json]\n",
    );
    return;
  }

  const cwd = process.cwd();
  const results = {
    bridgeUrl: options.bridgeUrl,
    launcherRecovery: null,
    hiddenSharedRuntime: null,
    listSessions: null,
    tableReadUpdate: null,
    invalidParagraphSpacing: null,
  };

  if (options.launch) {
    await runCommand("bash", ["scripts/bridge/launch-word-taskpane.sh", "--mode", "open", "--timeout", "20"], { cwd });
    results.launcherRecovery = { success: true };
  }

  const client = createMcpClient({ bridgeUrl: options.bridgeUrl });
  try {
    await client.initialize();

    const listSessionsResult = await client.callTool("list_sessions");
    const listSessionsText = extractTextContent(listSessionsResult);
    const listSessionsParsed = JSON.parse(listSessionsText);
    const sessions = listSessionsParsed.sessions ?? [];
    const selectedSession = getSelectedSession(sessions, options.session);
    const sessionId = selectedSession.snapshot.sessionId;

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
