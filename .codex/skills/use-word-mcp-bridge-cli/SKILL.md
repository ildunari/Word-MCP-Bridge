---
name: use-word-mcp-bridge-cli
description: Use for Word-MCP-Bridge operations with the office-bridge CLI or stdio MCP—bridge up/down, session discovery, summary/snapshot/events/watch-selection/status, serve vs mcp-serve, repo scripts (pnpm bridge:serve, pnpm exec office-bridge), and fixes for empty sessions or URL/token issues. Triggers include office-bridge, mcp-serve, Word MCP Bridge, localhost 4017, bridge CLI troubleshooting.
---

# Word MCP Bridge — `office-bridge` CLI (Codex)

Use this skill when the user is driving **Word-MCP-Bridge** from the terminal or wiring MCP. Ground truth: `packages/bridge/src/cli.ts` (printed usage), `packages/bridge/README.md`, repo root `README.md`.

## Model behavior

- Prefer **real commands** from this repo; do not invent flags.
- State clearly that **two processes** are required for MCP: **`office-bridge serve`** (bridge the add-in connects to) and **`office-bridge mcp-serve`** (stdio MCP adapter). `mcp-serve` does not replace `serve`.
- Default bridge URL is **`https://localhost:4017`** unless the user overrides `--url`.
- Treat `office-bridge list` and `office-bridge status` as the first liveness check. With multiple Word docs open, do not assume `word` is a safe selector; prefer the explicit session id or the document title shown by `list`.
- If a replace call fails because `targetMatchIndexes` is not available in the current candidate set, that is the intended safe refusal path, not a bridge failure.

## Quick reference

| Task | Command |
|------|---------|
| Sessions | `office-bridge list` |
| Word-only session inventory | `office-bridge status` or `office-bridge list` |
| Short status line | `office-bridge summary word` |
| Structured snapshot | `office-bridge snapshot word` (+ `--compact`, `--fields`) |
| Recent events | `office-bridge events word --limit N` |
| Live selection stream | `office-bridge watch-selection word` |
| Bridge/status | `office-bridge status` |
| Focus or reopen the taskpane surface | `scripts/bridge/launch-word-taskpane.sh --mode open` |
| Start bridge | `office-bridge serve` |
| MCP stdio | `office-bridge mcp-serve [--url https://localhost:4017]` |

Repo shortcuts (`package.json`): **`pnpm bridge:serve`**, **`pnpm bridge:mcp`**, **`pnpm exec office-bridge <subcommand>`**. Full install flow: root **`README.md`**.

## MCP host snippets

Use the JSON/TOML/CLI examples in **`packages/bridge/README.md`** for Claude Desktop, Cursor, Claude Code, and Codex (`codex mcp add`, `config.toml`).

## Troubleshooting checklist

1. Is **`serve`** running?
2. Is the **Word taskpane** open and connected?
3. Same **`--url`** for CLI and `mcp-serve`?
4. If auth issues: **`OFFICE_BRIDGE_TOKEN`** per **`packages/bridge/README.md`**.
5. If Word was reopened and `list` is empty, the add-in install is still persistent; the missing piece is usually the taskpane or shared runtime not being reattached yet, so reopen `Word MCP Bridge` from Word's Add-ins UI or use `scripts/bridge/launch-word-taskpane.sh --mode open`.
6. If `list` shows two docs but one looks stale, compare `lastSeen`/`pendingCount` from `status` before treating it as a tool bug.

## Current live behavior notes

- The bridge now exposes a **31-tool Word surface** including precise range reads/writes, scoped formatting, revision scope reads, and fuller comment lifecycle tools.
- Multi-document sessions are supported. Use the exact `session=` value from `office-bridge list` for destructive checks instead of relying on the generic `word` selector.
- Shared-runtime mode can keep the bridge alive even after the panel is hidden. A hidden session is healthy if `status` still shows a live Word session.
- `word_search_and_replace` intentionally fails closed when `targetMatchIndexes` refers to a truncated or unavailable candidate. The safe flow is:
  1. `word_search_text`
  2. inspect returned `matchIndex` values
  3. rerun `word_search_and_replace` with a matching `maxMatches`

## Scope

This repo’s add-in is Word-focused; the CLI may list other app names if present—confirm with **`list`**. For every subcommand, **`packages/bridge/src/cli.ts`** is the source of truth.

## Copy out of repo

Copy `.codex/skills/use-word-mcp-bridge-cli/` into a global Codex skills directory if desired.
