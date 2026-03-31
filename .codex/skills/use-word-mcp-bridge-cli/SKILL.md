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

## Quick reference

| Task | Command |
|------|---------|
| Sessions | `office-bridge list` |
| Short status line | `office-bridge summary word` |
| Structured snapshot | `office-bridge snapshot word` (+ `--compact`, `--fields`) |
| Recent events | `office-bridge events word --limit N` |
| Live selection stream | `office-bridge watch-selection word` |
| Bridge/status | `office-bridge status` |
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

## Scope

This repo’s add-in is Word-focused; the CLI may list other app names if present—confirm with **`list`**. For every subcommand, **`packages/bridge/src/cli.ts`** is the source of truth.

## Copy out of repo

Copy `.codex/skills/use-word-mcp-bridge-cli/` into a global Codex skills directory if desired.
