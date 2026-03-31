---
name: use-word-mcp-bridge-cli
description: Use when operating Word-MCP-Bridge via the office-bridge CLI—live Word taskpane sessions, bridge health, MCP stdio wiring, session inspection (list, summary, snapshot, events, watch-selection, status), serve vs mcp-serve, or troubleshooting localhost bridge connectivity. Triggers on office-bridge, Word MCP Bridge, localhost 4017, bridge serve, mcp-serve, watch-selection, bridge CLI.
---

# Word MCP Bridge — `office-bridge` CLI

Operate the **Word MCP Bridge** stack after the bridge package and Word add-in are installed. Authoritative command text lives in `packages/bridge/src/cli.ts` (usage block); operator docs in `packages/bridge/README.md` and repo `README.md`.

## When to use the CLI vs MCP

- **CLI (`office-bridge …`)**: ad-hoc debugging, scripts, CI-style checks, quick reads of session state, tailing events or selection, confirming the bridge is up.
- **`office-bridge mcp-serve`**: stdio MCP server for Claude Desktop, Cursor, Claude Code, Codex, etc. It **wraps** a running bridge; it does **not** replace `serve` (`packages/bridge/README.md`).

## Two processes (do not skip)

1. **`office-bridge serve`** — local HTTPS/WebSocket bridge the Word taskpane connects to (default `https://localhost:4017`).
2. **`office-bridge mcp-serve`** — MCP host launches this; it talks to the bridge URL (often `--url https://localhost:4017`).

From the repo: `pnpm bridge:serve` starts the server; `pnpm bridge:mcp` runs `pnpm exec office-bridge mcp-serve` (`package.json`). Hosted-add-in users still run `serve` locally per `README.md`.

## Session argument

Word sessions are typically selected with the **`word`** positional (see examples in `packages/bridge/src/cli.ts`). Use `office-bridge list` if unsure.

## Command selection

| Goal | Command |
|------|---------|
| See if anything is connected | `office-bridge list` |
| One-line human status | `office-bridge summary word` |
| Full session snapshot (fields, compact) | `office-bridge snapshot word` / `--compact` / `--fields key1,key2` |
| Recent bridge events | `office-bridge events word --limit 20` |
| Stream live selection | `office-bridge watch-selection word` |
| Bridge server / CLI self-check | `office-bridge status` |
| Start bridge | `office-bridge serve` |
| MCP stdio for hosts | `office-bridge mcp-serve [--url URL]` |
| Rich state dump | `office-bridge state word --compact` |
| Event stream / polling | `office-bridge poll word` (interval, event filters) |

Global output shaping: `--compact`, `--fields`, `--max-tokens` (events); `--url` for non-default bridge base (`packages/bridge/src/cli.ts`).

## Common workflows

**Local dev (repo)** — from `README.md`: `pnpm install`, `pnpm setup:word`, `pnpm bridge:serve`, `pnpm dev-server:word`, `pnpm start:word`, open the Word MCP Bridge taskpane; then e.g. `pnpm exec office-bridge list`, `summary word`, `snapshot word`, `status`, `pnpm bridge:mcp`.

**MCP host** — copy-paste patterns in `packages/bridge/README.md` (Claude Desktop JSON, Cursor `mcp.json`, `claude mcp add`, `codex mcp add`, `~/.codex/config.toml`).

**Auth** — default token path is usually enough when bridge and MCP run as the same user; else `OFFICE_BRIDGE_TOKEN` + `mcp-serve` (`packages/bridge/README.md`).

## Troubleshooting

- **Empty `list` / MCP tools fail**: ensure `serve` is running, Word taskpane is open, add-in connected to the bridge URL.
- **Wrong bridge**: pass `--url` consistently on CLI subcommands and `mcp-serve`.
- **TLS / localhost**: bridge TLS is localhost-oriented (`packages/bridge/README.md`).
- **More commands** (`inspect`, `tool`, `exec`, `vfs`, `diag`, `dom`, …): see usage in `packages/bridge/src/cli.ts`.

## Copy to global skills

This file is repo-local under `.claude/skills/`. To reuse elsewhere, copy the folder into your user/global Claude skills location unchanged.
