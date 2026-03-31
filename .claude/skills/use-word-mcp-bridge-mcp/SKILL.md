---
name: use-word-mcp-bridge-mcp
description: Operate the Word MCP Bridge stack — local HTTPS bridge plus stdio MCP server — to read live Word taskpane context, list sessions, call bridge tools, use VFS helpers, and optionally run privileged Office.js when the session allows it. Use when the user works in Word-MCP-Bridge, `@word-mcp-bridge/bridge`, `office-bridge`, `mcp-serve`, localhost Word bridge (default https://localhost:4017), MCP host setup (Claude Desktop, Cursor, Claude Code), or needs to choose between MCP tools and the `office-bridge` CLI. Triggers on "word mcp bridge", "bridge MCP", "call_bridge_tool", "list_sessions", live Word selection/context via bridge, or troubleshooting "no bridge sessions" / add-in not connected.
---

## When to use MCP vs CLI

**Prefer the MCP server** when you are inside an MCP-capable host (Claude Code, Claude Desktop, Cursor, Codex) and want structured tool calls: session discovery, snapshots, live context, recent events, `call_bridge_tool`, VFS read/write, or gated `run_unsafe_office_js`. Implementation lives in `packages/bridge/src/mcp.ts`.

**Prefer the CLI** (`packages/bridge/src/cli.ts`, `office-bridge` command) for terminal-first workflows, shell scripts, CI, long-running streams, or features not exposed as MCP tools: `watch-selection`, `watch-context`, `poll`, `screenshot`, `dom`, `assert`, `bench`, `vfs pull`/`push` to local disk paths, `reset`, etc. See printed usage in the CLI and `packages/bridge/README.md`.

## Prerequisites

1. **Bridge server running** — `office-bridge serve` (or from repo: `pnpm bridge:serve`). MCP does not replace this process.
2. **Word add-in connected** — Taskpane open and connected to the bridge (repo default flow uses the Word MCP Bridge add-in). See root `README.md`.
3. **MCP process** — `office-bridge mcp-serve --url https://localhost:4017` (or repo: `pnpm bridge:mcp` after `pnpm bridge:serve`). Host config examples: `packages/bridge/README.md`.

Same-machine installs usually pick up the bridge auth token automatically; if tools fail with auth errors, set `OFFICE_BRIDGE_TOKEN` as documented in `packages/bridge/README.md`.

## Setup flow (repo)

From root `README.md`:

```bash
pnpm install
pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Register MCP in the host using `office-bridge mcp-serve --url https://localhost:4017` (or `npx -y @word-mcp-bridge/bridge` variants). Copy-paste blocks for Claude Desktop, Cursor, Claude Code, and Codex: `packages/bridge/README.md`.

## MCP tools (stdio server)

Registered in `packages/bridge/src/mcp.ts`:

| Tool | Role |
|------|------|
| `list_sessions` | List connected bridge sessions. |
| `get_session_snapshot` | Refresh and return full session snapshot (state, capabilities, gateway context). Optional `session` selector. |
| `get_live_context` | Live context slice (e.g. selection) via `refresh_session`. Optional `session`. |
| `get_recent_events` | Recent events; `limit` 1–200, default 20. Optional `session`. |
| `call_bridge_tool` | Run a named bridge tool with `args`. Optional `session`. Respect session capability boundaries. |
| `run_unsafe_office_js` | Privileged Office.js **only** if the session exposes `unsafe_office_js`. `code` required; `explanation` optional. |
| `vfs_list` | List VFS paths; optional `prefix`. |
| `vfs_read` | Read VFS file; `path`; `encoding` `text` or `base64`. |
| `vfs_write` | Write VFS file; `path`; `text` and/or `dataBase64`. |
| `vfs_delete` | Delete VFS path. |

**Session selector:** If exactly one session exists, `session` can be omitted. If multiple sessions exist, pass an unambiguous selector (same matching rules as CLI session resolution — see bridge session matching in `packages/bridge/src/server.ts` and CLI behavior).

## Troubleshooting

- **"No bridge sessions" / MCP errors on session resolution** — Confirm `office-bridge serve` is up, Word taskpane is open, and add-in shows a connected bridge. Run `pnpm exec office-bridge list` or `pnpm exec office-bridge status` from the repo.
- **Wrong or empty live context** — Selection may be empty; confirm focus/selection in Word. Use `get_session_snapshot` vs `get_live_context` depending on whether you need the full gateway payload.
- **`run_unsafe_office_js` fails or is unavailable** — Session must advertise the capability; do not assume it is on. Prefer `call_bridge_tool` with documented tool names when possible.
- **Auth / HTTPS issues** — Bridge is localhost TLS-oriented; ensure MCP `--url` matches the running bridge. Set `OFFICE_BRIDGE_TOKEN` if needed (`packages/bridge/README.md`).

## Reference files

- `README.md` — Repo overview, dev commands, hosted add-in mode, bundle output.
- `packages/bridge/README.md` — Install modes, CLI cheat sheet, MCP host snippets, `OFFICE_BRIDGE_TOKEN`.
- `packages/bridge/src/mcp.ts` — MCP tool registry and behavior.
- `packages/bridge/src/cli.ts` — Full CLI surface and usage text.
