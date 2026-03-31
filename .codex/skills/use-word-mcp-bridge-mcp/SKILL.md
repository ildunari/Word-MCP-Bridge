---
name: use-word-mcp-bridge-mcp
description: Word MCP Bridge workflow for Codex/GPT-style agents — run the local bridge (`office-bridge serve`), attach the stdio MCP server (`office-bridge mcp-serve`), then use MCP tools to inspect Word sessions, read live context, call bridge tools, manage VFS, or run gated unsafe Office.js. Use for Word-MCP-Bridge, `@word-mcp-bridge/bridge`, MCP registration (`codex mcp add`), localhost bridge https://localhost:4017, or deciding MCP vs terminal CLI. Triggers on bridge MCP setup, `list_sessions`, `call_bridge_tool`, Word taskpane not connecting, or "no sessions" from MCP.
---

## MCP vs CLI (pick one surface)

- **MCP (`packages/bridge/src/mcp.ts`)** — Use inside Codex/Cursor/Claude when the host should call tools: sessions, snapshot, live context, events, `call_bridge_tool`, VFS, `run_unsafe_office_js` (capability-gated).
- **CLI (`packages/bridge/src/cli.ts`)** — Use from a shell for anything MCP does not wrap: `watch-selection`, `watch-context`, `poll`, screenshots, `dom`, `assert`, `bench`, disk `vfs pull`/`push`, `reset`, etc. Full command list is in the CLI help text and `packages/bridge/README.md`.

## Hard requirements

1. Bridge process: `office-bridge serve` (repo: `pnpm bridge:serve`).
2. MCP process: `office-bridge mcp-serve --url https://localhost:4017` (repo: `pnpm bridge:mcp`). MCP talks to the bridge over HTTP; it is not a substitute for `serve`.
3. Word add-in taskpane connected to that bridge. Flow and commands: root `README.md`.

If tool calls fail with auth errors, use `OFFICE_BRIDGE_TOKEN` as in `packages/bridge/README.md`.

## Quick setup (from repo)

```bash
pnpm install && pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Register MCP with Codex (example from `packages/bridge/README.md`):

```bash
codex mcp add word-mcp-bridge -- office-bridge mcp-serve --url https://localhost:4017
```

Use `npx -y @word-mcp-bridge/bridge` in place of `office-bridge` when not globally installed.

## MCP tool list

Source of truth: `packages/bridge/src/mcp.ts`.

- `list_sessions`
- `get_session_snapshot` — optional `session`
- `get_live_context` — optional `session`
- `get_recent_events` — optional `session`, `limit` (1–200)
- `call_bridge_tool` — `toolName`, optional `args`, optional `session`
- `run_unsafe_office_js` — `code`; optional `explanation`, `session`; **only if** session has `unsafe_office_js`
- `vfs_list` — optional `prefix`, `session`
- `vfs_read` — `path`, optional `encoding` (`text`|`base64`), `session`
- `vfs_write` — `path`, optional `text` / `dataBase64`, `session`
- `vfs_delete` — `path`, optional `session`

With a single connected session, `session` may be omitted; with multiple sessions, pass a unique selector (same idea as CLI session matching).

## If something breaks

- No sessions: verify bridge is running (`pnpm exec office-bridge status` / `list`) and Word pane is open and connected.
- Empty context: check selection/focus in Word; try `get_session_snapshot` for the full picture.
- Unsafe JS rejected: expected without capability — use `call_bridge_tool` or safer RPC paths instead.

## Docs map

- `README.md` — repo layout, dev vs hosted add-in, packaging.
- `packages/bridge/README.md` — npm/npx, CLI examples, MCP host configs, token.
- `packages/bridge/src/mcp.ts` — MCP tools.
- `packages/bridge/src/cli.ts` — CLI entry and usage.
