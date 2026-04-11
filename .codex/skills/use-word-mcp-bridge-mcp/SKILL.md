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
3. Word add-in taskpane connected to that bridge. The add-in install is persistent, but the open taskpane state may need to be restored after Word or a document window is reopened. Flow and commands: root `README.md`.

Once the shared runtime is attached, the visible panel does not have to remain open. Hidden-but-active sessions are still valid bridge targets.

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
- `get_bridge_status`
- `get_session_snapshot` — optional `session`
- `get_live_context` — optional `session`
- `get_recent_events` — optional `session`, `limit` (1–200)
- `call_bridge_tool` — `toolName`, optional `args`, optional `session`
- `word_list_documents`
- `run_unsafe_office_js` — `code`; optional `explanation`, `session`; **only if** session has `unsafe_office_js`
- `vfs_list` — optional `prefix`, `session`
- `vfs_read` — `path`, optional `encoding` (`text`|`base64`), `session`
- `vfs_write` — `path`, optional `text` / `dataBase64`, `session`
- `vfs_delete` — `path`, optional `session`

With a single connected session, `session` may be omitted; with multiple sessions, pass a unique selector from `list_sessions` or `word_list_documents`. Do not assume `word` is a safe selector when two docs are open.

## Current live behavior notes

- The bridge currently exposes a **31-tool Word surface** including exact paragraph range reads/writes, scoped formatting, revision scope reads, and fuller native comment lifecycle tools.
- `get_bridge_status` is the fastest high-level liveness read. Use it before deeper MCP calls when you suspect stale sessions or reconnect churn.
- `word_list_documents` is the easiest way to confirm which live Word documents are actually attached before running a write tool.
- `list_sessions` now returns a compact session record on fresh MCP hosts. If an older host still shows huge `recentEvents` payloads, restart that host's `office-bridge mcp-serve` process.
- `word_search_and_replace` intentionally fails closed when `targetMatchIndexes` points at a truncated or unavailable candidate. That is the safe path, not a bridge failure.
- Preferred mutation flow for repeated phrases:
  1. `call_bridge_tool` with `toolName: "word_search_text"`
  2. inspect returned `matchIndex`, `paragraphIndex`, and offsets
  3. rerun `word_search_and_replace` with matching `targetMatchIndexes` and `maxMatches`, or prefer `word_replace_text_range` for exact edits

## If something breaks

- No sessions: verify bridge is running (`pnpm exec office-bridge status` / `list`). If Word was reopened, the add-in install is still present; the usual missing piece is the taskpane or shared runtime not being reattached yet, so reopen `Word MCP Bridge` from Word's Add-ins UI or use `scripts/bridge/launch-word-taskpane.sh --mode open`.
- Empty context: check selection/focus in Word; try `get_session_snapshot` for the full picture.
- One session looks stale while another is healthy: compare `get_bridge_status` output and `get_recent_events` before assuming the tool layer is broken.
- Unsafe JS rejected: expected without capability — use `call_bridge_tool` or safer RPC paths instead.

## Docs map

- `README.md` — repo layout, dev vs hosted add-in, packaging.
- `packages/bridge/README.md` — npm/npx, CLI examples, MCP host configs, token.
- `packages/bridge/src/mcp.ts` — MCP tools.
- `packages/bridge/src/cli.ts` — CLI entry and usage.
