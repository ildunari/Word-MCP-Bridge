# Word-MCP-Bridge

`Word-MCP-Bridge` packages a live local bridge for Microsoft Word so MCP hosts, CLI tools, and local agents can inspect and control an open Word taskpane session.

This repo is intentionally smaller than `office-agents-hybrid`. It contains:

- `packages/bridge`: standalone HTTPS/WebSocket bridge server, CLI, and stdio MCP server
- `packages/word-addin`: minimal Word taskpane add-in that exposes bridge status, document metadata, live selection context, and capabilities
- `apps/mac-helper`: native SwiftUI menu bar helper for bridge lifecycle, counters, and quick actions
- `scripts/`: packaging helpers for the Word add-in bundle

The full chat side panel and old SDK/core runtime stack are intentionally out of scope.

## Repo surfaces

- Bridge server: `office-bridge serve`
- MCP server: `office-bridge mcp-serve`
- CLI inspection: `office-bridge list`, `summary`, `snapshot`, `events`, `watch-selection`
- Word add-in: local manifest and hosted manifest

## Local developer install

Use this mode when you are running everything from the repo on your machine.

```bash
pnpm install
pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Then open the `Word MCP Bridge` taskpane in Word.

Useful local commands:

```bash
pnpm exec office-bridge list
pnpm exec office-bridge summary word
pnpm exec office-bridge snapshot word
pnpm exec office-bridge status
pnpm bridge:mcp
pnpm helper:run
```

## Hosted add-in mode

Use this mode when the add-in UI is already hosted and you mainly want to run the local bridge and MCP server.

1. Deploy `packages/word-addin/dist` to your static host.
2. Use `packages/word-addin/manifest.prod.xml` as the production manifest.
3. Sideload or distribute that manifest to Word users.
4. Have each user run the local bridge server on their machine:

```bash
office-bridge serve
office-bridge mcp-serve
```

The hosted taskpane still connects back to the local bridge on `https://localhost:4017`.

## MCP host setup

See [`packages/bridge/README.md`](packages/bridge/README.md) for copy-paste setup examples for:

- Claude Desktop
- Cursor
- Claude Code
- Codex CLI
- generic stdio MCP hosts

## Word add-in bundle

Build and package the distributable add-in bundle:

```bash
pnpm release:bundle
```

That produces:

- `release/word-addin-bundle/`
- `release/word-addin-bundle.zip`
- `release/mac-helper/Word MCP Bridge Helper.app`

The bundle contains:

- `manifest.xml`
- `manifest.prod.xml`
- the built `dist/` taskpane assets
- the native macOS helper app bundle

## macOS helper app

The helper app is a lightweight menu bar controller for local development and distribution demos. It can:

- poll `https://127.0.0.1:4017/status`
- start the local bridge with `pnpm bridge:serve`
- stop the local bridge through the bridge shutdown endpoint
- show live totals for sessions, tool calls, errors, drops, and pending requests
- copy a ready-to-paste MCP config block

Build or run it directly from the repo:

```bash
pnpm helper:build
pnpm helper:run
pnpm package:helper
```

The packaged `.app` assumes it can still find this repo checkout, or that `WORD_MCP_BRIDGE_REPO_ROOT` points at one.

## Repo-local skills

This repo ships agent-facing usage skills in both `.claude/skills/` and `.codex/skills/` for:

- `use-word-mcp-bridge-mcp`
- `use-word-mcp-bridge-cli`

They are meant to work as repo-local guidance first, and can also be copied into a global skill setup if desired.

## What the add-in exposes

- connection state
- bridge URL and session identifiers
- document ID and lightweight document stats
- live selection preview, focus target, and tracking mode
- capability state for the current bridge session
- setup guidance when the bridge session is not yet connected

## Notes

- The add-in forces bridge mode on by default.
- The current capability set is intentionally small: live observation plus privileged raw Office.js execution.
- The add-in uses the selected protocol-node icon set in `packages/word-addin/public/assets/`.
