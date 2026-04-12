# Word-MCP-Bridge

`Word-MCP-Bridge` packages a live local bridge for Microsoft Word so MCP hosts, CLI tools, and local agents can inspect and control an open Word taskpane session.

This repo is intentionally smaller than `office-agents-hybrid`. It contains:

- `packages/bridge`: standalone HTTPS/WebSocket bridge server, CLI, and stdio MCP server
- `packages/word-addin`: minimal Word taskpane add-in that shows the live document dashboard, selection context, and review-ready details for the current Word file
- `apps/mac-helper`: native SwiftUI menu bar helper for bridge lifecycle, counters, and quick actions
- `scripts/`: packaging helpers for the Word add-in bundle

The full chat side panel and old SDK/core runtime stack are intentionally out of scope.

## Repo surfaces

- Bridge server: `office-bridge serve`
- MCP server: `office-bridge mcp-serve`
- CLI inspection: `office-bridge list`, `summary`, `snapshot`, `events`, `watch-selection`
- Word add-in: local production manifest and local development manifest

## Best user flow

The intended user flow is:

1. Install the Word add-in once with the local production manifest.
2. Put `Word MCP Bridge Helper.app` in `/Applications`.
3. Optionally enable `Launch helper at login` in the helper settings.
4. For normal use, open the helper, open Word, and open the `Word MCP Bridge` taskpane.

After the taskpane connects, shared-runtime mode can keep the bridge alive even if the panel is hidden. The add-in install is persistent, but the visible open-pane state may still need to be restored for a given Word window after Word restarts or document windows are reopened.

The helper app now includes:

- a first-run setup window
- a local-first install path
- quick actions for Word, manifests, docs, and MCP config
- optional bridge health notifications
- startup preferences for launch-at-login and auto-start bridge

## One-time install

### 1. Install the helper app

Build and package everything:

```bash
pnpm release:bundle
```

That produces:

- `release/word-addin-bundle/`
- `release/word-addin-bundle.zip`
- `release/mac-helper/Word MCP Bridge Helper.app`

Copy the helper app into `/Applications` and open it.

### 2. Install the Word add-in once

Use the helper app's `Getting Started` flow and choose the local production install path.

That flow will:

- reveal the local production manifest
- open Word
- tell you the exact next step inside Word
- let you verify when the taskpane has connected back to the bridge

If you want to do it manually, use:

- local production manifest: `packages/word-addin/manifest.prod.xml`
- local dev manifest: `packages/word-addin/manifest.xml`

## Daily workflow

After the one-time install:

1. Open `Word MCP Bridge Helper.app`.
2. Open Word.
3. Open the `Word MCP Bridge` taskpane in Word.
4. Use your MCP-capable host.

If the taskpane is not visible after reopening Word, use the helper or `scripts/bridge/launch-word-taskpane.sh --mode open` to restore it quickly. A hidden shared-runtime session still counts as healthy once it is attached.

If you enabled `Launch helper at login` and `Auto-start bridge when helper opens`, the helper should handle most of the local bridge setup automatically.

## Local developer install

Use this mode when you are running everything from the repo on your machine and want the local manifest + dev server path.

```bash
pnpm install
pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Then open the `Word MCP Bridge` taskpane in Word.

For faster recovery during local development, you can also use:

```bash
scripts/bridge/launch-word-taskpane.sh --mode open
```

Useful local commands:

```bash
pnpm exec office-bridge list
pnpm exec office-bridge summary word
pnpm exec office-bridge snapshot word
pnpm exec office-bridge status
pnpm bridge:mcp
pnpm bridge:smoke
pnpm helper:run
```

`pnpm bridge:smoke` runs a focused live smoke check for:

- launcher recovery
- compact `list_sessions` MCP output
- shared-runtime hidden-session liveness reporting
- table read/update behavior
- rejection of unsafe zero `lineSpacing`

To keep the active sidebar stable, the smoke run is read-first by default. Add `--launch` only when you want it to reopen the pane, and add `--write` only when you want it to exercise a live table cell update.

## Local production mode

Use this mode for normal users. The helper app serves the taskpane UI locally on the user machine, and the taskpane connects to the local bridge on `https://localhost:4017`.

1. Package the helper app and add-in assets.
2. Use `packages/word-addin/manifest.prod.xml` as the production manifest.
3. Install that manifest in Word once.
4. Keep `Word MCP Bridge Helper.app` running while using the taskpane.

## MCP host setup

See [`packages/bridge/README.md`](packages/bridge/README.md) for copy-paste setup examples for:

- Claude Desktop
- Cursor
- Claude Code
- Codex CLI
- generic stdio MCP hosts

## macOS helper app

The helper app is now the main setup and operations surface. It can:

- poll `https://localhost:4017/status`
- start the local bridge with `pnpm bridge:serve`
- stop the local bridge through the bridge shutdown endpoint
- show live totals for sessions, tool calls, errors, drops, and pending requests
- open Word, manifests, and setup docs
- copy a ready-to-paste MCP config block
- guide first-run installation
- optionally launch at login and auto-start the bridge
- optionally notify when Word disconnects or reconnects

Build or run it directly from the repo:

```bash
pnpm helper:build
pnpm helper:run
pnpm package:helper
```

The packaged `.app` uses bundled setup assets and bundled taskpane files for local production use. When running from the repo, it can still fall back to repo-local manifests and docs.

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
- shared-runtime visibility state, including hidden-but-active sessions
- capability state for the current bridge session
- setup guidance when the bridge session is not yet connected

## Notes

- The add-in forces bridge mode on by default.
- With shared-runtime enabled, the bridge can remain usable after the panel is hidden, but Word may still reopen some windows with the taskpane closed.
- The current capability set is intentionally small: live observation plus privileged raw Office.js execution.
- The add-in uses the selected protocol-node icon set in `packages/word-addin/public/assets/`.
