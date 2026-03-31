# Word-MCP-Bridge

Minimal Word bridge project for live Office session control through a local bridge server, CLI, and stdio MCP server.

This repo is intentionally smaller than `office-agents-hybrid`. It contains:

- `packages/bridge`: standalone HTTPS/WebSocket bridge server, CLI, and stdio MCP server
- `packages/word-addin`: tiny Word taskpane add-in that shows bridge status, document metadata, live selection context, and enabled capabilities

The full chat side panel and old SDK/core runtime stack are intentionally out of scope for this repo.

## Quick Start

```bash
pnpm install
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Then open the `Word MCP Bridge` taskpane in Word. Once the pane is open, use:

```bash
pnpm exec office-bridge list
pnpm exec office-bridge summary word
pnpm exec office-bridge snapshot word
pnpm exec office-bridge mcp-serve
```

## What The Add-in Exposes

- connection state
- bridge URL and session identifiers
- document ID and lightweight document stats
- live selection preview, focus target, and tracking mode
- capability state for the current bridge session

## Notes

- The add-in forces bridge mode on by default.
- The current capability set is intentionally small: live observation plus privileged raw Office.js execution.
- Icon assets are minimal placeholders for local sideloading and manifest validation.
