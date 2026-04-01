# AGENTS.md

## Cursor Cloud specific instructions

### Overview

Word-MCP-Bridge is a pnpm workspace monorepo with two TypeScript packages (`packages/bridge`, `packages/word-addin`) and a macOS-only Swift helper app (`apps/mac-helper`). No databases, Docker, or external APIs are required. See the root `README.md` for full documentation.

### Key commands

| Task | Command |
|---|---|
| Install deps | `pnpm install` |
| Build all | `pnpm build` |
| Run tests | `pnpm test` (Vitest, 149 tests across both packages) |
| Lint | `pnpm exec biome check .` (pre-existing issues in `dist/` and `.claude/skills/` are expected) |
| Type-check bridge | `cd packages/bridge && pnpm exec tsc -p tsconfig.build.json --noEmit` |
| Type-check word-addin | `cd packages/word-addin && pnpm exec svelte-check` (2 pre-existing errors) |
| Start bridge server | `pnpm bridge:serve` (HTTPS on `https://localhost:4017`) |
| Start word-addin dev server | `pnpm dev-server:word` (Vite on `https://localhost:3013`) |
| CLI status | `pnpm exec office-bridge status` |
| CLI list sessions | `pnpm exec office-bridge list` |

### Gotchas

- **Self-signed TLS certs**: The build step (`pnpm build`) auto-generates dev certs via `office-addin-dev-certs` on first run. Both the bridge server and Vite dev server use localhost TLS. When using `curl`, pass `-k` to skip certificate verification.
- **esbuild build scripts ignored**: pnpm shows a warning about ignored build scripts (esbuild, keytar, etc.). This does **not** block builds or tests; Vite/esbuild work fine without the native install script.
- **No E2E without Word desktop**: Full end-to-end testing requires Microsoft Word with the taskpane add-in loaded, which is not available in the Linux VM. The bridge server, dev server, CLI commands, and all automated tests work without Word.
- **macOS helper**: `apps/mac-helper` is Swift/SwiftUI and only builds on macOS. Skip it in Linux environments.
- **Bridge server must be running for CLI commands**: Commands like `office-bridge status`, `list`, `summary`, etc. require the bridge server to be running on port 4017.
- **The taskpane auto-connects to the bridge**: When the Vite dev server is open in a browser while the bridge server runs, the taskpane WebSocket client connects automatically.
