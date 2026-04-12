# Word MCP Bridge Setup Guide

## Recommended one-time install

1. Open `Word MCP Bridge Helper.app`.
2. Click `Install in Word`.
3. If the helper says Word must restart, restart Word once.
4. Click `Open Word MCP Bridge`.

## Daily workflow

1. Open `Word MCP Bridge Helper.app`.
2. Open Word, or let the helper reopen it.
3. Click `Open Word MCP Bridge` if the panel is not already visible.
4. Use your MCP-capable agent.

## If you are developing locally

Use the development manifest instead:

```bash
pnpm install
pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Then open the taskpane in Word and continue with your normal CLI or MCP workflow.

## MCP reminder

The Word taskpane server, add-in, and local bridge are separate:

- the helper serves the local taskpane UI that Word loads
- `office-bridge serve` runs the local bridge that Word connects to.
- `office-bridge mcp-serve` runs the stdio MCP server your agents connect to.

The helper now owns the normal production install flow. Use the dev manifest only from the advanced developer section when you are actively working from this repo.
