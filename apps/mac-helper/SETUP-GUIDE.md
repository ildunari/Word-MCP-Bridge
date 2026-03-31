# Word MCP Bridge Setup Guide

## Recommended one-time install

1. Open `Word MCP Bridge Helper.app`.
2. Reveal the hosted manifest from the helper.
3. Open Microsoft Word.
4. In Word, install the add-in using that manifest file.
5. Open the `Word MCP Bridge` taskpane from the Word ribbon.

## Daily workflow

1. Open `Word MCP Bridge Helper.app`.
2. Open Word.
3. Open the `Word MCP Bridge` taskpane.
4. Use your MCP-capable agent.

## If you are developing locally

Use the local manifest instead:

```bash
pnpm install
pnpm setup:word
pnpm bridge:serve
pnpm dev-server:word
pnpm start:word
```

Then open the taskpane in Word and continue with your normal CLI or MCP workflow.

## MCP reminder

The Word add-in and the local bridge are separate:

- `office-bridge serve` runs the local bridge that Word connects to.
- `office-bridge mcp-serve` runs the stdio MCP server your agents connect to.
