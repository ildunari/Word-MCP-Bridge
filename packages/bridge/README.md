# @word-mcp-bridge/bridge

Local HTTPS/WebSocket bridge, CLI, and stdio MCP server for `Word-MCP-Bridge`.

This package is the operator surface for:

- running the local bridge server
- exposing the bridge as an MCP stdio server
- inspecting live Word sessions from the CLI

The expected end-user flow is:

1. Install the Word add-in once in Word.
2. Open `Word MCP Bridge Helper.app`.
3. Let the helper start the local bridge.
4. Let your MCP host run `office-bridge mcp-serve`.

## Install

### Global npm install

```bash
npm install -g @word-mcp-bridge/bridge
```

Then use:

```bash
office-bridge serve
office-bridge mcp-serve
office-bridge list
```

### One-off `npx`

```bash
npx -y @word-mcp-bridge/bridge serve
npx -y @word-mcp-bridge/bridge mcp-serve
```

### From the repo

```bash
pnpm install
pnpm bridge:serve
pnpm exec office-bridge mcp-serve
```

## Runtime model

There are two separate processes:

1. `office-bridge serve`
   Starts the local HTTPS/WebSocket bridge that the Word add-in connects to.

2. `office-bridge mcp-serve`
   Starts the stdio MCP server that MCP hosts launch.

The MCP server is only useful when:

- the bridge server is running
- the Word add-in shared runtime is connected to the bridge

The visible taskpane does not have to stay open once the shared runtime is attached, but the open-pane state may still need to be restored after Word restarts or document windows are reopened.

If you are using the packaged helper app, the helper can now guide the first-run setup and reveal the local production manifest for the one-time Word install step.

## Common CLI commands

```bash
office-bridge serve
office-bridge list
office-bridge summary word
office-bridge snapshot word
office-bridge state word --compact
office-bridge events word --limit 20
office-bridge watch-selection word
office-bridge watch-context word
office-bridge mcp-serve
```

Helpful recovery command from the repo:

```bash
scripts/bridge/launch-word-taskpane.sh --mode open
```

Focused live smoke check from the repo:

```bash
pnpm bridge:smoke
```

## MCP host setup

The examples below assume the bridge server is already running locally at `https://localhost:4017`.

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word-mcp-bridge": {
      "command": "office-bridge",
      "args": ["mcp-serve", "--url", "https://localhost:4017"]
    }
  }
}
```

If you are not using a global install, replace `command` and `args` with an `npx` form:

```json
{
  "mcpServers": {
    "word-mcp-bridge": {
      "command": "npx",
      "args": ["-y", "@word-mcp-bridge/bridge", "mcp-serve", "--url", "https://localhost:4017"]
    }
  }
}
```

### Cursor

Add this to project-local `.cursor/mcp.json` or global `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "word-mcp-bridge": {
      "command": "office-bridge",
      "args": ["mcp-serve", "--url", "https://localhost:4017"]
    }
  }
}
```

### Claude Code

Add the server with the Claude Code CLI:

```bash
claude mcp add --transport stdio word-mcp-bridge -- office-bridge mcp-serve --url https://localhost:4017
```

If you prefer not to install globally:

```bash
claude mcp add --transport stdio word-mcp-bridge -- npx -y @word-mcp-bridge/bridge mcp-serve --url https://localhost:4017
```

### Codex CLI

Add the server with the Codex CLI:

```bash
codex mcp add word-mcp-bridge -- office-bridge mcp-serve --url https://localhost:4017
```

Or use `npx`:

```bash
codex mcp add word-mcp-bridge -- npx -y @word-mcp-bridge/bridge mcp-serve --url https://localhost:4017
```

Equivalent `~/.codex/config.toml` entry:

```toml
[mcp_servers.word-mcp-bridge]
command = "office-bridge"
args = ["mcp-serve", "--url", "https://localhost:4017"]
```

### Generic stdio MCP hosts

If your host accepts a `command` + `args` stdio config, use:

```json
{
  "command": "office-bridge",
  "args": ["mcp-serve", "--url", "https://localhost:4017"]
}
```

## Authentication

The local bridge server creates an auth token on first run. When the MCP server and bridge run under the same user account on the same machine, the package can usually read the default token path automatically.

If you want to pass it explicitly, set:

```bash
export OFFICE_BRIDGE_TOKEN="your-token"
```

And then run:

```bash
office-bridge mcp-serve --url https://localhost:4017
```

## Notes

- Local bridge TLS is intentionally localhost-oriented.
- `mcp-serve` wraps the live bridge; it does not replace `serve`.
- The Word add-in is distributed separately from this package.
- If a long-lived MCP host still shows the older verbose `list_sessions` payload after an update, restart that host's `office-bridge mcp-serve` process so it reconnects to the current bridge code.
