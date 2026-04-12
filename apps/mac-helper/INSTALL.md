# Word MCP Bridge Install

## Fast path

1. Drag `Word MCP Bridge Helper.app` into `/Applications`.
2. Open the helper.
3. Click `Install in Word`.
4. Click `Open Word MCP Bridge`.
5. Paste the MCP config from the helper into your agent app.

## What the helper does

- installs or repairs the bundled production Word add-in manifest
- serves the local taskpane on `https://localhost:3014`
- starts the local bridge on `https://localhost:4017`
- reopens the Word MCP Bridge panel when Word needs it

## Notes

- If Word was already open while the helper refreshed the add-in manifest, restart Word once and then click `Open Word MCP Bridge` again.
- `Launch helper at login` and `Auto-open Word MCP Bridge when Word launches` can be enabled from helper settings after the first run.
- Use the helper’s `Copy MCP Config` action for the exact MCP JSON block.
