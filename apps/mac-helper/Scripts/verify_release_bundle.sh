#!/usr/bin/env bash
set -euo pipefail

APP_BUNDLE="${1:?expected app bundle path}"
APP_ZIP="${2:?expected app zip path}"

require_file() {
  local path="$1"
  if [[ ! -e "$path" ]]; then
    echo "Missing required release artifact: $path" >&2
    exit 1
  fi
}

require_file "$APP_BUNDLE"
require_file "$APP_ZIP"
require_file "$APP_BUNDLE/Contents/MacOS/WordMCPBridgeHelper"
require_file "$APP_BUNDLE/Contents/Resources/setup/manifest.prod.xml"
require_file "$APP_BUNDLE/Contents/Resources/setup/SETUP-GUIDE.md"
require_file "$APP_BUNDLE/Contents/Resources/setup/release-metadata.json"
require_file "$APP_BUNDLE/Contents/Resources/taskpane/serve_taskpane.mjs"
require_file "$APP_BUNDLE/Contents/Resources/taskpane-dist/taskpane.html"
require_file "$APP_BUNDLE/Contents/Resources/runtime/node/bin/node"
require_file "$APP_BUNDLE/Contents/Resources/runtime/node/lib/node_modules/npm/package.json"
require_file "$APP_BUNDLE/Contents/Resources/runtime/bin/office-bridge"
require_file "$APP_BUNDLE/Contents/Resources/runtime/bin/word-mcp-bridge-word-launcher"
require_file "$APP_BUNDLE/Contents/Resources/runtime/bridge/node_modules/@word-mcp-bridge/bridge/bin/office-bridge.js"

if grep -q 'pnpm' "$APP_BUNDLE/Contents/Resources/runtime/bin/office-bridge"; then
  echo "Packaged office-bridge wrapper still references pnpm." >&2
  exit 1
fi

if grep -q 'WORD_MCP_BRIDGE_REPO_ROOT' "$APP_BUNDLE/Contents/Resources/runtime/bin/office-bridge"; then
  echo "Packaged office-bridge wrapper still references repo-root fallback." >&2
  exit 1
fi
