#!/usr/bin/env bash
set -euo pipefail

APP_BUNDLE="${1:?expected app bundle path}"
OFFICE_BRIDGE="$APP_BUNDLE/Contents/Resources/runtime/bin/office-bridge"
TASKPANE_NODE="$APP_BUNDLE/Contents/Resources/runtime/node/bin/node"
TASKPANE_SERVER="$APP_BUNDLE/Contents/Resources/taskpane/serve_taskpane.mjs"
TASKPANE_LAUNCHER="$APP_BUNDLE/Contents/Resources/runtime/bin/word-mcp-bridge-word-launcher"

TMP_DIR="$(mktemp -d)"
PORT="4417"
TASKPANE_PORT="3414"
BRIDGE_PID=""
TASKPANE_PID=""

cleanup() {
  if [[ -n "$BRIDGE_PID" ]]; then
    kill "$BRIDGE_PID" >/dev/null 2>&1 || true
  fi
  if [[ -n "$TASKPANE_PID" ]]; then
    kill "$TASKPANE_PID" >/dev/null 2>&1 || true
  fi
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

require_file() {
  local path="$1"
  if [[ ! -e "$path" ]]; then
    echo "Missing required bundle path: $path" >&2
    exit 1
  fi
}

wait_for_url() {
  local url="$1"
  local attempts="${2:-20}"
  for _ in $(seq 1 "$attempts"); do
    if curl -ksSf "$url" >/dev/null 2>&1; then
      return 0
    fi
    sleep 0.5
  done
  echo "Timed out waiting for $url" >&2
  exit 1
}

require_file "$OFFICE_BRIDGE"
require_file "$TASKPANE_NODE"
require_file "$TASKPANE_SERVER"
require_file "$TASKPANE_LAUNCHER"

"$OFFICE_BRIDGE" >/dev/null
"$TASKPANE_LAUNCHER" --help >/dev/null

TASKPANE_ROOT="$TMP_DIR/taskpane"
mkdir -p "$TASKPANE_ROOT"
printf '<!doctype html><html><body>ok</body></html>' > "$TASKPANE_ROOT/taskpane.html"

/usr/bin/openssl req -x509 -nodes -newkey rsa:2048 -sha256 -days 2 \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
  -keyout "$TMP_DIR/localhost.key" \
  -out "$TMP_DIR/localhost.crt" >/dev/null 2>&1

OFFICE_BRIDGE_CERT="$TMP_DIR/localhost.crt" \
OFFICE_BRIDGE_KEY="$TMP_DIR/localhost.key" \
"$OFFICE_BRIDGE" serve --port "$PORT" >/dev/null 2>&1 &
BRIDGE_PID="$!"

wait_for_url "https://localhost:${PORT}/health"

"$TASKPANE_NODE" "$TASKPANE_SERVER" \
  --root "$TASKPANE_ROOT" \
  --port "$TASKPANE_PORT" \
  --cert "$TMP_DIR/localhost.crt" \
  --key "$TMP_DIR/localhost.key" >/dev/null 2>&1 &
TASKPANE_PID="$!"

wait_for_url "https://localhost:${TASKPANE_PORT}/healthz"
curl -ksSf "https://localhost:${TASKPANE_PORT}/taskpane.html" >/dev/null
