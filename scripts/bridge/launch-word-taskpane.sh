#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
SWIFT_HELPER="$ROOT_DIR/scripts/bridge/word_ax.swift"

mode="open"
wait_for_bridge=1
bridge_url="${BRIDGE_URL:-https://localhost:4017}"
timeout_seconds=8
word_app="${WORD_APP:-Microsoft Word}"

usage() {
  cat <<'EOF'
Usage: scripts/bridge/launch-word-taskpane.sh [options]

Open or focus the Word MCP Bridge taskpane in Microsoft Word on macOS.

Options:
  --mode <open|status|show-addins>
      `open` tries to focus an existing pane, then clicks the `Open Word MCP Bridge`
      ribbon button, then falls back to `Add-ins`.
      `status` only reports whether the pane/button can be found.
      `show-addins` only opens the generic Add-ins surface.
  --bridge-url <url>
      Local bridge URL to poll after launching. Default: https://localhost:4017
  --timeout <seconds>
      How long to wait for a live Word bridge session after a successful open.
      Default: 8
  --no-bridge-wait
      Skip bridge-session polling after UI automation succeeds.
  --help
      Show this help.
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --mode)
      mode="${2:-}"
      shift 2
      ;;
    --bridge-url)
      bridge_url="${2:-}"
      shift 2
      ;;
    --timeout)
      timeout_seconds="${2:-}"
      shift 2
      ;;
    --no-bridge-wait)
      wait_for_bridge=0
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 64
      ;;
  esac
done

if [[ ! -f "$SWIFT_HELPER" ]]; then
  echo "Missing helper: $SWIFT_HELPER" >&2
  exit 66
fi

baseline_word_session_count="0"
baseline_word_session_ids="[]"

if [[ "$mode" == "open" && $wait_for_bridge -eq 1 ]] && command -v pnpm >/dev/null 2>&1; then
  baseline_status_output="$(pnpm exec office-bridge --url "$bridge_url" status 2>/dev/null || true)"
  if [[ -n "$baseline_status_output" ]]; then
    baseline_probe="$(printf '%s' "$baseline_status_output" | node -e '
let raw = "";
process.stdin.on("data", (chunk) => raw += chunk);
process.stdin.on("end", () => {
  try {
    const status = JSON.parse(raw);
    const sessions = Array.isArray(status.sessions)
      ? status.sessions.filter((session) => session?.snapshot?.app === "word")
      : [];
    const ids = sessions
      .map((session) => session?.snapshot?.sessionId)
      .filter((value) => typeof value === "string" && value.length > 0);
    process.stdout.write(JSON.stringify({ count: sessions.length, ids }));
  } catch {
    process.stdout.write(JSON.stringify({ count: 0, ids: [] }));
  }
});
')"
    baseline_word_session_count="$(printf '%s' "$baseline_probe" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(parsed.count ?? 0)); });')"
    baseline_word_session_ids="$(printf '%s' "$baseline_probe" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(JSON.stringify(parsed.ids ?? [])); });')"
  fi
fi

if [[ "$mode" != "status" ]]; then
  if [[ "$(osascript -e "application \"$word_app\" is running" 2>/dev/null || true)" != "true" ]]; then
    echo "Launching $word_app..."
    open -a "$word_app"
    sleep 3
  else
    osascript -e "tell application \"$word_app\" to activate" >/dev/null 2>&1 || true
  fi

  osascript <<EOF >/dev/null 2>&1 || true
tell application "$word_app"
  activate
  if (count of documents) = 0 then
    make new document
  end if
end tell
EOF
  sleep 2
fi

launch_started_at="$(node -e 'console.log(Date.now())')"

set +e
swift_output="$(swift "$SWIFT_HELPER" --mode "$mode" 2>&1)"
swift_status=$?
set -e

printf '%s\n' "$swift_output"

if [[ $swift_status -ne 0 && ! ( "$mode" == "open" && $swift_status -eq 2 ) ]]; then
  exit "$swift_status"
fi

if [[ "$mode" != "open" || $wait_for_bridge -eq 0 ]]; then
  if [[ $swift_status -eq 2 ]]; then
    exit 2
  fi
  exit 0
fi

if ! command -v pnpm >/dev/null 2>&1; then
  echo "pnpm not found; skipping bridge-session verification." >&2
  exit 0
fi

deadline=$((SECONDS + timeout_seconds))
while (( SECONDS < deadline )); do
  if status_output="$(pnpm exec office-bridge --url "$bridge_url" status 2>/dev/null)"; then
    probe_result="$(printf '%s' "$status_output" | node -e "
let raw = '';
process.stdin.on('data', (chunk) => raw += chunk);
process.stdin.on('end', () => {
  try {
    const status = JSON.parse(raw);
    const sessions = Array.isArray(status.sessions)
      ? status.sessions.filter((session) => session?.snapshot?.app === 'word')
      : [];
    const count = sessions.length;
    const hasRecent = sessions.some((session) => Number(session?.connectedAt ?? 0) >= Number(process.argv[1]));
    const hasRecentSeen = sessions.some((session) =>
      Number(session?.lastSeenAt ?? 0) >= Number(process.argv[1]) ||
      Number(session?.snapshot?.updatedAt ?? 0) >= Number(process.argv[1])
    );
    const baselineIds = new Set(JSON.parse(process.argv[2] ?? '[]'));
    const currentIds = sessions
      .map((session) => session?.snapshot?.sessionId)
      .filter((value) => typeof value === 'string' && value.length > 0);
    const newIds = currentIds.filter((id) => !baselineIds.has(id));
    const allBaselineIdsPresent = currentIds.length > 0 && [...baselineIds].every((id) => currentIds.includes(id));
    process.stdout.write(JSON.stringify({ count, hasRecent, hasRecentSeen, newIds, allBaselineIdsPresent }));
  } catch {
    process.stdout.write(JSON.stringify({ count: 0, hasRecent: false, hasRecentSeen: false, newIds: [], allBaselineIdsPresent: false }));
  }
});
" "$launch_started_at" "$baseline_word_session_ids")"
    current_count="$(printf '%s' "$probe_result" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(parsed.count ?? 0)); });')"
    has_recent="$(printf '%s' "$probe_result" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(Boolean(parsed.hasRecent))); });')"
    has_recent_seen="$(printf '%s' "$probe_result" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(Boolean(parsed.hasRecentSeen))); });')"
    new_id_count="$(printf '%s' "$probe_result" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(Array.isArray(parsed.newIds) ? parsed.newIds.length : 0)); });')"
    all_baseline_ids_present="$(printf '%s' "$probe_result" | node -e 'let raw=""; process.stdin.on("data", c => raw += c); process.stdin.on("end", () => { const parsed = JSON.parse(raw); process.stdout.write(String(Boolean(parsed.allBaselineIdsPresent))); });')"

    if [[ "$has_recent" == "true" || "$has_recent_seen" == "true" || "$new_id_count" -gt 0 ]]; then
      echo "Live Word bridge session confirmed."
      exit 0
    fi

    if [[ "$swift_output" == *"already open"* && "$baseline_word_session_count" -ge 1 && "$current_count" -ge "$baseline_word_session_count" && "$all_baseline_ids_present" == "true" && "$has_recent_seen" == "true" ]]; then
      echo "Live Word bridge session confirmed."
      exit 0
    fi
  fi
  sleep 1
done

echo "Taskpane launch was attempted, but no live Word bridge session appeared within ${timeout_seconds}s." >&2
exit 2
