#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_ROOT="$(cd "$APP_DIR/../.." && pwd)"
BUILD_DIR="$APP_DIR/.build/release"
OUTPUT_DIR="$REPO_ROOT/release/mac-helper"
METADATA_FILE="$APP_DIR/release-metadata.json"
INSTALL_GUIDE="$REPO_ROOT/release/INSTALL.md"
VERIFY_SCRIPT="$APP_DIR/Scripts/verify_release_bundle.sh"

metadata_value() {
  node -e 'const fs=require("fs"); const data=JSON.parse(fs.readFileSync(process.argv[1],"utf8")); process.stdout.write(String(data[process.argv[2]] ?? ""));' \
    "$METADATA_FILE" "$1"
}

APP_NAME="${WORD_MCP_BRIDGE_APP_NAME:-$(metadata_value appName)}.app"
BUNDLE_ID="${WORD_MCP_BRIDGE_BUNDLE_ID:-$(metadata_value bundleIdentifier)}"
APP_VERSION="${WORD_MCP_BRIDGE_APP_VERSION:-$(metadata_value version)}"
BUILD_NUMBER="${WORD_MCP_BRIDGE_BUILD_NUMBER:-$(metadata_value buildNumber)}"
MINIMUM_MACOS_VERSION="$(metadata_value minimumMacOSVersion)"
APP_BUNDLE="$OUTPUT_DIR/$APP_NAME"
APP_ZIP="$OUTPUT_DIR/Word-MCP-Bridge-Helper.zip"
TEMP_NOTARY_ZIP="$OUTPUT_DIR/Word-MCP-Bridge-Helper.notary.zip"
SETUP_DIR="$APP_BUNDLE/Contents/Resources/setup"
TASKPANE_DIR="$APP_BUNDLE/Contents/Resources/taskpane"
TASKPANE_DIST_DIR="$APP_BUNDLE/Contents/Resources/taskpane-dist"
RUNTIME_DIR="$APP_BUNDLE/Contents/Resources/runtime"
RUNTIME_BIN_DIR="$RUNTIME_DIR/bin"
RUNTIME_NODE_ROOT="$RUNTIME_DIR/node"
RUNTIME_NODE_BIN_DIR="$RUNTIME_NODE_ROOT/bin"
RUNTIME_BRIDGE_DIR="$RUNTIME_DIR/bridge"
ICONSET_DIR="$APP_DIR/AppIcon.iconset"
ICON_FILE="$APP_BUNDLE/Contents/Resources/WordMCPBridgeHelper.icns"

NODE_VERSION="${WORD_MCP_BRIDGE_NODE_VERSION:-$(node -p 'process.version')}"
NODE_PLATFORM="${WORD_MCP_BRIDGE_NODE_PLATFORM:-$(node -p 'process.platform')}"
NODE_ARCH="${WORD_MCP_BRIDGE_NODE_ARCH:-$(node -p 'process.arch')}"
NODE_DIST_FILENAME="node-${NODE_VERSION}-${NODE_PLATFORM}-${NODE_ARCH}.tar.gz"
NODE_DIST_URL="${WORD_MCP_BRIDGE_NODE_DIST_URL:-https://nodejs.org/dist/${NODE_VERSION}/${NODE_DIST_FILENAME}}"

swift build --configuration release --package-path "$APP_DIR"
python3 "$APP_DIR/Scripts/build_helper_icon.py"
rm -f "$APP_DIR/WordMCPBridgeHelper.icns"
iconutil -c icns "$ICONSET_DIR" -o "$APP_DIR/WordMCPBridgeHelper.icns"

pnpm --dir "$REPO_ROOT" --filter @word-mcp-bridge/bridge build
pnpm --dir "$REPO_ROOT" --filter @word-mcp-bridge/word-addin build

rm -rf "$APP_BUNDLE" "$APP_ZIP" "$TEMP_NOTARY_ZIP"
mkdir -p \
  "$APP_BUNDLE/Contents/MacOS" \
  "$APP_BUNDLE/Contents/Resources" \
  "$SETUP_DIR" \
  "$TASKPANE_DIR" \
  "$TASKPANE_DIST_DIR" \
  "$RUNTIME_BIN_DIR" \
  "$RUNTIME_NODE_ROOT" \
  "$RUNTIME_BRIDGE_DIR"

cat > "$APP_BUNDLE/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDevelopmentRegion</key>
  <string>en</string>
  <key>CFBundleExecutable</key>
  <string>WordMCPBridgeHelper</string>
  <key>CFBundleIdentifier</key>
  <string>${BUNDLE_ID}</string>
  <key>CFBundleInfoDictionaryVersion</key>
  <string>6.0</string>
  <key>CFBundleIconFile</key>
  <string>WordMCPBridgeHelper</string>
  <key>CFBundleName</key>
  <string>${APP_NAME%.app}</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>${APP_VERSION}</string>
  <key>CFBundleVersion</key>
  <string>${BUILD_NUMBER}</string>
  <key>LSMinimumSystemVersion</key>
  <string>${MINIMUM_MACOS_VERSION}</string>
  <key>LSUIElement</key>
  <true/>
</dict>
</plist>
PLIST

cp "$BUILD_DIR/WordMCPBridgeHelper" "$APP_BUNDLE/Contents/MacOS/WordMCPBridgeHelper"
cp "$APP_DIR/WordMCPBridgeHelper.icns" "$ICON_FILE"
cp "$REPO_ROOT/apps/mac-helper/SETUP-GUIDE.md" "$SETUP_DIR/SETUP-GUIDE.md"
cp "$REPO_ROOT/apps/mac-helper/INSTALL.md" "$INSTALL_GUIDE"
cp "$REPO_ROOT/apps/mac-helper/release-metadata.json" "$SETUP_DIR/release-metadata.json"
cp "$REPO_ROOT/packages/word-addin/manifest.prod.xml" "$SETUP_DIR/manifest.prod.xml"
cp "$REPO_ROOT/packages/word-addin/manifest.xml" "$SETUP_DIR/manifest.xml"
cp "$REPO_ROOT/apps/mac-helper/Scripts/serve_taskpane.mjs" "$TASKPANE_DIR/serve_taskpane.mjs"
cp -R "$REPO_ROOT/packages/word-addin/dist/." "$TASKPANE_DIST_DIR/"

BRIDGE_PACK_DIR="$(mktemp -d)"
NODE_DIST_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "$BRIDGE_PACK_DIR"
  rm -rf "$NODE_DIST_DIR"
}
trap cleanup EXIT

curl -fsSL "$NODE_DIST_URL" -o "$NODE_DIST_DIR/$NODE_DIST_FILENAME"
tar -xzf "$NODE_DIST_DIR/$NODE_DIST_FILENAME" -C "$RUNTIME_NODE_ROOT" --strip-components=1
chmod +x "$RUNTIME_NODE_BIN_DIR/node"

(
  cd "$REPO_ROOT/packages/bridge"
  npm pack --pack-destination "$BRIDGE_PACK_DIR" >/dev/null
)
BRIDGE_PACKAGE_TGZ="$(find "$BRIDGE_PACK_DIR" -maxdepth 1 -name '*.tgz' | head -n 1)"
if [[ -z "$BRIDGE_PACKAGE_TGZ" ]]; then
  echo "Could not create packaged bridge tarball." >&2
  exit 1
fi
npm install --omit=dev --ignore-scripts --prefix "$RUNTIME_BRIDGE_DIR" "$BRIDGE_PACKAGE_TGZ" >/dev/null

cat > "$RUNTIME_BIN_DIR/office-bridge" <<'WRAPPER'
#!/bin/sh
set -eu
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"
RUNTIME_DIR="$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)"
NODE_BIN="$RUNTIME_DIR/node/bin/node"
BRIDGE_ENTRY="$RUNTIME_DIR/bridge/node_modules/@word-mcp-bridge/bridge/bin/office-bridge.js"
exec "$NODE_BIN" "$BRIDGE_ENTRY" "$@"
WRAPPER
chmod +x "$RUNTIME_BIN_DIR/office-bridge"

swiftc "$REPO_ROOT/scripts/bridge/word_ax.swift" -o "$RUNTIME_BIN_DIR/word-mcp-bridge-word-launcher"
chmod +x "$RUNTIME_BIN_DIR/word-mcp-bridge-word-launcher"

if [[ -n "${WORD_MCP_BRIDGE_CODESIGN_IDENTITY:-}" ]]; then
  codesign --force --deep --options runtime --sign "$WORD_MCP_BRIDGE_CODESIGN_IDENTITY" "$APP_BUNDLE"

  if [[ -n "${WORD_MCP_BRIDGE_NOTARY_PROFILE:-}" ]]; then
    ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE" "$TEMP_NOTARY_ZIP"
    xcrun notarytool submit "$TEMP_NOTARY_ZIP" \
      --keychain-profile "$WORD_MCP_BRIDGE_NOTARY_PROFILE" \
      --wait
    xcrun stapler staple "$APP_BUNDLE"
    rm -f "$TEMP_NOTARY_ZIP"
  fi
fi

ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE" "$APP_ZIP"

if [[ -x "$VERIFY_SCRIPT" ]]; then
  "$VERIFY_SCRIPT" "$APP_BUNDLE" "$APP_ZIP"
fi

echo "Packaged helper app at: $APP_BUNDLE"
echo "Packaged helper zip at: $APP_ZIP"
