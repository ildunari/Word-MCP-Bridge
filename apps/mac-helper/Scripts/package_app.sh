#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_ROOT="$(cd "$APP_DIR/../.." && pwd)"
BUILD_DIR="$APP_DIR/.build/release"
OUTPUT_DIR="$REPO_ROOT/release/mac-helper"
APP_NAME="Word MCP Bridge Helper.app"
APP_BUNDLE="$OUTPUT_DIR/$APP_NAME"
APP_ZIP="$OUTPUT_DIR/Word-MCP-Bridge-Helper.zip"
INSTALL_GUIDE="$REPO_ROOT/release/INSTALL.md"
SETUP_DIR="$APP_BUNDLE/Contents/Resources/setup"
TASKPANE_DIR="$APP_BUNDLE/Contents/Resources/taskpane"
TASKPANE_DIST_DIR="$APP_BUNDLE/Contents/Resources/taskpane-dist"
LAUNCHER_DIR="$APP_BUNDLE/Contents/Resources/bridge-launch/scripts/bridge"
ICONSET_DIR="$APP_DIR/AppIcon.iconset"
ICON_FILE="$APP_BUNDLE/Contents/Resources/WordMCPBridgeHelper.icns"

swift build --configuration release --package-path "$APP_DIR"
python3 "$APP_DIR/Scripts/build_helper_icon.py"
rm -f "$APP_DIR/WordMCPBridgeHelper.icns"
iconutil -c icns "$ICONSET_DIR" -o "$APP_DIR/WordMCPBridgeHelper.icns"

pnpm --dir "$REPO_ROOT" --filter @word-mcp-bridge/word-addin build

rm -rf "$APP_BUNDLE" "$APP_ZIP"
mkdir -p "$APP_BUNDLE/Contents/MacOS" "$APP_BUNDLE/Contents/Resources" "$SETUP_DIR" "$TASKPANE_DIR" "$TASKPANE_DIST_DIR" "$LAUNCHER_DIR"

cat > "$APP_BUNDLE/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDevelopmentRegion</key>
  <string>en</string>
  <key>CFBundleExecutable</key>
  <string>WordMCPBridgeHelper</string>
  <key>CFBundleIdentifier</key>
  <string>dev.wordmcpbridge.helper</string>
  <key>CFBundleInfoDictionaryVersion</key>
  <string>6.0</string>
  <key>CFBundleIconFile</key>
  <string>WordMCPBridgeHelper</string>
  <key>CFBundleName</key>
  <string>Word MCP Bridge Helper</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>1.0.0</string>
  <key>CFBundleVersion</key>
  <string>1</string>
  <key>LSMinimumSystemVersion</key>
  <string>14.0</string>
  <key>LSUIElement</key>
  <true/>
</dict>
</plist>
PLIST

cp "$BUILD_DIR/WordMCPBridgeHelper" "$APP_BUNDLE/Contents/MacOS/WordMCPBridgeHelper"
cp "$APP_DIR/WordMCPBridgeHelper.icns" "$ICON_FILE"
cp "$REPO_ROOT/apps/mac-helper/SETUP-GUIDE.md" "$SETUP_DIR/SETUP-GUIDE.md"
cp "$REPO_ROOT/apps/mac-helper/INSTALL.md" "$INSTALL_GUIDE"
cp "$REPO_ROOT/packages/word-addin/manifest.prod.xml" "$SETUP_DIR/manifest.prod.xml"
cp "$REPO_ROOT/packages/word-addin/manifest.xml" "$SETUP_DIR/manifest.xml"
cp "$REPO_ROOT/apps/mac-helper/Scripts/serve_taskpane.py" "$TASKPANE_DIR/serve_taskpane.py"
cp -R "$REPO_ROOT/packages/word-addin/dist/." "$TASKPANE_DIST_DIR/"
cp "$REPO_ROOT/scripts/bridge/launch-word-taskpane.sh" "$LAUNCHER_DIR/launch-word-taskpane.sh"
cp "$REPO_ROOT/scripts/bridge/word_ax.swift" "$LAUNCHER_DIR/word_ax.swift"
chmod +x "$LAUNCHER_DIR/launch-word-taskpane.sh"

if [[ -n "${WORD_MCP_BRIDGE_CODESIGN_IDENTITY:-}" ]]; then
  codesign --force --deep --options runtime --sign "$WORD_MCP_BRIDGE_CODESIGN_IDENTITY" "$APP_BUNDLE"
fi

ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE" "$APP_ZIP"

echo "Packaged helper app at: $APP_BUNDLE"
echo "Packaged helper zip at: $APP_ZIP"
