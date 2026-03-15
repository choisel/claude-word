#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# build-app.sh — Build "Claude Word Assistant.app" for end-user distribution
#
# Output: dist/Claude Word Assistant.app
#
# Requirements (dev machine only):
#   - Python 3.8+
#   - Internet access (to install build dependencies in a temp venv)
# ---------------------------------------------------------------------------
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$SCRIPT_DIR/.."
BUILD_VENV="$ROOT/.build-venv"
DIST_DIR="$ROOT/dist"
APP_NAME="Claude Word Assistant"

echo "=== Building $APP_NAME.app ==="
echo ""

# 1. Build venv with all dependencies + build tools
echo "--- Setting up build environment ---"
python3 -m venv "$BUILD_VENV"
source "$BUILD_VENV/bin/activate"
pip install --upgrade pip --quiet
pip install \
    fastapi==0.115.0 \
    "uvicorn[standard]==0.34.0" \
    python-dotenv==1.0.1 \
    sse-starlette==2.1.3 \
    rumps \
    pyinstaller \
    --quiet
echo "✓ Build dependencies installed"

# 2. Generate icon (simple PNG → icns via sips)
echo ""
echo "--- Generating app icon ---"
ICON_DIR="$ROOT/.build-icon"
ICONSET_DIR="$ICON_DIR/AppIcon.iconset"
mkdir -p "$ICONSET_DIR"
SRC_ICON="$ROOT/addin/assets/icon-80.png"

for size in 16 32 64 128 256 512; do
    sips -z $size $size "$SRC_ICON" --out "$ICONSET_DIR/icon_${size}x${size}.png" --silent
    double=$((size * 2))
    sips -z $double $double "$SRC_ICON" --out "$ICONSET_DIR/icon_${size}x${size}@2x.png" --silent
done
iconutil -c icns "$ICONSET_DIR" -o "$ICON_DIR/AppIcon.icns"
echo "✓ Icon generated"

# 3. PyInstaller spec — packages server + addin together
echo ""
echo "--- Building binary with PyInstaller ---"
cd "$ROOT"

pyinstaller \
    --noconfirm \
    --onedir \
    --windowed \
    --name "$APP_NAME" \
    --icon "$ICON_DIR/AppIcon.icns" \
    --distpath "$DIST_DIR" \
    --workpath "$ROOT/.build-work" \
    --specpath "$ROOT/.build-spec" \
    --add-data "addin:addin" \
    --add-data "server/main.py:." \
    --add-data "server/document.py:." \
    --add-data "server/session.py:." \
    --hidden-import "uvicorn.logging" \
    --hidden-import "uvicorn.loops" \
    --hidden-import "uvicorn.loops.auto" \
    --hidden-import "uvicorn.protocols" \
    --hidden-import "uvicorn.protocols.http" \
    --hidden-import "uvicorn.protocols.http.auto" \
    --hidden-import "uvicorn.protocols.websockets" \
    --hidden-import "uvicorn.protocols.websockets.auto" \
    --hidden-import "uvicorn.lifespan" \
    --hidden-import "uvicorn.lifespan.on" \
    --hidden-import "fastapi" \
    --hidden-import "sse_starlette" \
    server/app_launcher.py \
    2>&1 | tail -20

echo "✓ Binary built"

# 4. Cleanup temp files
deactivate
rm -rf "$BUILD_VENV" "$ICON_DIR" "$ROOT/.build-work" "$ROOT/.build-spec"

APP_PATH="$DIST_DIR/$APP_NAME.app"
echo ""
echo "=== Build complete ==="
echo ""
echo "App: $APP_PATH"
echo "Size: $(du -sh "$APP_PATH" | cut -f1)"
echo ""
echo "To distribute:"
echo "  1. Zip: cd dist && zip -r '$APP_NAME.zip' '$APP_NAME.app'"
echo "  2. Upload the .zip to a GitHub Release"
echo "  3. Users download, unzip, double-click — done."
