#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$SCRIPT_DIR/.."

echo "=== Claude Word Add-in Setup ==="

# 1. Python venv
echo ""
echo "--- Setting up Python virtual environment ---"
cd "$ROOT/server"
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip --quiet
pip install -r requirements.txt
echo "✓ Python deps installed"

# 2. SSL certs
echo ""
echo "--- Generating self-signed SSL certificate ---"
mkdir -p "$ROOT/certs"
openssl req -x509 \
  -newkey rsa:4096 \
  -keyout "$ROOT/certs/localhost.key" \
  -out "$ROOT/certs/localhost.crt" \
  -days 365 \
  -nodes \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
  2>/dev/null
echo "✓ Certificates generated in certs/"

# 3. Trust the cert on macOS
echo ""
echo "--- Trusting certificate in macOS keychain (may prompt for password) ---"
security add-trusted-cert \
  -d -r trustRoot \
  -k ~/Library/Keychains/login.keychain-db \
  "$ROOT/certs/localhost.crt" && echo "✓ Certificate trusted" \
  || echo "⚠ Could not add to keychain automatically — trust it manually in Keychain Access"

# 4. Word sideload directory
echo ""
echo "--- Creating Word add-in sideload directory ---"
WEF_DIR=~/Library/Containers/com.microsoft.Word/Data/Documents/wef
mkdir -p "$WEF_DIR"
cp "$ROOT/addin/manifest.xml" "$WEF_DIR/claude-word-manifest.xml"
echo "✓ Manifest copied to: $WEF_DIR"

echo ""
echo "=== Setup complete ==="
echo ""
echo "Next steps:"
echo "  1. Run: bash start.sh"
echo "  2. Open Word → Insert → Add-ins → My Add-ins → Claude Assistant"
echo "  3. Select some text in your document and ask Claude a question"
echo ""
echo "Logs will be written to: server/logs/server.log"
