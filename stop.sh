#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$SCRIPT_DIR/server/server.pid"

if [ ! -f "$PID_FILE" ]; then
  echo "No server running (no PID file found)."
  exit 0
fi

PID=$(cat "$PID_FILE")

if kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  rm "$PID_FILE"
  echo "Server stopped (PID $PID)."
else
  echo "Process $PID not found (already stopped?)."
  rm "$PID_FILE"
fi
