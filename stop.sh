#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$SCRIPT_DIR/server/server.pid"

if [ ! -f "$PID_FILE" ]; then
  echo "Aucun serveur en cours (pas de PID file)."
  exit 0
fi

PID=$(cat "$PID_FILE")

if kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  rm "$PID_FILE"
  echo "Serveur arrêté (PID $PID)."
else
  echo "Processus $PID introuvable (déjà arrêté ?)."
  rm "$PID_FILE"
fi
