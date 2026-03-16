#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$SCRIPT_DIR/server/server.pid"
LOG_FILE="$SCRIPT_DIR/server/logs/server.log"

# Already running?
if [ -f "$PID_FILE" ]; then
  OLD_PID=$(cat "$PID_FILE")
  if kill -0 "$OLD_PID" 2>/dev/null; then
    echo "Server already running (PID $OLD_PID). Run stop.sh to stop it."
    exit 1
  else
    rm "$PID_FILE"
  fi
fi

cd "$SCRIPT_DIR/server"
source venv/bin/activate

nohup uvicorn main:app \
  --host 0.0.0.0 \
  --port 5000 \
  --ssl-keyfile "$SCRIPT_DIR/certs/localhost.key" \
  --ssl-certfile "$SCRIPT_DIR/certs/localhost.crt" \
  --log-level info \
  >> "$LOG_FILE" 2>&1 &

echo $! > "$PID_FILE"
echo "Server started (PID $(cat $PID_FILE))"
echo "Logs: $LOG_FILE"
echo "      tail -f $LOG_FILE"
