#!/bin/bash
# Idempotently starts the two background services AI Assistant needs on macOS:
# (opencode serve + webpack dev server).

# Resolve the root directory (parent of scripts directory)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
DEV_PORT=3000
OPENCODE_PORT=4098

LOG_DIR="$HOME/Library/Logs/AIAssistant"
mkdir -p "$LOG_DIR"

test_port() {
  nc -z 127.0.0.1 "$1" >/dev/null 2>&1
  return $?
}

cd "$ROOT_DIR" || exit 1

if test_port "$DEV_PORT"; then
  echo "webpack dev server already listening on port $DEV_PORT - skipping"
else
  echo "Starting webpack dev server on port $DEV_PORT..."
  nohup npm run dev-server > "$LOG_DIR/dev-server.log" 2> "$LOG_DIR/dev-server.err.log" &
fi

if test_port "$OPENCODE_PORT"; then
  echo "opencode serve already listening on port $OPENCODE_PORT - skipping"
else
  echo "Starting opencode serve on port $OPENCODE_PORT..."
  nohup opencode serve --port "$OPENCODE_PORT" --hostname 127.0.0.1 --cors "https://localhost:$DEV_PORT" > "$LOG_DIR/opencode-serve.log" 2> "$LOG_DIR/opencode-serve.err.log" &
fi
