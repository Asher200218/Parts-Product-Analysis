#!/bin/bash
# Parts Analysis — local launcher (macOS). Double-click it in Finder to run.
# Double-click in Finder: starts a local server in THIS folder and opens your browser.
# Close this Terminal window to stop the server.

cd "$(dirname "$0")" || exit 1

PY="$(command -v python3 || command -v python)"
if [ -z "$PY" ]; then
  osascript -e 'display alert "Python 3 required" message "python3 was not found. Install Python 3, then double-click this again."' >/dev/null 2>&1
  echo "ERROR: python3 not found. Install Python 3 first."
  read -r -p "Press Enter to close… " _
  exit 1
fi

# pick a free port
PORT=4173
for p in 4173 4174 4175 4180 8099 8123; do
  if ! lsof -i ":$p" >/dev/null 2>&1; then PORT=$p; break; fi
done
URL="http://localhost:$PORT"

# open browser once server is up
( for _ in 1 2 3 4 5 6 7 8 9 10; do
    if curl -s -o /dev/null "$URL"; then open "$URL"; break; fi
    sleep 0.3
  done ) &

echo "──────────────────────────────────────────────"
echo "  Parts Analysis is running"
echo "  $URL"
echo "  Close this window to stop."
echo "──────────────────────────────────────────────"
exec "$PY" -m http.server "$PORT"
