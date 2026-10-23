#!/bin/bash
# Parts Analysis — 本地启动器 (macOS)。在「访达」里双击即可运行。
# Double-click in Finder: starts a local server in THIS folder and opens your browser.
# 关闭这个终端窗口即可停止服务器 / Close this Terminal window to stop the server.

cd "$(dirname "$0")" || exit 1

PY="$(command -v python3 || command -v python)"
if [ -z "$PY" ]; then
  osascript -e 'display alert "需要 Python 3 / Python 3 required" message "未找到 python3。请先安装 Python 3 后再双击运行。"' >/dev/null 2>&1
  echo "ERROR: python3 not found. Install Python 3 first."
  read -r -p "按回车关闭… / Press Enter to close… " _
  exit 1
fi

# 找一个空闲端口 / pick a free port
PORT=4173
for p in 4173 4174 4175 4180 8099 8123; do
  if ! lsof -i ":$p" >/dev/null 2>&1; then PORT=$p; break; fi
done
URL="http://localhost:$PORT"

# 等服务器起来后自动打开浏览器 / open browser once server is up
( for _ in 1 2 3 4 5 6 7 8 9 10; do
    if curl -s -o /dev/null "$URL"; then open "$URL"; break; fi
    sleep 0.3
  done ) &

echo "──────────────────────────────────────────────"
echo "  零件分析录入网页正在运行 / Parts Analysis is running"
echo "  $URL"
echo "  关闭此窗口即可停止 / Close this window to stop."
echo "──────────────────────────────────────────────"
exec "$PY" -m http.server "$PORT"
