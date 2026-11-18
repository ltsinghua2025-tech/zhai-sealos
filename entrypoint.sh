#!/bin/bash
set -euo pipefail
umask 077

project=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
runtime="$project/zhai"
cd "$project"

if [ -z "${ZHAI_ACCESS_PASSWORD:-}" ]; then
  echo "请先设置 ZHAI_ACCESS_PASSWORD，再启动应用。" >&2
  exit 1
fi
if [ ! -x "$runtime/.venv/bin/python3" ]; then
  echo "运行依赖未安装，请先使用 bash start.sh 启动。" >&2
  exit 1
fi

export ZHAI_CONFIG_FILE="${ZHAI_CONFIG_FILE:-$runtime/private-config.json}"
export ZHAI_DATA_DIR="${ZHAI_DATA_DIR:-$runtime/data}"
export ZHAI_WEB_HOST=0.0.0.0
export ZHAI_WEB_PORT="${ZHAI_WEB_PORT:-${PORT:-8080}}"
export ZHAI_ACCESS_USER="${ZHAI_ACCESS_USER:-zhai}"
export PYTHONUNBUFFERED=1
case "$(uname -m)" in
  aarch64|arm64) export OPENSSL_armcap="${OPENSSL_armcap:-0}" ;;
esac

"$runtime/.venv/bin/python3" - "$ZHAI_WEB_PORT" <<'PY'
import sys
port = sys.argv[1]
if not port.isascii() or not port.isdecimal() or not 1 <= int(port) <= 65535:
    raise SystemExit('ZHAI_WEB_PORT 或 PORT 必须是 1 至 65535 的端口号。')
PY
"$runtime/.venv/bin/python3" "$runtime/verify-package.py"
cd "$runtime"
exec "$runtime/.venv/bin/python3" server/app.py
