#!/bin/bash
set -euo pipefail
umask 077

project=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
runtime="$project/zhai"
cd "$project"

if [ "$(uname -s)" != Linux ]; then
  echo "请在 Sealos 的 Linux Python 3.12 环境中运行。" >&2
  exit 1
fi
python3 -c 'import sys; assert sys.version_info[:2] == (3, 12), "需要 Python 3.12"'
case "$(uname -m)" in
  x86_64|amd64) arch=x86_64 ;;
  aarch64|arm64) arch=aarch64; export OPENSSL_armcap="${OPENSSL_armcap:-0}" ;;
  *) echo "支持的服务器架构为 x86_64 或 ARM64。" >&2; exit 1 ;;
esac

python3 "$runtime/verify-package.py"
fingerprint=$(python3 - "$runtime" "$arch" <<'PY'
import hashlib
from pathlib import Path
import sys

root = Path(sys.argv[1])
architecture = sys.argv[2]
wheels = sorted((root / 'wheels' / architecture).glob('*.whl'))
if not wheels:
    raise SystemExit('缺少当前服务器架构的随包依赖，请使用完整的 Git 导入包。')
digest = hashlib.sha256()
digest.update((sys.version + '\n' + architecture).encode())
digest.update((root / 'engine' / 'requirements.txt').read_bytes())
for wheel in wheels:
    digest.update(wheel.name.encode())
    with wheel.open('rb') as source:
        digest.update(hashlib.file_digest(source, 'sha256').digest())
print(digest.hexdigest())
PY
)

python3_env="$runtime/.venv/bin/python3"
stamp="$runtime/.venv/.zhai-dependencies.sha256"
if [ ! -x "$python3_env" ]; then
  python3 -m venv "$runtime/.venv"
fi

check_dependencies() {
  "$python3_env" - "$runtime/engine/requirements.txt" <<'PY'
from importlib.metadata import version
from pathlib import Path
import re
import sys

assert sys.version_info[:2] == (3, 12), '虚拟环境需要 Python 3.12'
for line in Path(sys.argv[1]).read_text().splitlines():
    line = line.strip()
    if not line or line.startswith('#'):
        continue
    requirement = re.fullmatch(r'([A-Za-z0-9_.-]+)==([^\s;]+)', line)
    if requirement is None:
        raise SystemExit('随包依赖须使用明确的版本号。')
    name, expected = requirement.groups()
    assert version(name) == expected, f'运行依赖版本不符：{name}'
import docx, docxcompose, openpyxl, pdfplumber, pypdf, reportlab, lxml
PY
}

if [ "$(cat "$stamp" 2>/dev/null || true)" != "$fingerprint" ] || ! check_dependencies >/dev/null 2>&1; then
  echo "正在从随包 wheels 安装 Python 依赖……"
  "$python3_env" -m pip install --disable-pip-version-check --no-index --find-links="$runtime/wheels/$arch" -r "$runtime/engine/requirements.txt"
  check_dependencies
  printf '%s\n' "$fingerprint" > "$stamp.tmp"
  mv "$stamp.tmp" "$stamp"
else
  echo "随包运行依赖已就绪。"
fi

export ZHAI_CONFIG_FILE="${ZHAI_CONFIG_FILE:-$runtime/private-config.json}"
export ZHAI_DATA_DIR="${ZHAI_DATA_DIR:-$runtime/data}"
export ZHAI_WEB_HOST=0.0.0.0
export ZHAI_WEB_PORT="${ZHAI_WEB_PORT:-${PORT:-8080}}"
export PYTHONUNBUFFERED=1

"$python3_env" - "$ZHAI_WEB_PORT" <<'PY'
import sys
port = sys.argv[1]
if not port.isascii() or not port.isdecimal() or not 1 <= int(port) <= 65535:
    raise SystemExit('ZHAI_WEB_PORT 或 PORT 必须是 1 至 65535 的端口号。')
PY
cd "$runtime"
exec "$python3_env" server/app.py
