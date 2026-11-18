#!/usr/bin/env python3
"""Verify public runtime files; secret configuration is deliberately excluded."""
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
for relative, expected in manifest.items():
    path = root / relative
    if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
        raise SystemExit(f"文件不完整：{relative}。请重新从 Git 导入。")
print(f"部署文件完整性检查通过，共 {len(manifest)} 项")
