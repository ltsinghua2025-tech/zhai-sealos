#!/usr/bin/env python3
"""Run inside the deployed environment; never print credentials."""
import base64
import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

root = Path(__file__).resolve().parent
sys.path.insert(0, str(root / "server"))
from runtime_config import load_config
load_config(root)
parser = argparse.ArgumentParser()
parser.add_argument('--model', action='store_true', help='另外请求默认算力，消耗少量模型用量')
args = parser.parse_args()
port = os.environ.get("ZHAI_WEB_PORT", "8080")
credential = f"{os.environ.get('ZHAI_ACCESS_USER','zhai')}:{os.environ.get('ZHAI_ACCESS_PASSWORD','')}"
headers = {"Authorization": "Basic " + base64.b64encode(credential.encode()).decode(), "Content-Type": "application/json"}
try:
    with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/health", timeout=5) as response:
        health = json.load(response)
        assert health["ok"]
    print("网页服务：正常；版本：" + str(health.get('version', '旧版')))
    if not args.model:
        print('默认算力未调用；如需检查，追加 --model')
        raise SystemExit(0)
    request = urllib.request.Request(f"http://127.0.0.1:{port}/api/profiles/test", data=b"{}", headers=headers)
    with urllib.request.urlopen(request, timeout=150) as response:
        result = json.load(response)
    print("默认算力：" + ("连接成功" if result.get("success") else "未通过"))
except urllib.error.HTTPError as exc:
    print(f"检查未通过：HTTP {exc.code}。请查看平台日志或联系部署协助人员。")
    raise SystemExit(1)
except Exception:
    print("检查未通过：服务未启动、端口不一致或网络无法连接算力。")
    raise SystemExit(1)
