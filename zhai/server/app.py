#!/usr/bin/env python3
"""zhai Web API.

The desktop build keeps its Tauri/Rust harness.  This small HTTP service is the
web adapter: it owns uploaded files, sessions and the existing Python engine.
It intentionally keeps the API surface narrow and explicit so the browser never
receives a server filesystem path as an operation target.
"""

from __future__ import annotations

import base64
import html
import hashlib
import hmac
import ipaddress
import io
import json
import mimetypes
import os
import re
import sqlite3
import sys
import time
import socket
import threading
import urllib.error
import urllib.request
import uuid
import zipfile
from contextlib import redirect_stdout
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "engine"))
sys.path.insert(0, str(ROOT / "server"))
from runtime_config import load_config, model_settings, safe_error
from data_admin import data_lock
load_config(ROOT)
import material_workshop_engine as engine  # noqa: E402


DATA_DIR = Path(os.environ.get("ZHAI_DATA_DIR", "/data/zhai")).resolve()
UPLOAD_DIR = DATA_DIR / "uploads"
ARTIFACT_DIR = DATA_DIR / "artifacts"
DB_PATH = DATA_DIR / "zhai.sqlite3"
MAX_UPLOAD_BYTES = int(os.environ.get("ZHAI_MAX_UPLOAD_BYTES", str(50 * 1024 * 1024)))
ENGINE_LOCK = threading.Lock()
TURN_LOCK = threading.RLock()
TURNS: dict[str, dict[str, Any]] = {}
TURN_LOCAL = threading.local()
WEB_VERSION = "0.27.1"
BUILTIN_OUTPUT_CONTRACTS = json.loads((ROOT / "resources/builtin-output-contracts.json").read_text(encoding="utf-8"))


def check_cancelled() -> None:
    cancelled = getattr(TURN_LOCAL, "cancelled", None)
    if cancelled and cancelled.is_set():
        raise RuntimeError("任务已停止；已结束的模型请求仍可能产生计费")


def file_hash(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


def json_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def safe_name(value: str) -> str:
    name = Path(value).name.strip() or "未命名文件"
    name = re.sub(r"[^\w.()\-\u4e00-\u9fff ]+", "_", name)
    return name[:160]


def init_db() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS sessions (
              id TEXT PRIMARY KEY, title TEXT NOT NULL, profile_id TEXT NOT NULL,
              model TEXT NOT NULL, status TEXT NOT NULL, created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS messages (
              id TEXT PRIMARY KEY, session_id TEXT NOT NULL, role TEXT NOT NULL,
              content TEXT NOT NULL, created_at TEXT NOT NULL, turn_id TEXT,
              FOREIGN KEY(session_id) REFERENCES sessions(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS attachments (
              id TEXT PRIMARY KEY, session_id TEXT NOT NULL, kind TEXT NOT NULL,
              name TEXT NOT NULL, server_path TEXT, source_url TEXT, mime_type TEXT,
              size INTEGER, relative_path TEXT, created_at TEXT NOT NULL,
              source_group_id TEXT, source_group_name TEXT,
              FOREIGN KEY(session_id) REFERENCES sessions(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS skills (
              id TEXT PRIMARY KEY, payload TEXT NOT NULL, deleted INTEGER NOT NULL DEFAULT 0,
              created_at TEXT NOT NULL, updated_at TEXT NOT NULL
            );
            """
        )


def row_session(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "title": row["title"],
        "profileId": row["profile_id"],
        "model": row["model"],
        "status": row["status"],
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "workspacePath": "web://zhai/sessions/" + row["id"],
    }


def attachment_json(row: sqlite3.Row, artifact: bool = False) -> dict[str, Any]:
    return {
        "id": row["id"],
        "sessionId": row["session_id"],
        "kind": "artifact" if artifact else row["kind"],
        "name": row["name"],
        "mimeType": row["mime_type"],
        "localSnapshotPath": row["server_path"],
        "sourceUrl": row["source_url"],
        "size": row["size"],
        "createdAt": row["created_at"],
        "sourceGroupId": row["source_group_id"],
        "sourceGroupName": row["source_group_name"],
        "relativePath": row["relative_path"] or row["name"],
    }


def fetch_public_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("只允许读取 HTTP(S) 公开网址")
    host = parsed.hostname
    try:
        addresses = {item[4][0] for item in socket.getaddrinfo(host, parsed.port or (443 if parsed.scheme == "https" else 80), type=socket.SOCK_STREAM)}
    except socket.gaierror as exc:
        raise ValueError("网址域名无法解析") from exc
    if any(not ipaddress.ip_address(address).is_global for address in addresses):
        raise ValueError("为防止 SSRF，不能读取本机、局域网或链路本地地址")
    request = urllib.request.Request(url, headers={"User-Agent": "ZHAI-Web-Agent/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            raw = response.read(MAX_UPLOAD_BYTES + 1)
            content_type = response.headers.get_content_type()
            charset = response.headers.get_content_charset() or "utf-8"
    except (urllib.error.URLError, TimeoutError) as exc:
        raise ValueError(f"公开网址读取失败：{exc}") from exc
    if len(raw) > MAX_UPLOAD_BYTES:
        raise ValueError("公开网址内容超过大小限制")
    text = raw.decode(charset, errors="replace")
    if content_type in {"text/html", "application/xhtml+xml"}:
        text = re.sub(r"(?is)<(script|style|noscript).*?>.*?</\1>", " ", text)
        text = re.sub(r"(?s)<[^>]+>", " ", html.unescape(text))
    return re.sub(r"\s+", " ", text).strip()[:20_000]


def get_session(session_id: str) -> dict[str, Any] | None:
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        row = db.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        if not row:
            return None
        messages = [
            dict(message)
            for message in db.execute(
                "SELECT id, session_id, role, content, created_at, turn_id FROM messages WHERE session_id = ? ORDER BY created_at",
                (session_id,),
            ).fetchall()
        ]
        messages = [
            {
                "id": item["id"],
                "sessionId": item["session_id"],
                "role": item["role"],
                "content": item["content"],
                "createdAt": item["created_at"],
                "turnId": item["turn_id"],
            }
            for item in messages
        ]
        attachments = [
            attachment_json(item)
            for item in db.execute(
                "SELECT * FROM attachments WHERE session_id = ? ORDER BY created_at",
                (session_id,),
            ).fetchall()
        ]
        return {"session": row_session(row), "messages": messages, "attachments": attachments, "toolRuns": []}


def file_context(session_id: str, limit: int = 40000) -> str:
    chunks: list[str] = []
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        rows = db.execute(
            "SELECT name, server_path FROM attachments WHERE session_id = ? AND server_path IS NOT NULL AND kind != 'artifact' ORDER BY created_at",
            (session_id,),
        ).fetchall()
    remaining = limit
    for row in rows:
        path = Path(row["server_path"])
        if not path.is_file() or remaining <= 0:
            continue
        try:
            text = engine.extract_text_preview(path)
        except Exception as exc:  # noqa: BLE001 - one bad file must not hide the rest
            text = f"[无法读取：{exc}]"
        text = text[:remaining]
        chunks.append(f"【{row['name']}】\n{text}")
        remaining -= len(text)
    return "\n\n".join(chunks)


def call_openai_compatible_payload(
    base_url: str,
    api_key: str,
    model: str,
    messages: list[dict[str, Any]],
    tools: list[dict[str, Any]] | None = None,
    tool_choice: str | None = None,
    response_format: dict[str, Any] | None = None,
) -> dict[str, Any]:
    url = base_url.rstrip("/")
    if not url.endswith("/chat/completions"):
        url += "/chat/completions"
    request_payload: dict[str, Any] = {"model": model, "messages": messages, "temperature": 0.2}
    if tools:
        request_payload["tools"] = tools
        request_payload["tool_choice"] = tool_choice or "auto"
    if response_format:
        request_payload["response_format"] = response_format
    body = json_bytes(request_payload)
    request = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, headers, newurl):
            return None

    check_cancelled()
    try:
        with urllib.request.build_opener(NoRedirect()).open(request, timeout=120) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        messages = {401: "算力密钥未通过认证", 403: "算力账号没有该模型权限", 429: "算力繁忙或额度不足"}
        raise RuntimeError(f"模型服务返回 HTTP {exc.code}：{messages.get(exc.code, '请检查接口地址、模型权限或稍后重试')}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"无法连接模型服务：{exc.reason}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("模型服务返回格式无法识别")
    check_cancelled()
    return payload


def call_openai_compatible(base_url: str, api_key: str, model: str, messages: list[dict[str, Any]]) -> str:
    payload = call_openai_compatible_payload(base_url, api_key, model, messages)
    try:
        content = payload["choices"][0]["message"].get("content", "")
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("模型服务返回格式无法识别") from exc
    return str(content or "")


def profile_list() -> list[dict[str, Any]]:
    base_url = os.environ.get("ZHAI_LLM_BASE_URL", "https://api.deepseek.com/v1")
    model = os.environ.get("ZHAI_LLM_MODEL", "")
    return [{
        "id": "default",
        "provider": "openai_compatible",
        "displayName": os.environ.get("ZHAI_LLM_NAME", "自配模型"),
        "baseUrl": base_url,
        "defaultModel": model,
        "keyStored": bool(os.environ.get("ZHAI_LLM_API_KEY")),
    }]


def builtin_tools() -> list[dict[str, Any]]:
    def tool(
        tool_id: str,
        name: str,
        description: str,
        extensions: list[str],
        available: bool = True,
        parameters: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        return {
            "id": tool_id,
            "name": name,
            "description": description,
            "source": "builtin",
            "version": WEB_VERSION,
            "permissions": ["read_selected_files", "write_task_output"],
            "supportedExtensions": extensions,
            "outputFormats": BUILTIN_OUTPUT_CONTRACTS[tool_id],
            "available": available,
            "availabilityDetail": "服务器材料处理引擎可用" if available else "服务器环境未配置该转换组件",
            "parameters": parameters or [],
            "steps": [{"toolId": tool_id, "parameters": {}}],
        }
    return [
        tool("docx_to_gongwen", "现有 Word 一键公文排版", "按保守公文规则生成DOCX副本；修正附件、缩进，保留表格与已有页脚。复杂结构保留并告知。旧DOC请先用WPS另存为DOCX。", ["docx"], parameters=[
            {"name": "attachmentIds", "type": "array", "description": "要排版的DOCX附件ID；不填则处理当前会话中的Word"},
            {"name": "preserveBold", "type": "boolean", "description": "保留原文加粗，默认true"},
            {"name": "formatTables", "type": "boolean", "description": "默认false保留表格；true仅调整安全区域字体，不改变几何结构"},
            {"name": "applyPageNumbers", "type": "boolean", "description": "添加页码但保护已有页脚，默认true"},
            {"name": "lineSpacingPt", "type": "number", "description": "正文行距，默认28磅"},
            {"name": "recipientMode", "type": "string", "description": "auto自动识别主送；none明确无主送"},
            {"name": "signatureMode", "type": "string", "description": "默认preserve保留落款；仅用户明确无章单一机关时可用unsigned"},
        ]),
        tool("normalize_docx_styles", "Word 样式规范化", "清理 Word 样式异常并生成独立 DOCX 副本；旧DOC请先用WPS另存为DOCX。", ["docx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的 Word 附件 ID"}, {"name": "preset", "type": "string", "description": "可选 gongwen 公文版式"}]),
        tool("redact_content", "敏感信息脱敏", "识别常见姓名、手机号、身份证号、地址、邮箱、统一社会信用代码等敏感信息，生成独立副本。", ["docx", "xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要脱敏的附件 ID；不填则处理当前会话全部可处理材料"}]),
        tool("normalize_whitespace", "文字清理", "清理多余空格和手动换行，保留文件结构并生成新版本。", ["docx", "xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要清理的附件 ID"}]),
        tool("remove_line_breaks", "清除手动换行", "清除段内多余手动换行，生成新版本文件。", ["docx", "xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的附件 ID"}, {"name": "joinParagraphs", "type": "boolean", "description": "是否尝试合并相邻段落"}]),
        tool("replace_text", "批量文字替换", "按明确的替换对批量修改材料，原文件不覆盖。", ["docx", "xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的附件 ID"}, {"name": "replacements", "type": "array", "description": "替换对数组，每项包含 before 和 after"}]),
        tool("delete_blank_rows", "清理 Excel 空行", "删除 Excel 工作表中的空白行并生成新文件。", ["xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的 Excel 附件 ID"}]),
        tool("set_xlsx_cell", "填写 Excel 单元格", "按用户明确指定的工作表和单元格写入值，生成新文件。", ["xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的 Excel 附件 ID"}, {"name": "sheet", "type": "string", "description": "工作表名"}, {"name": "cell", "type": "string", "description": "单元格地址"}, {"name": "value", "type": "string", "description": "写入值"}]),
        tool("extract_pdf_text", "提取 PDF 文本", "按页导出 PDF 文本层为 TXT，便于继续审阅和整理。", ["pdf"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要处理的 PDF 附件 ID"}]),
        tool("merge_xlsx", "按表头汇总 Excel", "按字段名对齐多个单行表头 XLSX，保留来源；复杂表头和公式明确拒绝。", ["xlsx"], parameters=[{"name": "attachmentIds", "type": "array", "description": "要汇总的 Excel 附件 ID；省略时使用全部 XLSX"}]),
        tool("word_to_pdf", "Word 转 PDF", "当前网页容器未配置 WPS/LibreOffice。", ["docx"], False),
    ]


def skill_defaults(tool: dict[str, Any]) -> dict[str, Any]:
    """Fill the fields the React skill library expects from a stored draft."""
    value = dict(tool)
    value.setdefault("source", "designed")
    value.setdefault("kind", "workflow")
    value.setdefault("version", "0.1.0")
    value.setdefault("author", "网页用户")
    value.setdefault("permissions", ["read_selected_files", "write_task_output"])
    value.setdefault("testStatus", "untested")
    value.setdefault("status", "draft")
    value.setdefault("supportedExtensions", ["docx", "xlsx", "pdf", "csv", "md", "txt"])
    value.setdefault("available", False)
    value.setdefault("availabilityDetail", "网页技能草稿；请先上传样本测试")
    value.setdefault("parameters", [])
    value.setdefault("steps", [])
    return value


def stored_skills(include_deleted: bool = False) -> list[dict[str, Any]]:
    with sqlite3.connect(DB_PATH) as db:
        rows = db.execute(
            "SELECT payload FROM skills WHERE deleted = 0 OR ? = 1 ORDER BY updated_at DESC",
            (1 if include_deleted else 0,),
        ).fetchall()
    return [json.loads(row[0]) for row in rows]


def deleted_skills() -> list[dict[str, Any]]:
    with sqlite3.connect(DB_PATH) as db:
        rows = db.execute("SELECT payload FROM skills WHERE deleted = 1 ORDER BY updated_at DESC").fetchall()
    return [json.loads(row[0]) for row in rows]


def all_skills(include_deleted: bool = False) -> list[dict[str, Any]]:
    return builtin_tools() + stored_skills(include_deleted=include_deleted)


def find_skill(skill_id: str, include_deleted: bool = False) -> dict[str, Any] | None:
    for tool in all_skills(include_deleted=include_deleted):
        if tool.get("id") == skill_id:
            return tool
    return None


def save_skill(tool: dict[str, Any], deleted: bool = False) -> dict[str, Any]:
    value = skill_defaults(tool)
    if value.get("source") != "designed":
        raise ValueError("内置能力不能覆盖保存")
    value["id"] = str(value.get("id") or uuid.uuid4())
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,100}", value["id"]):
        raise ValueError("能力包标识无效")
    timestamp = now()
    with sqlite3.connect(DB_PATH) as db:
        exists = db.execute("SELECT 1 FROM skills WHERE id = ?", (value["id"],)).fetchone()
        if exists:
            db.execute(
                "UPDATE skills SET payload = ?, deleted = ?, updated_at = ? WHERE id = ?",
                (json.dumps(value, ensure_ascii=False), 1 if deleted else 0, timestamp, value["id"]),
            )
        else:
            db.execute(
                "INSERT INTO skills VALUES (?, ?, ?, ?, ?)",
                (value["id"], json.dumps(value, ensure_ascii=False), 1 if deleted else 0, timestamp, timestamp),
            )
        db.commit()
    return value


def infer_skill_draft(prompt: str, kind: str = "workflow") -> dict[str, Any]:
    text = prompt.strip() or "材料处理助手"
    lower = text.lower()
    extensions = ["docx", "xlsx", "pdf", "csv"]
    if "word" in lower or "文稿" in text or "公文" in text:
        extensions = ["docx", "md", "txt"]
    elif "excel" in lower or "表格" in text:
        extensions = ["xlsx", "csv"]
    elif "pdf" in lower:
        extensions = ["pdf"]
    steps: list[dict[str, Any]] = []
    if "公文" in text or "套版" in text or "排版" in text:
        steps.append({"toolId": "docx_to_gongwen", "parameters": {"preserveBold": True, "formatTables": False, "signatureMode": "preserve", "recipientMode": "auto", "applyPageNumbers": True, "lineSpacingPt": 28}})
    if "脱敏" in text or "敏感" in text:
        steps.append({"toolId": "redact_content", "parameters": {}})
    if "检查" in text or "审查" in text or "规范" in text:
        steps.append({"toolId": "inspect_document", "parameters": {}})
    if "提取" in text and "pdf" in lower:
        steps.append({"toolId": "extract_pdf_text", "parameters": {}})
    if "csv" in lower and ("excel" in lower or "xlsx" in lower):
        steps.append({"toolId": "csv_to_xlsx", "parameters": {}})
    if "xlsx" in lower and "csv" in lower:
        steps.append({"toolId": "xlsx_to_csv", "parameters": {}})
    if not steps:
        steps = [{"toolId": "inspect_document", "parameters": {}}]
    name = text.split("：", 1)[0].split(":", 1)[0].strip()[:32] or "网页材料处理助手"
    value: dict[str, Any] = {
        "id": "designed-" + str(uuid.uuid4()),
        "name": name,
        "description": text,
        "source": "designed",
        "kind": kind,
        "version": "0.1.0",
        "author": "网页用户",
        "permissions": ["read_selected_files", "write_task_output"],
        "testStatus": "untested",
        "status": "draft",
        "supportedExtensions": extensions,
        "available": False,
        "availabilityDetail": "网页技能草稿；请先上传样本测试",
        "parameters": [],
        "steps": steps if kind == "workflow" else [],
    }
    if kind == "python":
        value["sourceCode"] = "def run(context, params=None):\n    # 网页端不会执行未审阅的 Python 代码\n    return {\"summary\": \"draft\", \"outputs\": []}\n"
        value["availabilityDetail"] = "网页端出于安全原因不执行自定义 Python；可改用工作流技能"
    return value


def parse_json_object(value: str) -> dict[str, Any] | None:
    cleaned = value.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.IGNORECASE)
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        result = json.loads(cleaned[start:end + 1])
    except json.JSONDecodeError:
        return None
    return result if isinstance(result, dict) else None


def design_skill(payload: dict[str, Any]) -> dict[str, Any]:
    kind = "workflow"
    draft = infer_skill_draft(str(payload.get("prompt") or ""), kind)
    base_url, api_key, model = model_settings(payload)
    if api_key:
        sample_context = file_context(str(payload.get("sessionId") or ""), limit=18000) if payload.get("sessionId") else ""
        instruction = (
            "请为政务文件处理应用设计一个可审阅的技能草稿。只返回 JSON，不要 Markdown。"
            "字段必须包括 name、description、supportedExtensions（字符串数组）、steps（数组，"
            "每项含 toolId 和 parameters）、parameters（数组）。toolId 只能使用："
            "inspect_document、docx_to_gongwen、normalize_docx_styles、redact_content、normalize_whitespace、"
            "remove_line_breaks、replace_text、delete_blank_rows、set_xlsx_cell、extract_pdf_text、"
            "merge_xlsx。"
            f"技能类型：{kind}；用户需求：{payload.get('prompt', '')}\n"
            f"样本材料内容（只用于理解版式和规则，不要泄露无关个人信息）：{sample_context or '（未提供可读取的样本正文）'}"
        )
        try:
            generated = parse_json_object(call_openai_compatible(base_url, api_key, model, [{"role": "user", "content": instruction}]))
            if generated:
                for key in ("name", "description", "supportedExtensions", "steps", "parameters"):
                    if key in generated:
                        draft[key] = generated[key]
        except Exception:
            # A model outage must not prevent a deterministic, reviewable draft.
            pass
    return skill_defaults(draft)


def materialize_skill_files(run_id: str, files: list[dict[str, Any]]) -> list[dict[str, Any]]:
    root = UPLOAD_DIR / "skills" / run_id
    root.mkdir(parents=True, exist_ok=True)
    materialized: list[dict[str, Any]] = []
    for item in files:
        raw = base64.b64decode(str(item.get("contentBase64") or ""), validate=True)
        if len(raw) > MAX_UPLOAD_BYTES:
            raise ValueError("单个技能样本超过大小限制")
        name = safe_name(str(item.get("filename") or "样本文件"))
        file_id = str(uuid.uuid4())
        target = root / file_id / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(raw)
        materialized.append({
            "id": file_id,
            "name": name,
            "path": str(target),
            "relativePath": str(item.get("relativePath") or name),
        })
    if not materialized:
        raise ValueError("至少需要一个技能样本")
    return materialized


def engine_operation(tool_id: str, parameters: dict[str, Any] | None = None) -> tuple[str, dict[str, Any]]:
    """Translate the web-facing tool vocabulary to the existing engine schema."""
    value = dict(parameters or {})
    if tool_id == "docx_to_gongwen":
        value.setdefault("preset", "gongwen")
        return "normalize_docx_styles", value
    return tool_id, value


def skill_operations(tool: dict[str, Any], override_parameters: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    steps = tool.get("steps") or [{"toolId": tool.get("id"), "parameters": {}}]
    if override_parameters is not None and tool.get("source") == "builtin":
        steps = [{"toolId": tool.get("id"), "parameters": override_parameters}]
    allowed = {item["id"] for item in builtin_tools() if item.get("available")}
    operations: list[dict[str, Any]] = []
    for index, step in enumerate(steps):
        tool_id = str(step.get("toolId") or "")
        if tool_id not in allowed:
            raise ValueError(f"技能包含网页端不支持的工具：{tool_id}")
        kind, parameters = engine_operation(tool_id, dict(step.get("parameters") or {}))
        operations.append({
            "id": f"skill-operation-{index}",
            "kind": kind,
            "title": tool_id,
            "description": tool_id,
            "fileCount": 0,
            "parameters": parameters,
            "requiresReview": False,
            "available": True,
        })
    return operations


def register_artifacts(session_id: str, outcome: dict[str, Any], source_group_name: str = "处理成果") -> list[dict[str, Any]]:
    artifacts: list[dict[str, Any]] = []
    candidates = list(outcome.get("outputs") or [])
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        for output in candidates:
            path = Path(str(output.get("path") or ""))
            if not path.is_file():
                continue
            artifact_id = str(output.get("id") or uuid.uuid4())
            display_name = path.name if path.name != "index.html" else "处理报告.html"
            pending = getattr(TURN_LOCAL, "pending_artifacts", None)
            if pending is not None:
                artifact = {"id": artifact_id, "sessionId": session_id, "kind": "artifact", "name": display_name,
                            "localSnapshotPath": str(path), "mimeType": mimetypes.guess_type(display_name)[0],
                            "size": path.stat().st_size, "relativePath": display_name, "createdAt": now(),
                            "sourceGroupName": source_group_name}
                pending.append(artifact)
                artifacts.append(artifact)
                continue
            db.execute(
                "INSERT INTO attachments (id, session_id, kind, name, server_path, mime_type, size, relative_path, created_at, source_group_name) VALUES (?, ?, 'artifact', ?, ?, ?, ?, ?, ?, ?)",
                (artifact_id, session_id, display_name, str(path), mimetypes.guess_type(display_name)[0], path.stat().st_size, display_name, now(), source_group_name),
            )
            row = db.execute("SELECT * FROM attachments WHERE id = ?", (artifact_id,)).fetchone()
            artifacts.append(attachment_json(row, artifact=True))
        db.commit()
    return artifacts


def run_engine_job(
    tool: dict[str, Any],
    files: list[dict[str, Any]],
    run_id: str,
    root: Path,
    session_id: str | None = None,
    override_parameters: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if tool.get("kind") == "python":
        raise ValueError("网页端不会执行自定义 Python 技能，请改用工作流技能")
    output_root, report_root = root / "output", root / "report"
    operations = skill_operations(tool, override_parameters)
    hashes = {item["path"]: file_hash(Path(item["path"])) for item in files}
    check_cancelled()
    with ENGINE_LOCK, open(os.devnull, "w", encoding="utf-8") as sink, redirect_stdout(sink):
        outcome = engine.run({
            "schemaVersion": 1,
            "neverOverwrite": True,
            "jobId": run_id,
            "files": files,
            "operations": operations,
            "outputDirectory": str(output_root),
            "reportDirectory": str(report_root),
        })
    for source, digest in hashes.items():
        if file_hash(Path(source)) != digest:
            raise RuntimeError("原件校验失败，本次成果未发布")
    if outcome.get("failedFiles") or outcome.get("status") != "completed":
        failure = (outcome.get("failures") or [{}])[0].get("message", "文件处理未完成")
        raise RuntimeError(str(failure))
    # Reopen every final document before exposing it in the browser.
    for output in outcome.get("outputs", []):
        path = Path(output["path"])
        if path.suffix.lower() == ".docx":
            from docx import Document
            Document(path)
        elif path.suffix.lower() == ".xlsx":
            from openpyxl import load_workbook
            workbook = load_workbook(path)
            workbook.close()
        elif path.suffix.lower() == ".pdf":
            from pypdf import PdfReader
            len(PdfReader(path).pages)
        elif path.suffix.lower() in {".txt", ".md"}:
            path.read_text(encoding="utf-8")
    check_cancelled()
    notices = outcome_notes(outcome)
    if hasattr(TURN_LOCAL, "notices"):
        TURN_LOCAL.notices.extend(notices)
    if session_id:
        artifacts = register_artifacts(session_id, outcome, tool.get("name", "处理成果"))
    else:
        artifacts = []
        for output in list(outcome.get("outputs") or []):
            path = Path(str(output.get("path") or ""))
            if path.is_file():
                artifacts.append({
                    "id": str(uuid.uuid4()), "sessionId": "skill:" + run_id, "kind": "artifact",
                    "name": path.name, "localSnapshotPath": str(path), "size": path.stat().st_size,
                    "createdAt": now(), "relativePath": path.name,
                })
    return outcome, artifacts


def run_skill_engine(tool: dict[str, Any], files: list[dict[str, Any]], run_id: str, parameters: dict[str, Any] | None = None) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    materialized = materialize_skill_files(run_id, files)
    root = ARTIFACT_DIR / "skills" / tool["id"] / run_id
    return run_engine_job(tool, materialized, run_id, root, override_parameters=parameters)


def session_files(session_id: str, attachment_ids: list[str] | None = None, extensions: list[str] | None = None) -> list[dict[str, Any]]:
    wanted = {str(item) for item in (attachment_ids or []) if str(item)}
    allowed_extensions = {str(item).lower().lstrip(".") for item in (extensions or [])}
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        rows = db.execute(
            "SELECT id, name, server_path, relative_path, mime_type FROM attachments WHERE session_id = ? AND server_path IS NOT NULL AND kind != 'artifact' ORDER BY created_at",
            (session_id,),
        ).fetchall()
    files: list[dict[str, Any]] = []
    for row in rows:
        if wanted and row["id"] not in wanted:
            continue
        path = Path(row["server_path"])
        if allowed_extensions and path.suffix.lower().lstrip(".") not in allowed_extensions:
            continue
        if path.is_file():
            files.append({"id": row["id"], "name": row["name"], "path": str(path), "relativePath": row["relative_path"] or row["name"], "mimeType": row["mime_type"]})
    return files


def run_session_tool(session_id: str, tool_id: str, parameters: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    tool = find_skill(tool_id)
    if not tool or not tool.get("available"):
        raise ValueError(f"网页端工具不可用：{tool_id}")
    attachment_ids = parameters.get("attachmentIds") or parameters.get("fileIds")
    if not isinstance(attachment_ids, list):
        attachment_ids = None
    files = session_files(session_id, attachment_ids, tool.get("supportedExtensions"))
    if not files:
        accepted = "、".join(tool.get("supportedExtensions") or [])
        raise ValueError(f"当前会话没有可处理的材料（需要：{accepted}）")
    run_id = str(uuid.uuid4())
    root = ARTIFACT_DIR / session_id / run_id
    outcome, artifacts = run_engine_job(tool, files, run_id, root, session_id=session_id, override_parameters=parameters)
    return outcome, artifacts


def run_redaction(session_id: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    job_id = str(uuid.uuid4())
    root = ARTIFACT_DIR / session_id / job_id
    output_root = root / "output"
    report_root = root / "report"
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        rows = db.execute(
            "SELECT id, name, server_path, relative_path FROM attachments WHERE session_id = ? AND server_path IS NOT NULL AND kind != 'artifact'",
            (session_id,),
        ).fetchall()
    files = [
        {"id": row["id"], "name": row["name"], "path": row["server_path"], "relativePath": row["relative_path"] or row["name"]}
        for row in rows
    ]
    operations = [{
        "id": "web-redact",
        "kind": "redact_content",
        "title": "生成脱敏副本",
        "description": "服务器规则引擎生成独立脱敏副本",
        "fileCount": len(files),
        "parameters": {},
        "requiresReview": True,
        "available": True,
    }]
    with open(os.devnull, "w", encoding="utf-8") as sink, redirect_stdout(sink):
        outcome = engine.run({
            "schemaVersion": 1,
            "neverOverwrite": True,
            "jobId": job_id,
            "files": files,
            "operations": operations,
            "outputDirectory": str(output_root),
            "reportDirectory": str(report_root),
        })
    artifacts: list[dict[str, Any]] = []
    with sqlite3.connect(DB_PATH) as db:
        for output in outcome.get("outputs", []):
            path = Path(output["path"])
            if not path.is_file():
                continue
            artifact_id = str(uuid.uuid4())
            db.execute(
                "INSERT INTO attachments (id, session_id, kind, name, server_path, mime_type, size, relative_path, created_at, source_group_name) VALUES (?, ?, 'artifact', ?, ?, ?, ?, ?, ?, ?)",
                (artifact_id, session_id, "脱敏-" + path.name, str(path), mimetypes.guess_type(path.name)[0], path.stat().st_size, path.name, now(), "脱敏结果"),
            )
            row = db.execute("SELECT * FROM attachments WHERE id = ?", (artifact_id,)).fetchone()
            artifacts.append(attachment_json(row, artifact=True))
        db.commit()
    return outcome, artifacts


def attachment_catalog(session_id: str) -> list[dict[str, Any]]:
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        rows = db.execute(
            "SELECT id, name, kind, mime_type, size, relative_path, source_url FROM attachments WHERE session_id = ? AND kind != 'artifact' ORDER BY created_at",
            (session_id,),
        ).fetchall()
    result: list[dict[str, Any]] = []
    for row in rows:
        name = row["name"]
        result.append({
            "id": row["id"],
            "name": name,
            "type": row["kind"],
            "extension": Path(name).suffix.lower().lstrip("."),
            "size": row["size"],
            "relativePath": row["relative_path"] or name,
            "url": row["source_url"] if row["kind"] == "url" else None,
        })
    return result


def openai_tool_schema(tool_id: str) -> dict[str, Any]:
    if tool_id == "list_skills":
        return {"type": "object", "properties": {}}
    if tool_id == "read_attachment":
        return {
            "type": "object",
            "properties": {"attachmentId": {"type": "string", "description": "附件 ID"}},
            "required": ["attachmentId"],
        }
    tool = find_skill(tool_id)
    parameters: dict[str, Any] = {"type": "object", "properties": {}}
    if tool:
        for item in tool.get("parameters") or []:
            name = str(item.get("name") or "")
            if not name:
                continue
            value_type = str(item.get("type") or "string")
            schema: dict[str, Any] = {"type": value_type, "description": str(item.get("description") or "")}
            if value_type == "array" and name == "replacements":
                schema["items"] = {"type": "object", "properties": {"before": {"type": "string"}, "after": {"type": "string"}}, "required": ["before", "after"]}
            elif value_type == "array":
                schema["items"] = {"type": "string"}
            parameters["properties"][name] = schema
    return parameters


def openai_tool_specs(allow_file_read: bool) -> list[dict[str, Any]]:
    specs: list[dict[str, Any]] = [{
        "type": "function",
        "function": {"name": "list_skills", "description": "列出当前网页端可直接调用的内置能力和已发布技能。", "parameters": openai_tool_schema("list_skills")},
    }]
    if allow_file_read:
        specs.append({
            "type": "function",
            "function": {"name": "read_attachment", "description": "读取一份用户已上传材料的文本内容，供模型理解事实；只能读取当前会话附件。", "parameters": openai_tool_schema("read_attachment")},
        })
    for tool in builtin_tools():
        if not tool.get("available"):
            continue
        specs.append({
            "type": "function",
            "function": {"name": tool["id"], "description": tool["description"], "parameters": openai_tool_schema(tool["id"])},
        })
    specs.append({
        "type": "function",
        "function": {
            "name": "run_skill",
            "description": "运行一个已经通过样本测试并发布的自建工作流技能。需要先列出技能并使用 skillId。",
            "parameters": {"type": "object", "properties": {"skillId": {"type": "string"}, "attachmentIds": {"type": "array", "items": {"type": "string"}}, "parameters": {"type": "object"}}, "required": ["skillId"]},
        },
    })
    return specs


def session_skill_run(session_id: str, skill_id: str, parameters: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    skill = find_skill(skill_id)
    if not skill or skill.get("source") != "designed" or not skill.get("available"):
        raise ValueError("技能不存在、未发布或当前不可用")
    attachment_ids = parameters.get("attachmentIds")
    if not isinstance(attachment_ids, list):
        attachment_ids = None
    files = session_files(session_id, attachment_ids, skill.get("supportedExtensions"))
    if not files:
        raise ValueError("当前会话没有符合该技能格式要求的材料")
    run_id = str(uuid.uuid4())
    return run_engine_job(skill, files, run_id, ARTIFACT_DIR / session_id / run_id, session_id=session_id, override_parameters=parameters.get("parameters") if isinstance(parameters.get("parameters"), dict) else None)


def read_session_attachment(session_id: str, attachment_id: str) -> dict[str, Any]:
    with sqlite3.connect(DB_PATH) as db:
        db.row_factory = sqlite3.Row
        row = db.execute(
            "SELECT id, name, server_path, source_url, mime_type, size, relative_path FROM attachments WHERE id = ? AND session_id = ? AND kind != 'artifact'",
            (attachment_id, session_id),
        ).fetchone()
    if not row:
        raise ValueError("附件不存在或不是可读取的文件")
    if row["source_url"]:
        text = fetch_public_url(str(row["source_url"]))
    else:
        if not row["server_path"]:
            raise ValueError("附件没有可读取内容")
        path = Path(row["server_path"])
        if not path.is_file():
            raise ValueError("附件文件已不存在")
        text = engine.extract_text_preview(path)
    return {"attachmentId": row["id"], "name": row["name"], "relativePath": row["relative_path"] or row["name"], "text": text[:18000], "characters": len(text)}


def outcome_notes(outcome: dict[str, Any]) -> list[str]:
    return list(dict.fromkeys(
        safe_error(Exception(str(note)))
        for output in outcome.get("outputs", []) for note in output.get("notes", [])
    ))


def tool_result_summary(outcome: dict[str, Any], artifacts: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "status": outcome.get("status"),
        "completedFiles": outcome.get("completedFiles", 0),
        "failedFiles": outcome.get("failedFiles", 0),
        "failures": outcome.get("failures", [])[:10],
        "notes": outcome_notes(outcome),
        "artifacts": [{"id": item.get("id"), "name": item.get("name"), "size": item.get("size")} for item in artifacts],
        "report": next((item.get("name") for item in artifacts if str(item.get("name", "")).endswith(".html")), None),
    }


def execute_agent_tool(session_id: str, tool_id: str, arguments: dict[str, Any], allow_file_read: bool) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if tool_id == "list_skills":
        return {"skills": [{"id": item["id"], "name": item["name"], "description": item["description"], "supportedExtensions": item.get("supportedExtensions", []), "outputFormats": item.get("outputFormats", []), "available": item.get("available", False)} for item in all_skills()]}, []
    if tool_id == "read_attachment":
        if not allow_file_read:
            raise ValueError("本轮已关闭模型读取材料正文；可以继续执行本地规则工具，但不会把正文发给模型")
        return read_session_attachment(session_id, str(arguments.get("attachmentId") or "")), []
    if tool_id == "run_skill":
        outcome, artifacts = session_skill_run(session_id, str(arguments.get("skillId") or ""), arguments)
        return tool_result_summary(outcome, artifacts), artifacts
    outcome, artifacts = run_session_tool(session_id, tool_id, arguments)
    return tool_result_summary(outcome, artifacts), artifacts


def model_text(message: dict[str, Any]) -> str:
    content = message.get("content", "")
    if isinstance(content, list):
        return "".join(str(item.get("text", "")) for item in content if isinstance(item, dict))
    return str(content or "")


def agent_system_prompt(session_id: str, allow_file_read: bool) -> str:
    catalog = json.dumps(attachment_catalog(session_id), ensure_ascii=False)
    privacy = (
        "本轮允许你按需读取附件正文，但只使用当前会话材料，不要编造事实。"
        if allow_file_read
        else
        "本轮禁止把附件正文发送给你；你可以调用确定性文件工具直接处理附件，并根据工具结果汇报，不要猜测正文内容。"
    )
    return f"""你是 ZHAI（珠海 AI）政务材料 Agent，不是只会聊天的问答机器人。你的职责是理解任务、选择已验证工具、实际处理用户材料，并把成果交付给用户。

执行原则：
1. {privacy}
2. 用户要求“按公文格式/公文排版/套红/Word规范”时，优先调用 docx_to_gongwen；已有 DOCX 必须生成新文件，绝不覆盖原件。
3. 用户要求脱敏、检查、提取、转换、清理、批量修改时，必须调用对应工具，不要只描述应该怎么做。
4. 工具参数中的 attachmentIds 必须使用当前会话材料目录里的 id；不确定时省略，让工具处理当前会话匹配格式的文件。
5. 工具失败要如实说明原因并给替代建议；工具成功后说明处理数量、生成的成果名称和需要人工复核的地方。
6. 需要引用材料事实时先调用 read_attachment；不能以文件名推断内容。涉及发布、承诺、个人信息和政策结论时保持审慎。
7. 可以先检查再处理；复杂任务可以连续调用多个工具，但每一步都必须有实际作用。不要调用当前不可用的 Word 转 PDF。

当前会话附件目录（id 是工具调用的唯一标识）：
{catalog}

回答可以使用 Markdown，但最终要简明、可验收。"""


def emit_agent_event(events: list[dict[str, Any]], event: dict[str, Any]) -> None:
    if event.get("type") == "artifact_created" and getattr(TURN_LOCAL, "pending_artifacts", None) is not None:
        return
    events.append(event)


def run_agent_turn(session_id: str, content: str, payload: dict[str, Any], turn_id: str) -> tuple[str, list[dict[str, Any]], list[dict[str, Any]]]:
    base_url, api_key, model = model_settings(payload)
    allow_file_read = bool(payload.get("allowModelFileAccess"))
    events: list[dict[str, Any]] = getattr(TURN_LOCAL, "events", [])
    artifacts: list[dict[str, Any]] = []
    if not api_key:
        def local_run(tool_id: str, label: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
            run_id = str(uuid.uuid4())
            events.append({"type": "tool_started", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "toolId": tool_id, "summary": f"正在执行：{label}"})
            events.append({"type": "tool_progress", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "progress": 18, "detail": "规则引擎已启动"})
            result, created = execute_agent_tool(session_id, tool_id, {}, allow_file_read)
            for artifact in created:
                events.append({"type": "artifact_created", "sessionId": session_id, "turnId": turn_id, "artifact": artifact})
            events.append({"type": "tool_progress", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "progress": 100, "detail": "处理完成并已生成独立成果"})
            return result, created
        lower = content.lower()
        if any(key in content for key in ("公文", "排版", "套版")) or "gongwen" in lower:
            result, artifacts = local_run("docx_to_gongwen", "现有 Word 一键公文排版")
            outcome = result
            return f"已按公文版式生成 {outcome.get('completedFiles', 0)} 个 Word 新版本。原文件未覆盖，请在成果区下载；如有特殊发文机关标志、印章或页码要求，请人工复核。", artifacts, events
        if "脱敏" in content or "敏感信息" in content:
            outcome, artifacts = local_run("redact_content", "敏感信息脱敏")
            return f"已完成脱敏处理：成功生成 {outcome.get('completedFiles', 0)} 个独立副本，原文件未覆盖。", artifacts, events
        if any(key in content for key in ("检查文件", "检查材料", "审查材料", "规范性")):
            outcome, artifacts = local_run("inspect_document", "检查文件")
            return f"已完成材料检查：处理 {outcome.get('completedFiles', 0)} 个文件，并生成检查报告。", artifacts, events
        raise ValueError("请在“模型服务”中配置 API Key；不配置模型时，网页端仍支持“公文排版”“脱敏”“文件检查”等明确的本地规则任务。")

    history = get_session(session_id) or {"messages": []}
    messages: list[dict[str, Any]] = [{"role": "system", "content": agent_system_prompt(session_id, allow_file_read)}]
    # Earlier assistant answers can quote file contents. Do not resend those
    # when the current turn revokes model access to an attached document.
    previous = history.get("messages", [])[-20:] if allow_file_read or not attachment_catalog(session_id) else []
    for item in previous:
        if item.get("role") in {"user", "assistant"}:
            messages.append({"role": item["role"], "content": item["content"]})
    if not messages or messages[-1].get("content") != content:
        messages.append({"role": "user", "content": content})
    tools = openai_tool_specs(allow_file_read)
    last_error: Exception | None = None
    native = True
    for _ in range(8):
        check_cancelled()
        try:
            response = call_openai_compatible_payload(base_url, api_key, model, messages, tools=tools if native else None)
            choice = response["choices"][0]
            assistant = choice.get("message") or {}
        except Exception as exc:  # noqa: BLE001 - fallback keeps gateways without tool support usable
            check_cancelled()
            last_error = exc
            if native and ("HTTP 400" in str(exc) or "HTTP 422" in str(exc)):
                native = False
                messages = [
                    ({"role": "user", "content": f"前一步工具结果：{item.get('content', '')}"} if item.get("role") == "tool" else item)
                    for item in messages
                ]
                messages[0]["content"] = str(messages[0]["content"]) + "\n\n你的接口不支持原生 tools。请改用 JSON 工具协议：只返回 {\\\"message\\\":\\\"给用户看的话\\\",\\\"calls\\\":[{\\\"toolId\\\":\\\"工具名\\\",\\\"parameters\\\":{}}]}；需要执行工具时 calls 不得为空。"
                continue
            raise
        text = model_text(assistant)
        tool_calls = assistant.get("tool_calls") or []
        if not native and not tool_calls:
            fallback = parse_json_object(text)
            calls = fallback.get("calls") if fallback else []
            if isinstance(calls, list) and calls:
                tool_calls = [{"id": str(uuid.uuid4()), "function": {"name": str(item.get("toolId") or ""), "arguments": json.dumps(item.get("parameters") or {}, ensure_ascii=False)}} for item in calls if isinstance(item, dict)]
                text = str(fallback.get("message") or "")
            elif fallback and str(fallback.get("message") or "").strip():
                return str(fallback["message"]).strip(), artifacts, events
        if not tool_calls:
            if text.strip():
                return text.strip(), artifacts, events
            break
        assistant_turn = {"role": "assistant", "content": text, "tool_calls": tool_calls}
        for field in ("reasoning_content", "reasoning_text"):
            if field in assistant:
                assistant_turn[field] = assistant[field]
        messages.append(assistant_turn)
        for call in tool_calls:
            check_cancelled()
            function = call.get("function") or {}
            tool_id = str(function.get("name") or "")
            raw_args = function.get("arguments") or "{}"
            run_id = str(uuid.uuid4())
            try:
                arguments = json.loads(raw_args) if isinstance(raw_args, str) else raw_args
                if not isinstance(arguments, dict):
                    raise ValueError("工具参数必须是 JSON 对象")
                tool = find_skill(tool_id)
                label = str(tool.get("name") if tool else tool_id)
                emit_agent_event(events, {"type": "tool_started", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "toolId": tool_id, "summary": f"正在执行：{label}"})
                emit_agent_event(events, {"type": "tool_progress", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "progress": 18, "detail": "已确认材料范围，开始调用服务器文档引擎"})
                result, created = execute_agent_tool(session_id, tool_id, arguments, allow_file_read)
                artifacts.extend(created)
                for artifact in created:
                    emit_agent_event(events, {"type": "artifact_created", "sessionId": session_id, "turnId": turn_id, "artifact": artifact})
                detail = result if isinstance(result, dict) else {"result": result}
                emit_agent_event(events, {"type": "tool_progress", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "progress": 100, "detail": json.dumps(detail, ensure_ascii=False)[:800]})
                if native:
                    messages.append({"role": "tool", "tool_call_id": call.get("id", run_id), "name": tool_id, "content": json.dumps(result, ensure_ascii=False)})
                else:
                    messages.append({"role": "user", "content": f"工具 {tool_id} 已执行，结果如下：{json.dumps(result, ensure_ascii=False)}。请根据结果继续；如需其他工具，继续返回 JSON 工具协议。"})
            except Exception as exc:  # noqa: BLE001 - one tool failure is fed back to the model
                emit_agent_event(events, {"type": "tool_failed", "sessionId": session_id, "turnId": turn_id, "runId": run_id, "message": str(exc)})
                if native:
                    messages.append({"role": "tool", "tool_call_id": call.get("id", run_id), "name": tool_id, "content": json.dumps({"error": str(exc)}, ensure_ascii=False)})
                else:
                    messages.append({"role": "user", "content": f"工具 {tool_id} 执行失败：{exc}。请返回 JSON 工具协议并决定是否换一个工具。"})
    if last_error:
        raise last_error
    raise RuntimeError("模型没有返回可执行的回答")


def generate_citizen_reply(payload: dict[str, Any]) -> dict[str, Any]:
    department = str(payload.get("department") or "承办单位").strip()
    request = str(payload.get("citizenRequest") or "").strip()
    verification = str(payload.get("verification") or "").strip()
    handling = str(payload.get("handling") or "").strip()
    case_result = str(payload.get("caseResult") or "").strip()
    next_step = str(payload.get("nextStep") or "").strip()
    platform = str(payload.get("platform") or "general")
    if not request:
        raise ValueError("请至少填写市民诉求")
    findings: list[dict[str, Any]] = []
    if not verification:
        findings.append({"level": "warning", "code": "verification", "message": "缺少核查时间、对象或依据，发布前必须补充。"})
    if not handling:
        findings.append({"level": "warning", "code": "handling", "message": "缺少承办措施，不能只写“已处理”。"})
    if not next_step:
        findings.append({"level": "warning", "code": "next_step", "message": "缺少后续跟进安排或联系人。"})
    if re.search(r"(绝对|一定|保证|彻底解决|100%)", verification + handling + next_step):
        findings.append({"level": "error", "code": "absolute_promise", "message": "存在绝对化或保证性表述，请改为可核验、可兑现的表达。"})
    opening = "网友您好！\n\n" if platform == "people" else "您好！\n\n" if platform == "provincial" else "您好，感谢您反映情况。\n\n"
    deterministic = f"{opening}关于您反映的“{request[:120]}”，{department}已开展核查。\n\n"
    deterministic += (f"经核查，{verification}\n\n" if verification else "经核查，相关情况仍需补充核实。\n\n")
    deterministic += (f"处理情况：{handling}" + (f"{case_result}" if case_result else "") + "\n\n") if handling else "处理情况：承办单位正在核实处理，请补充具体措施后发布。\n\n"
    deterministic += f"后续安排：{next_step or '请补充后续跟进时间、方式和责任人。'}"
    base_url, api_key, model = model_settings(payload)
    if api_key:
        instruction = f"""你是珠海政务工单答复审校助手。请根据以下字段起草一版可审阅的中文回复，不补造事实，不擅自承诺。只返回 JSON：{{\"answer\":\"回复正文\",\"findings\":[{{\"level\":\"error|warning|info\",\"code\":\"代码\",\"message\":\"问题\"}}]}}。答复要包含诉求回应、核查事实、处理措施、后续安排；缺失内容用[待补充]标记，并指出缺项。\n平台：{platform}\n承办单位：{department}\n市民诉求：{request}\n核查情况：{verification}\n处理措施：{handling}\n处理结果：{case_result}\n后续安排：{next_step}\n特别要求：{payload.get('instructions') or ''}"""
        try:
            generated = parse_json_object(call_openai_compatible(base_url, api_key, model, [{"role": "user", "content": instruction}]))
            if generated and str(generated.get("answer") or "").strip():
                return {"answer": str(generated["answer"]).strip(), "findings": findings + list(generated.get("findings") or []), "references": [{"source": "ZHAI", "title": "政务工单答复规范"}], "returnedModel": model}
        except Exception:
            # The deterministic draft remains usable if the gateway is unavailable.
            pass
    return {"answer": deterministic, "findings": findings or [{"level": "info", "code": "review", "message": "请在发布前复核事实、时间、责任单位和联系方式。"}], "references": [{"source": "ZHAI", "title": "政务工单答复规范"}], "returnedModel": payload.get("model")}


def start_background_turn(session_id: str, content: str, payload: dict[str, Any]) -> dict[str, Any]:
    model_settings(payload)
    with TURN_LOCK:
        if any(not state.get("finished") for state in TURNS.values()):
            raise ValueError("试用版一次处理一个任务，请等当前任务结束后再提交")
        # Only retain the recent terminal states; documents and chat live in SQLite.
        for key in list(TURNS)[:-20]:
            del TURNS[key]
        turn_id = str(uuid.uuid4())
        state = {"turnId": turn_id, "sessionId": session_id, "status": "running", "events": [], "error": None, "cancelled": threading.Event()}
        TURNS[turn_id] = state
        with sqlite3.connect(DB_PATH) as db:
            db.execute("INSERT INTO messages VALUES (?, ?, 'user', ?, ?, ?)", (str(uuid.uuid4()), session_id, content, now(), turn_id))
            db.execute("UPDATE sessions SET status = 'running', updated_at = ? WHERE id = ?", (now(), session_id))
        threading.Thread(target=finish_background_turn, args=(state, content, dict(payload)), daemon=True).start()
    return {"turnId": turn_id, "status": "running"}


def finish_background_turn(state: dict[str, Any], content: str, payload: dict[str, Any]) -> None:
    session_id, turn_id = state["sessionId"], state["turnId"]
    TURN_LOCAL.cancelled = state["cancelled"]
    TURN_LOCAL.events = state["events"]
    TURN_LOCAL.pending_artifacts = []
    TURN_LOCAL.notices = []
    try:
        answer, artifacts, events = run_agent_turn(session_id, content, payload, turn_id)
        check_cancelled()
        notices = [n for n in dict.fromkeys(TURN_LOCAL.notices)
                   if any(word in n for word in ("字体", "保留", "未", "跳过", "复杂", "提示"))]
        missing = [n for n in notices if n not in answer]
        if missing:
            answer += "\n\n处理说明：\n" + "\n".join("- " + n for n in missing)
        TURN_LOCAL.pending_artifacts = None
        # All results become visible together, after the full turn succeeds.
        artifacts = register_artifacts(session_id, {"outputs": [{"id": item["id"], "path": item["localSnapshotPath"]} for item in artifacts]})
        for artifact in artifacts:
            events.append({"type": "artifact_created", "sessionId": session_id, "turnId": turn_id, "artifact": artifact})
        events.append({"type": "assistant_delta", "sessionId": session_id, "turnId": turn_id, "text": answer})
        message_id = str(uuid.uuid4())
        with sqlite3.connect(DB_PATH) as db:
            db.execute("INSERT INTO messages VALUES (?, ?, 'assistant', ?, ?, ?)", (message_id, session_id, answer, now(), turn_id))
        events.append({"type": "turn_completed", "sessionId": session_id, "turnId": turn_id, "messageId": message_id})
        state["status"] = "completed"
    except Exception as exc:
        state["error"] = safe_error(exc)
        state["status"] = "cancelled" if state["cancelled"].is_set() else "failed"
    finally:
        with sqlite3.connect(DB_PATH) as db:
            db.execute("UPDATE sessions SET status = 'idle', updated_at = ? WHERE id = ?", (now(), session_id))
        # Complete status is published only after the session is idle.
        state["finished"] = True
        TURN_LOCAL.__dict__.clear()
        payload.clear()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def parse_request(self) -> bool:
        if not super().parse_request():
            return False
        if self.command == "GET" and urlparse(self.path).path == "/api/health":
            return True
        expected = f"{os.environ.get('ZHAI_ACCESS_USER', 'zhai')}:{os.environ.get('ZHAI_ACCESS_PASSWORD', '')}"
        actual = ""
        try:
            scheme, value = self.headers.get("Authorization", "").split(" ", 1)
            if scheme.lower() == "basic":
                actual = base64.b64decode(value, validate=True).decode("utf-8")
        except (ValueError, UnicodeError):
            pass
        if not os.environ.get("ZHAI_ACCESS_PASSWORD") or not hmac.compare_digest(actual.encode(), expected.encode()):
            self.close_connection = True
            self.send_response(401)
            self.send_header("WWW-Authenticate", 'Basic realm="zhai", charset="UTF-8"')
            self.send_header("Content-Length", "0")
            self.end_headers()
            return False
        if self.command in {"POST", "PATCH", "DELETE"}:
            origin = self.headers.get("Origin")
            if origin and urlparse(origin).netloc != self.headers.get("Host"):
                self.close_connection = True
                self.send_error(403, "Cross-origin request denied")
                return False
            if self.command in {"POST", "PATCH"} and self.headers.get_content_type() != "application/json":
                self.close_connection = True
                self.send_error(415, "JSON required")
                return False
        return True

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        super().end_headers()

    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write("[zhai-web] " + (format % args) + "\n")

    def send_json(self, payload: Any, status: int = 200) -> None:
        body = safe_error(Exception(json_bytes(payload).decode("utf-8"))).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD_BYTES * 2:
            raise ValueError("请求体过大")
        raw = self.rfile.read(length)
        value = json.loads(raw.decode("utf-8"))
        if not isinstance(value, dict):
            raise ValueError("请求必须是 JSON 对象")
        return value

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PATCH, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:
        try:
            parsed = urlparse(self.path)
            path = parsed.path
            if path == "/api/health":
                return self.send_json({"ok": True, "service": "zhai-web", "version": WEB_VERSION, "instance": os.environ.get("ZHAI_INSTANCE_ID", ""), "time": now()})
            if path == "/api/profiles":
                return self.send_json(profile_list())
            if path == "/api/tools":
                return self.send_json(all_skills())
            if path == "/api/skills":
                return self.send_json(all_skills())
            if path == "/api/skills/recent":
                return self.send_json([item["id"] for item in stored_skills()[:12]])
            if path == "/api/skills/deleted":
                return self.send_json(deleted_skills())
            if path == "/api/capabilities":
                return self.send_json({
                    "webAgent": {"available": True, "status": "ready", "detail": "支持原生 tools 与 JSON 工具协议回退"},
                    "gongwenDocx": {"available": True, "status": "ready", "detail": "服务器端可生成独立公文格式 DOCX"},
                    "skills": {"available": True, "status": "ready", "detail": "支持设计、样本测试、发布、运行和回收站"},
                    "citizenReply": {"available": True, "status": "ready", "detail": "支持结构化答复草稿与风险提示"},
                    "pythonSidecar": {"available": True, "status": "ready", "detail": "服务器 Python 文档处理引擎可用"},
                    "visionOcr": {"available": False, "status": "unavailable", "detail": "网页容器未配置 macOS Vision OCR"},
                    "paddleOcr": {"available": False, "status": "optional_not_installed", "detail": "未安装可选 OCR 组件"},
                })
            if path == "/api/sessions":
                with sqlite3.connect(DB_PATH) as db:
                    db.row_factory = sqlite3.Row
                    rows = db.execute("SELECT * FROM sessions ORDER BY updated_at DESC").fetchall()
                return self.send_json([row_session(row) for row in rows])
            match = re.fullmatch(r"/api/turns/([A-Za-z0-9-]+)", path)
            if match:
                with TURN_LOCK:
                    state = TURNS.get(match.group(1))
                    if not state:
                        return self.send_json({"error": "任务状态已过期或服务器已重启，请刷新会话"}, 404)
                    offset = max(0, int(parse_qs(parsed.query).get("offset", ["0"])[0]))
                    events = list(state["events"])
                    return self.send_json({"status": state["status"] if state.get("finished") else "running", "error": state["error"],
                                           "events": events[offset:], "offset": len(events), "detail": get_session(state["sessionId"]) if state.get("finished") else None})
            match = re.fullmatch(r"/api/sessions/([^/]+)", path)
            if match:
                detail = get_session(unquote(match.group(1)))
                if not detail:
                    return self.send_json({"error": "会话不存在"}, 404)
                return self.send_json(detail)
            match = re.fullmatch(r"/api/skills/([^/]+)", path)
            if match:
                skill = find_skill(unquote(match.group(1)))
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                return self.send_json(skill)
            if path == "/api/artifacts/download":
                raw_path = unquote(parse_qs(parsed.query).get("path", [""])[0])
                target = Path(raw_path).resolve()
                if ARTIFACT_DIR not in target.parents or not target.exists() or not (target.is_file() or target.is_dir()):
                    return self.send_json({"error": "成果路径无效"}, 400)
                download_name = target.name
                if target.is_dir():
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
                        for item in target.rglob("*"):
                            if item.is_file():
                                archive.write(item, item.relative_to(target))
                    body = buffer.getvalue()
                    download_name += ".zip"
                    content_type = "application/zip"
                else:
                    body = target.read_bytes()
                    content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                fallback_name = re.sub(r"[^A-Za-z0-9._-]+", "_", download_name) or "zhai-artifact"
                self.send_header("Content-Disposition", f"attachment; filename=\"{fallback_name}\"; filename*=UTF-8''{quote(download_name)}")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            if not path.startswith("/api/"):
                static_root = ROOT / "dist"
                requested = (static_root / unquote(path.lstrip("/"))).resolve()
                if requested.is_file() and static_root in requested.parents:
                    body = requested.read_bytes()
                    self.send_response(200)
                    self.send_header("Content-Type", mimetypes.guess_type(requested.name)[0] or "application/octet-stream")
                    self.send_header("Content-Length", str(len(body)))
                    self.end_headers()
                    self.wfile.write(body)
                    return
                index = static_root / "index.html"
                if index.is_file():
                    body = index.read_bytes()
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(body)))
                    self.end_headers()
                    self.wfile.write(body)
                    return
            self.send_json({"error": "未找到接口"}, 404)
        except Exception as exc:  # noqa: BLE001 - API returns a safe message
            self.send_json({"error": str(exc)}, 500)

    def do_POST(self) -> None:
        try:
            parsed = urlparse(self.path)
            path = parsed.path
            payload = self.read_json()
            if path == "/api/citizen-reply":
                return self.send_json(generate_citizen_reply(payload))
            if path == "/api/skills/design":
                return self.send_json(design_skill(payload))
            if path == "/api/skills":
                tool = payload.get("tool")
                if not isinstance(tool, dict):
                    return self.send_json({"error": "缺少技能内容"}, 400)
                return self.send_json(save_skill({**tool, "source": "designed", "status": "draft", "testStatus": "untested", "available": False}))
            match = re.fullmatch(r"/api/skills/([^/]+)/test", path)
            if match:
                skill_id = unquote(match.group(1))
                skill = find_skill(skill_id)
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                if skill.get("source") != "designed":
                    return self.send_json({"error": "内置能力无需发布"}, 400)
                run_id = "test-" + str(uuid.uuid4())
                try:
                    outcome, _ = run_skill_engine(skill, list(payload.get("files") or []), run_id, payload.get("parameters"))
                    completed = int(outcome.get("completedFiles", 0))
                    failed = int(outcome.get("failedFiles", 0))
                    tested = {
                        **skill,
                        "testStatus": "passed" if failed == 0 and completed > 0 else "failed",
                        "status": "draft" if failed == 0 and completed > 0 else "test_failed",
                        "available": False,
                        "availabilityDetail": f"样本测试{'通过' if failed == 0 and completed > 0 else '未通过'}（成功 {completed}，失败 {failed}）",
                    }
                    save_skill(tested)
                    return self.send_json(tested)
                except Exception as exc:
                    failed = {**skill, "testStatus": "failed", "status": "test_failed", "available": False, "availabilityDetail": str(exc)}
                    save_skill(failed)
                    return self.send_json(failed)
            match = re.fullmatch(r"/api/skills/([^/]+)/publish", path)
            if match:
                skill = find_skill(unquote(match.group(1)))
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                if skill.get("source") != "designed":
                    return self.send_json({"error": "内置能力已可直接使用"}, 400)
                if skill.get("testStatus") != "passed":
                    return self.send_json({"error": "请先用真实样本测试通过"}, 400)
                published = {**skill, "status": "published", "available": True, "availabilityDetail": "已发布，可由网页 Agent 或能力包页面运行"}
                return self.send_json(save_skill(published))
            match = re.fullmatch(r"/api/skills/([^/]+)/disable", path)
            if match:
                skill = find_skill(unquote(match.group(1)))
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                if skill.get("source") != "designed":
                    return self.send_json({"error": "内置能力不能停用"}, 400)
                return self.send_json(save_skill({**skill, "status": "disabled", "available": False, "availabilityDetail": "已停用"}))
            match = re.fullmatch(r"/api/skills/([^/]+)/restore", path)
            if match:
                skill = find_skill(unquote(match.group(1)), include_deleted=True)
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                return self.send_json(save_skill(skill, deleted=False))
            if path == "/api/skills/run":
                skill = find_skill(str(payload.get("skillId") or ""))
                if not skill:
                    return self.send_json({"error": "技能不存在"}, 404)
                if not skill.get("available"):
                    return self.send_json({"error": "技能尚未发布或当前不可用"}, 400)
                run_id = str(uuid.uuid4())
                outcome, artifacts = run_skill_engine(skill, list(payload.get("files") or []), run_id, payload.get("parameters"))
                return self.send_json({
                    "summary": f"“{skill['name']}”已处理 {outcome.get('completedFiles', 0)} 个输入" + ("\n" + "\n".join(outcome_notes(outcome)) if outcome_notes(outcome) else ""),
                    "outputDirectory": str(ARTIFACT_DIR / "skills" / skill["id"] / run_id),
                    "artifacts": artifacts,
                })
            if path == "/api/sessions":
                session_id = str(uuid.uuid4())
                timestamp = now()
                title = str(payload.get("title") or "新会话").strip()[:120]
                with sqlite3.connect(DB_PATH) as db:
                    db.execute(
                        "INSERT INTO sessions VALUES (?, ?, ?, ?, 'idle', ?, ?)",
                        (session_id, title, str(payload.get("profileId") or "default"), str(payload.get("model") or ""), timestamp, timestamp),
                    )
                detail = get_session(session_id)
                return self.send_json(detail["session"] if detail else {})
            match = re.fullmatch(r"/api/sessions/([^/]+)/attachments", path)
            if match:
                session_id = unquote(match.group(1))
                if not get_session(session_id):
                    return self.send_json({"error": "会话不存在"}, 404)
                kind = str(payload.get("kind") or "file")
                if kind == "url":
                    attachment_id = str(uuid.uuid4())
                    with sqlite3.connect(DB_PATH) as db:
                        db.execute(
                            "INSERT INTO attachments (id, session_id, kind, name, source_url, created_at) VALUES (?, ?, 'url', ?, ?, ?)",
                            (attachment_id, session_id, "公开网址", str(payload.get("url") or ""), now()),
                        )
                        db.commit()
                        db.row_factory = sqlite3.Row
                        row = db.execute("SELECT * FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
                    return self.send_json(attachment_json(row))
                encoded = str(payload.get("contentBase64") or "")
                raw = base64.b64decode(encoded, validate=True)
                if len(raw) > MAX_UPLOAD_BYTES:
                    return self.send_json({"error": "单个文件超过大小限制"}, 413)
                filename = safe_name(str(payload.get("filename") or "未命名文件"))
                attachment_id = str(uuid.uuid4())
                target_dir = UPLOAD_DIR / session_id / attachment_id
                target_dir.mkdir(parents=True, exist_ok=True)
                target = (target_dir / filename).resolve()
                if target.parent != target_dir.resolve():
                    return self.send_json({"error": "文件名无效"}, 400)
                target.write_bytes(raw)
                with sqlite3.connect(DB_PATH) as db:
                    db.execute(
                        "INSERT INTO attachments (id, session_id, kind, name, server_path, mime_type, size, relative_path, created_at, source_group_id, source_group_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (attachment_id, session_id, kind, filename, str(target), payload.get("mimeType"), len(raw), str(payload.get("relativePath") or filename), now(), str(payload.get("groupId") or "") or None, "上传文件夹" if kind == "folder" else None),
                    )
                    db.commit()
                    db.row_factory = sqlite3.Row
                    row = db.execute("SELECT * FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
                return self.send_json(attachment_json(row))
            match = re.fullmatch(r"/api/sessions/([^/]+)/messages", path)
            if match:
                session_id = unquote(match.group(1))
                detail = get_session(session_id)
                if not detail:
                    return self.send_json({"error": "会话不存在"}, 404)
                turn_id = str(uuid.uuid4())
                content = str(payload.get("content") or "").strip()
                if not content:
                    return self.send_json({"error": "消息不能为空"}, 400)
                return self.send_json(start_background_turn(session_id, content, payload), 202)
            match = re.fullmatch(r"/api/sessions/([^/]+)/cancel", path)
            if match:
                with TURN_LOCK:
                    for state in TURNS.values():
                        if state["sessionId"] == unquote(match.group(1)) and not state.get("finished"):
                            state["cancelled"].set()
                return self.send_json({"status": "stopping"})
            if path == "/api/profiles":
                if (payload.get("provider") or "openai_compatible") not in {"deepseek", "openai_compatible"}:
                    return self.send_json({"error": "仅支持 DeepSeek 和兼容接口；请核对地址后保存为兼容接口。"}, 400)
                return self.send_json({
                    "id": payload.get("id") or "default",
                    "provider": payload.get("provider") or "openai_compatible",
                    "displayName": payload.get("displayName") or "自配模型",
                    "baseUrl": payload.get("baseUrl") or os.environ.get("ZHAI_LLM_BASE_URL", "https://api.deepseek.com/v1"),
                    "defaultModel": payload.get("defaultModel") or os.environ.get("ZHAI_LLM_MODEL", ""),
                    "keyStored": bool(model_settings({**payload, "model": payload.get("defaultModel")})[1]),
                })
            if path == "/api/profiles/test":
                base_url, api_key, model = model_settings(payload)
                if not api_key:
                    return self.send_json({"success": False, "message": "未提供 API Key", "testedAt": now()}, 400)
                started = time.monotonic()
                call_openai_compatible(
                    base_url, api_key, model,
                    [{"role": "user", "content": "请只回复：连接成功"}],
                )
                return self.send_json({"success": True, "returnedModel": payload.get("model"), "latencyMs": round((time.monotonic() - started) * 1000), "testedAt": now(), "message": "连接成功；未发送任何材料内容"})
            self.send_json({"error": "未找到接口"}, 404)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, 500)

    def do_PATCH(self) -> None:
        try:
            match = re.fullmatch(r"/api/sessions/([^/]+)", urlparse(self.path).path)
            if not match:
                return self.send_json({"error": "未找到接口"}, 404)
            payload = self.read_json()
            title = str(payload.get("title") or "新会话").strip()[:120]
            with sqlite3.connect(DB_PATH) as db:
                db.execute("UPDATE sessions SET title = ?, updated_at = ? WHERE id = ?", (title, now(), unquote(match.group(1))))
                db.commit()
            detail = get_session(unquote(match.group(1)))
            if not detail:
                return self.send_json({"error": "会话不存在"}, 404)
            self.send_json(detail["session"])
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, 500)

    def do_DELETE(self) -> None:
        try:
            parsed = urlparse(self.path)
            permanent_match = re.fullmatch(r"/api/skills/([^/]+)/permanent", parsed.path)
            skill_match = re.fullmatch(r"/api/skills/([^/]+)", parsed.path)
            match = re.fullmatch(r"/api/sessions/([^/]+)", parsed.path)
            attachment_match = re.fullmatch(r"/api/sessions/([^/]+)/attachments/([^/]+)", parsed.path)
            with sqlite3.connect(DB_PATH) as db:
                db.execute("PRAGMA foreign_keys = ON")
                if permanent_match:
                    db.execute("DELETE FROM skills WHERE id = ?", (unquote(permanent_match.group(1)),))
                elif skill_match:
                    skill = find_skill(unquote(skill_match.group(1)))
                    if not skill or skill.get("source") != "designed":
                        return self.send_json({"error": "只能删除自己创建的技能"}, 400)
                    db.execute("UPDATE skills SET deleted = 1, updated_at = ? WHERE id = ?", (now(), unquote(skill_match.group(1))))
                elif attachment_match:
                    db.execute("DELETE FROM attachments WHERE session_id = ? AND id = ?", (unquote(attachment_match.group(1)), unquote(attachment_match.group(2))))
                elif match:
                    db.execute("DELETE FROM sessions WHERE id = ?", (unquote(match.group(1)),))
                else:
                    return self.send_json({"error": "未找到接口"}, 404)
                db.commit()
            self.send_json({})
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, 500)


def main() -> int:
    if not os.environ.get("ZHAI_ACCESS_PASSWORD"):
        raise SystemExit("未配置访问密码，请使用部署包的启动脚本")
    # One writer per workspace, including when two Devbox terminals start it.
    with data_lock(DATA_DIR):
        init_db()
        with sqlite3.connect(DB_PATH) as db:
            db.execute("UPDATE sessions SET status = 'idle' WHERE status = 'running'")
        host = os.environ.get("ZHAI_WEB_HOST", "0.0.0.0")
        port = int(os.environ.get("ZHAI_WEB_PORT", "8000"))
        server = ThreadingHTTPServer((host, port), Handler)
        print(f"zhai Web API listening on {host}:{port}", flush=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            return 0
        finally:
            server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
