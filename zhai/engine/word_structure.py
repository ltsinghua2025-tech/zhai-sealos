"""Deterministic DOCX structure/locator editing without rebuilding its package."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _load(path: Path):
    if not path.is_absolute() or path.suffix.lower() != ".docx" or not path.is_file():
        raise ValueError("结构定位工具仅支持已有的绝对路径 DOCX；旧 DOC 需先明确转换")
    if path.stat().st_size > 80_000_000: raise ValueError("DOCX 文件超过结构读取的 80 MB 限制")
    with ZipFile(path) as archive:
        if sum(info.file_size for info in archive.infolist()) > 120_000_000:
            raise ValueError("DOCX 解压体积超过结构读取限制")
        root = etree.fromstring(archive.read("word/document.xml"), parser=etree.XMLParser(resolve_entities=False, no_network=True))
    return root


def _paragraph_text(paragraph):
    return "".join(node.text or "" if node.tag == W + "t" else "\t" if node.tag == W + "tab" else "\n" for node in paragraph.iter() if node.tag in {W + "t", W + "tab", W + "br", W + "cr"})


def _editable(paragraph):
    # Targets containing hyperlinks, fields, drawings, revision or content
    # controls must not be flattened. They remain visible in the read result.
    return all(child.tag in {W + "pPr", W + "r"} and (child.tag != W + "r" or all(node.tag in {W + "rPr", W + "t"} for node in child)) for child in paragraph)


def _structure(root):
    locators = {}

    def blocks(parent, prefix):
        result = []
        p_index = t_index = 0
        for child in parent:
            if child.tag == W + "p":
                locator = f"{prefix}/p:{p_index}"
                locators[locator] = child
                result.append({"kind": "paragraph", "locator": locator, "text": _paragraph_text(child), "editable": _editable(child)})
                p_index += 1
            elif child.tag == W + "tbl":
                locator = f"{prefix}/t:{t_index}"
                rows = []
                vertical = {}
                for row_index, row in enumerate(child.findall(W + "tr")):
                    cells = []
                    before = row.find(f"{W}trPr/{W}gridBefore")
                    grid_column = int(before.get(W + "val", "0")) if before is not None else 0
                    for cell_index, cell in enumerate(row.findall(W + "tc")):
                        cell_locator = f"{locator}/r:{row_index}/c:{cell_index}"
                        span_node = cell.find(f"{W}tcPr/{W}gridSpan")
                        span = int(span_node.get(W + "val", "1")) if span_node is not None else 1
                        merge = cell.find(f"{W}tcPr/{W}vMerge")
                        continuation = merge is not None and merge.get(W + "val") != "restart"
                        merged_into = vertical.get(grid_column) if continuation else None
                        if merge is not None and not continuation:
                            vertical[grid_column] = cell_locator
                        elif merge is None:
                            vertical.pop(grid_column, None)
                        items = blocks(cell, cell_locator)
                        # Continuation cells are not independently editable.
                        if continuation:
                            for key, value in locators.items():
                                if key.startswith(cell_locator + "/"):
                                    value.set("zhai-merged-continuation", "true")
                            def protect(items):
                                for item in items:
                                    if item["kind"] == "paragraph": item["editable"] = False
                                    elif item["kind"] == "table":
                                        for row in item["rows"]:
                                            for cell in row["cells"]: protect(cell["blocks"])
                            protect(items)
                        cells.append({"locator": cell_locator, "gridColumn": grid_column, "columnSpan": span, "mergedInto": merged_into, "verticalContinuation": continuation, "blocks": items})
                        grid_column += span
                    rows.append({"index": row_index, "cells": cells})
                unsupported = [etree.QName(node).localname for node in child if node.tag not in {W + "tblPr", W + "tblGrid", W + "tr"}]
                result.append({"kind": "table", "locator": locator, "rows": rows, "unsupportedChildren": unsupported})
                t_index += 1
            elif child.tag not in {W + "sectPr", W + "tcPr"}:
                result.append({"kind":"unsupported", "xmlTag":etree.QName(child).localname, "textPreview":"".join(node.text or "" for node in child.iter(W + "t"))[:1000], "editable":False, "note":"此结构（如内容控件/修订）未展开定位，不能将未列出的内容判断为空或据此填表"})
        return result

    body = root.find(W + "body")
    if body is None: raise ValueError("DOCX 缺少正文结构")
    items = blocks(body, "body")
    return items, locators


def read_structure(path: Path):
    original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    root = _load(path)
    items, _ = _structure(root)
    if hashlib.sha256(path.read_bytes()).hexdigest() != original_hash: raise ValueError("读取期间材料已变化，请重新读取结构")
    result = {"schemaVersion": 1, "sourceHash": original_hash, "blocks": items, "notes": ["定位索引从 0 开始；合并格按实际 XML 单元格去重；受保护结构可读取但不可编辑。旧 DOC 不支持定位编辑。", "unsupported 块及表格 unsupportedChildren 不展开定位，不能把未列出的内容当作空值；需要人工确认。"]}
    if len(json.dumps(result, ensure_ascii=False).encode("utf-8")) > 1_500_000: raise ValueError("Word 结构超过单次工具读取上限，未返回不完整结构；请使用专用离线批处理能力")
    return result


def _replace_text(paragraph, old: str, text: str):
    nodes = list(paragraph.iter(W + "t"))
    if not nodes:
        run = paragraph.find(W + "r")
        if run is None: run = etree.SubElement(paragraph, W + "r")
        nodes = [etree.SubElement(run, W + "t")]
    start = 0
    while start < min(len(old), len(text)) and old[start] == text[start]: start += 1
    suffix = 0
    while suffix < min(len(old), len(text)) - start and old[-suffix - 1] == text[-suffix - 1]: suffix += 1
    stop = len(old) - suffix
    replacement = text[start:len(text) - suffix if suffix else len(text)]
    positions = []
    offset = 0
    for node in nodes:
        value = node.text or ""
        positions.append((node, value, offset, offset + len(value)))
        offset += len(value)
    anchor = next((node for node, _, _, end in positions if end > start), nodes[-1])
    for node, value, begin, end in positions:
        left = value[:max(0, min(len(value), start - begin))]
        right = value[max(0, min(len(value), stop - begin)):]
        if end <= start and node is not anchor: continue
        if begin >= stop and node is not anchor: continue
        node.text = left + (replacement if node is anchor else "") + right
        if node.text and (node.text[0].isspace() or node.text[-1].isspace()): node.set(XML_SPACE, "preserve")


def edit_structure(path: Path, output: Path, source_hash: str, edits: list):
    if not output.is_absolute() or output.suffix.lower() != ".docx" or output.exists() or output.resolve() == path.resolve():
        raise ValueError("编辑成果必须是尚不存在的新 DOCX 绝对路径，禁止覆盖原件或既有成果")
    original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    if not source_hash or source_hash != original_hash: raise ValueError("材料已变化或缺少 sourceHash，请重新读取结构后编辑")
    if not isinstance(edits, list) or not edits or len(edits) > 500: raise ValueError("edits 必须包含 1 至 500 个明确定位编辑")
    root = _load(path)
    _, locators = _structure(root)
    seen = set()
    changes = []
    for edit in edits:
        if not isinstance(edit, dict): raise ValueError("编辑项必须是对象")
        locator = edit.get("locator")
        if not isinstance(locator, str): raise ValueError("locator 必须是读取结构返回的字符串")
        if locator in seen: raise ValueError(f"同一定位重复编辑：{locator}")
        seen.add(locator)
        paragraph = locators.get(locator)
        if paragraph is None: raise ValueError(f"定位不存在：{locator}")
        if not _editable(paragraph) or paragraph.get("zhai-merged-continuation"):
            raise ValueError(f"定位含字段、链接、对象、修订、控件或合并续格，拒绝破坏格式：{locator}")
        old = _paragraph_text(paragraph)
        if edit.get("expected") != old: raise ValueError(f"定位原文与 expected 不符，拒绝猜测：{locator}")
        text = edit.get("text")
        if not isinstance(text, str) or any(c in text for c in "\n\r\t") or len(text) > 30_000:
            raise ValueError("段落新文字必须是单段纯文字，最多 30000 字符")
        changes.append((locator, paragraph, old, text))
    for _, paragraph, old, text in changes: _replace_text(paragraph, old, text)
    # Structure annotations are never serialized into the user's file.
    for paragraph in locators.values(): paragraph.attrib.pop("zhai-merged-continuation", None)
    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    created = False
    try:
        # Exclusive creation prevents accidental replacement even on collision.
        with output.open("xb") as stream:
            created = True
            with ZipFile(path) as source, ZipFile(stream, "w") as target:
                for item in source.infolist():
                    target.writestr(item, payload if item.filename == "word/document.xml" else source.read(item.filename))
        reopened = _load(output)
        _, reopened_locators = _structure(reopened)
        if set(locators) != set(reopened_locators): raise ValueError("成果重开后结构定位变化")
        expected_values = {locator: text for locator, _, _, text in changes}
        for locator, paragraph in locators.items():
            if _paragraph_text(reopened_locators[locator]) != expected_values.get(locator, _paragraph_text(paragraph)):
                raise ValueError("成果重开后文字核验失败")
        # Also reopen with the office reader used by the rest of the engine.
        from docx import Document
        Document(output)
        if hashlib.sha256(path.read_bytes()).hexdigest() != original_hash: raise ValueError("原件保护核验失败")
    except Exception:
        if created: output.unlink(missing_ok=True)
        raise
    return {"summary": f"已按定位编辑 {len(edits)} 处；成果重新打开核验通过，原件哈希不变；新文字沿用替换起点的字符格式。", "outputs": [str(output)], "edited": [locator for locator, *_ in changes]}


def run_request(request_path: Path):
    request = json.loads(request_path.read_text(encoding="utf-8"))
    if request.get("schemaVersion") != 1 or request.get("neverOverwrite") is not True: raise ValueError("Word 结构请求要求 schemaVersion=1、neverOverwrite=true")
    path = Path(request["inputPath"])
    if request.get("mode") == "read": return read_structure(path)
    if request.get("mode") == "edit": return edit_structure(path, Path(request["outputPath"]), request.get("sourceHash", ""), request.get("edits"))
    raise ValueError("Word 结构 mode 必须是 read 或 edit")
