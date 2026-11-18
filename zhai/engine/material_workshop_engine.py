#!/usr/bin/env python3
"""材料助手隔离处理引擎。

入口只接受 Rust 层写出的 JSON 请求文件。stdout 严格使用 NDJSON 事件，
便于 Tauri 主进程校验并转发；诊断信息只写 stderr。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import io
import io
import inspect
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from contextlib import redirect_stderr, redirect_stdout
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Iterable

if sys.platform.startswith("linux"):
    os.environ.setdefault("OPENSSL_armcap", "0")

from redaction_engine import redact_text, redaction_replacements


SUPPORTED_EXTENSIONS = {"docx", "xlsx", "pdf", "csv", "md", "txt"}
LEGACY_EXTENSIONS = {"doc", "xls"}
PDF_OPERATIONS = {
    "safe_copy", "inspect_document", "review_content", "cross_file_consistency_check",
    "merge_or_split", "split_pdf", "rotate_pdf", "watermark_pdf",
    "extract_pdf_text", "ocr_pdf", "clean_metadata", "redact_content",
    "convert_to_pdf",  # An existing PDF can pass through a mixed conversion batch.
}


def validate_pdf_operations(operation_kinds: set[str]) -> None:
    unsupported = operation_kinds - PDF_OPERATIONS
    if unsupported:
        raise ValueError("PDF 不支持以下操作，未生成成果：" + ", ".join(sorted(unsupported)))
    if "inspect_document" in operation_kinds and operation_kinds - {"inspect_document", "safe_copy", "review_content", "cross_file_consistency_check"}:
        raise ValueError("PDF 检查与文件修改不能在同一步执行，请分开运行")


def emit(event: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def read_request(path: Path) -> dict[str, Any]:
    if not path.is_absolute():
        raise ValueError("请求文件必须使用绝对路径")
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schemaVersion") != 1:
        raise ValueError("不支持的请求协议版本")
    if data.get("neverOverwrite") is not True:
        raise ValueError("安全策略要求 neverOverwrite=true")
    for key in ("jobId", "files", "operations", "outputDirectory", "reportDirectory"):
        if key not in data:
            raise ValueError(f"请求缺少字段：{key}")
    return data


def ensure_inside(candidate: Path, parent: Path) -> Path:
    candidate_resolved = candidate.resolve(strict=False)
    parent_resolved = parent.resolve(strict=False)
    if candidate_resolved != parent_resolved and parent_resolved not in candidate_resolved.parents:
        raise ValueError(f"拒绝写入输出目录以外的路径：{candidate}")
    return candidate_resolved


def safe_relative_path(value: str, fallback_name: str) -> Path:
    raw = Path(value)
    pieces = [part for part in raw.parts if part not in ("", ".", "..", os.sep)]
    if not pieces:
        return Path(fallback_name)
    return Path(*pieces)


def unique_output_path(output_root: Path, relative: Path) -> Path:
    target = ensure_inside(output_root / relative, output_root)
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    for index in range(2, 10_000):
        candidate = target.with_name(f"{stem}-{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法为输出文件生成唯一名称：{target.name}")


def atomic_writer(target: Path, write: Callable[[Path], None]) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{uuid.uuid4().hex}.part")
    try:
        write(temporary)
        with temporary.open("rb") as stream:
            os.fsync(stream.fileno())
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


def copy_atomic(source: Path, target: Path) -> None:
    atomic_writer(target, lambda temporary: shutil.copy2(source, temporary))


def checksum(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def import_docx() -> Any:
    try:
        import docx
    except ImportError as exc:
        raise RuntimeError("DOCX 组件未安装") from exc
    return docx


def _set_run_font(run: Any, east_asia: str, size_pt: float, bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()  # noqa: SLF001
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), east_asia)


def _fix_chinese_quotes(value: str) -> str:
    output: list[str] = []
    opening = True
    for character in value:
        if character == '"':
            output.append("“" if opening else "”")
            opening = not opening
        else:
            output.append(character)
    return "".join(output)


def _parse_gongwen_markdown(value: str) -> tuple[str, str, list[tuple[str, str]]]:
    title = ""
    author = ""
    sections: list[tuple[str, str]] = []
    for raw_line in value.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("# ") and not title:
            title = line[2:].strip()
        elif line.startswith("#### "):
            sections.append(("heading3", _fix_chinese_quotes(line[5:].strip())))
        elif line.startswith("### "):
            sections.append(("heading2", _fix_chinese_quotes(line[4:].strip())))
        elif line.startswith("## "):
            sections.append(("heading1", _fix_chinese_quotes(line[3:].strip())))
        elif line.startswith("作者："):
            author = line.removeprefix("作者：").strip()
        else:
            sections.append(("body", _fix_chinese_quotes(line)))
    if not title:
        raise ValueError("未找到标题；请使用“# 标题”开头")
    if not sections:
        raise ValueError("未找到正文内容")
    return _fix_chinese_quotes(title), _fix_chinese_quotes(author), sections


@lru_cache(maxsize=1)
def _installed_font_families() -> str:
    if sys.platform == 'darwin':
        try:
            import ctypes as ct
            cf = ct.CDLL('/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation')
            core = ct.CDLL('/System/Library/Frameworks/CoreText.framework/CoreText')
            core.CTFontManagerCopyAvailableFontFamilyNames.restype = ct.c_void_p
            cf.CFArrayGetCount.argtypes = [ct.c_void_p]; cf.CFArrayGetCount.restype = ct.c_long
            cf.CFArrayGetValueAtIndex.argtypes = [ct.c_void_p, ct.c_long]; cf.CFArrayGetValueAtIndex.restype = ct.c_void_p
            cf.CFStringGetCString.argtypes = [ct.c_void_p, ct.c_char_p, ct.c_long, ct.c_uint32]
            cf.CFRelease.argtypes = [ct.c_void_p]
            families = core.CTFontManagerCopyAvailableFontFamilyNames()
            names = []
            try:
                for i in range(cf.CFArrayGetCount(families)):
                    buffer = ct.create_string_buffer(2048)
                    if cf.CFStringGetCString(cf.CFArrayGetValueAtIndex(families,i), buffer, 2048, 0x08000100):
                        names.append(buffer.value.decode('utf-8'))
            finally:
                if families: cf.CFRelease(families)
            if names: return _font_token('\n'.join(names))
        except (OSError, ValueError, AttributeError):
            pass
    commands = []
    if shutil.which("fc-list"):
        commands.append(["fc-list", ":", "family"])
    if sys.platform == "darwin":
        commands.append(["system_profiler", "SPFontsDataType"])
    for command in commands:
        try:
            completed = subprocess.run(command, capture_output=True, text=True, timeout=12, check=False)
            if completed.returncode == 0 and completed.stdout:
                return completed.stdout.lower().replace("_", "").replace(" ", "")
        except (OSError, subprocess.TimeoutExpired):
            continue
    return ""


def _font_token(value: str) -> str:
    return value.lower().replace("_", "").replace(" ", "").replace("-", "")


def _font_is_installed(font: str, installed: str) -> bool:
    # These are naming variants for the same font family, not visual fallbacks.
    aliases = {
        "仿宋_GB2312": ["仿宋gb2312", "simfang", "fangsonggb2312"],
        "楷体_GB2312": ["楷体gb2312", "simkai", "kaitigb2312"],
        "宋体": ["宋体", "simsun"],
        "方正小标宋简体": ["方正小标宋简体", "fzxiaobiaosong", "fzxbs"],
        "黑体": ["黑体", "simhei"],
        "Times New Roman": ["timesnewroman"],
    }
    return any(_font_token(alias) in installed for alias in aliases.get(font, [font]))


def document_font_names(source: Path) -> set[str]:
    names: set[str] = set()
    try:
        if source.suffix.lower() == ".docx":
            from zipfile import ZipFile
            from lxml import etree
            from docx.oxml.ns import qn
            with ZipFile(source) as package:
                styles_used = set()
                for name in package.namelist():
                    if name.startswith('word/') and name.endswith('.xml') and not name.endswith(('styles.xml','fontTable.xml')):
                        parsed = etree.fromstring(package.read(name))
                        styles_used.update(e.get(qn('w:val')) for e in parsed.iter() if e.tag in {qn('w:pStyle'),qn('w:rStyle'),qn('w:tblStyle')})
                for part in package.namelist():
                    if not part.startswith('word/') or not part.endswith('.xml') or part.endswith('fontTable.xml'):
                        continue
                    root = etree.fromstring(package.read(part))
                    if part == 'word/styles.xml':
                        styles = {s.get(qn('w:styleId')): s for s in root if s.tag == qn('w:style')}
                        styles_used.update(k for k, s in styles.items() if s.get(qn('w:default')) == '1')
                        pending = list(styles_used)
                        while pending:
                            style = styles.get(pending.pop())
                            parent = style.find(qn('w:basedOn')) if style is not None else None
                            parent_id = parent.get(qn('w:val')) if parent is not None else None
                            if parent_id and parent_id not in styles_used:
                                styles_used.add(parent_id)
                                pending.append(parent_id)
                        for style in list(root):
                            if style.tag == qn('w:style') and style.get(qn('w:styleId')) not in styles_used and style.get(qn('w:default')) != '1':
                                root.remove(style)
                    for element in root.iter():
                        if element.tag.endswith('}rFonts'):
                            names.update(value.split(';')[0] for key, value in element.attrib.items() if key.rsplit('}',1)[-1] in {'ascii','hAnsi','eastAsia','cs'})
                        if element.tag.endswith('}textpath'):
                            match = re.search(r'font-family:["\']?([^;"\']+)', element.get('style',''))
                            if match: names.add(match.group(1))
            document = import_docx().Document(source)
            for paragraph in iter_docx_paragraphs(document):
                for run in paragraph.runs:
                    if run.font.name:
                        names.add(run.font.name)
                    fonts = run._element.rPr.rFonts if run._element.rPr is not None else None  # noqa: SLF001
                    if fonts is not None:
                        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                            value = fonts.get(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{key}")
                            if value:
                                names.add(value)
        elif source.suffix.lower() == ".xlsx":
            workbook = import_openpyxl().load_workbook(source, read_only=False, data_only=False)
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.font and cell.font.name:
                            names.add(cell.font.name)
            workbook.close()
    except Exception:  # Font inspection must never corrupt or block document reading.
        return set()
    return {name.split(';')[0].strip() for name in names if name and name.strip()}


def missing_document_fonts(source: Path) -> list[str]:
    installed = _installed_font_families()
    if not installed:
        return []
    return sorted(name for name in document_font_names(source) if not _font_is_installed(name, installed))


def font_notices(source: Path) -> list[str]:
    installed = _installed_font_families()
    if not installed:
        return ['字体检查：未能读取本机字体清单，实际显示字体需在办公软件中确认']
    missing = sorted(name for name in document_font_names(source) if not _font_is_installed(name, installed))
    if not missing:
        return []
    return ['字体提示：本机缺少 ' + '、'.join(missing) + '。文件中保留原字体名称；打开或导出时办公软件可能替换字体，实际替代字体尚未确认。']


def generate_gongwen_docx(source: Path, target: Path, parameters: dict[str, Any]) -> dict[str, Any]:
    docx = import_docx()
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    # python-docx 1.2 从 `docx/parts/../templates` 加载新页脚模板。
    # PyInstaller 会收集 templates，但纯 Python 的 parts 目录不会物理解包，
    # 导致操作系统无法解析中间的 `parts/..`。只在 sidecar 临时解包目录
    # 创建空的中间目录，不写入用户文件系统。
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        (Path(bundle_root) / "docx" / "parts").mkdir(parents=True, exist_ok=True)

    value = source.read_text(encoding=str(parameters.get("encoding") or "utf-8-sig"))
    title, author, sections = _parse_gongwen_markdown(value)
    document = docx.Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)
    section.footer_distance = Cm(1.75)

    settings = document.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is None:
        settings.append(OxmlElement("w:evenAndOddHeaders"))

    def configure_paragraph(paragraph: Any, alignment: Any, first_line_pt: float | None = None) -> None:
        paragraph.alignment = alignment
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Pt(28)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if first_line_pt is not None:
            paragraph.paragraph_format.first_line_indent = Pt(first_line_pt)

    def add_text_paragraph(text: str, font: str, size: float = 16, bold: bool = False, align: Any = WD_ALIGN_PARAGRAPH.JUSTIFY, indent: bool = False) -> Any:
        paragraph = document.add_paragraph()
        configure_paragraph(paragraph, align, size * 2 if indent else None)
        _set_run_font(paragraph.add_run(text), font, size, bold)
        return paragraph

    for _ in range(2):
        paragraph = document.add_paragraph()
        configure_paragraph(paragraph, WD_ALIGN_PARAGRAPH.LEFT)
    add_text_paragraph(title, "方正小标宋简体", 22, False, WD_ALIGN_PARAGRAPH.CENTER)
    configure_paragraph(document.add_paragraph(), WD_ALIGN_PARAGRAPH.LEFT)
    if author:
        add_text_paragraph(author, "楷体_GB2312", 16, False, WD_ALIGN_PARAGRAPH.CENTER)
        configure_paragraph(document.add_paragraph(), WD_ALIGN_PARAGRAPH.LEFT)
    for kind, text in sections:
        if kind == "heading1":
            add_text_paragraph(text, "黑体", 16, False, indent=True)
        elif kind == "heading2":
            add_text_paragraph(text, "楷体_GB2312", 16, True, indent=True)
        elif kind == "heading3":
            add_text_paragraph(text, "仿宋_GB2312", 16, True, indent=True)
        else:
            add_text_paragraph(text, "仿宋_GB2312", 16, False, indent=True)

    def make_footer(footer: Any, right: bool) -> None:
        footer.is_linked_to_previous = False
        paragraph = footer.paragraphs[0]
        configure_paragraph(paragraph, WD_ALIGN_PARAGRAPH.RIGHT if right else WD_ALIGN_PARAGRAPH.LEFT)
        paragraph.paragraph_format.right_indent = Pt(14) if right else None
        paragraph.paragraph_format.left_indent = Pt(14) if not right else None
        _set_run_font(paragraph.add_run("— "), "宋体", 14)
        run = paragraph.add_run()
        _set_run_font(run, "宋体", 14)
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = " PAGE "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.extend([begin, instruction, separate, end])  # noqa: SLF001
        _set_run_font(paragraph.add_run(" —"), "宋体", 14)

    make_footer(section.footer, True)
    make_footer(section.even_page_footer, False)
    clear_docx_metadata(document)
    atomic_writer(target, lambda temporary: document.save(temporary))
    reopened = docx.Document(target)
    required_fonts = ["方正小标宋简体", "仿宋_GB2312", "楷体_GB2312", "黑体", "宋体", "Times New Roman"]
    installed = _installed_font_families()
    missing = [font for font in required_fonts if installed and not _font_is_installed(font, installed)]
    notes = ["已按指定字体名写入 DOCX，未自动替换字体"]
    if missing:
        notes.append("当前电脑可能缺少字体：" + "、".join(missing))
    elif not installed:
        notes.append("未能读取系统字体清单，请在目标电脑人工检查字体")
    return {
        "changes": len(reopened.paragraphs),
        "notes": notes,
        "verification": f"docx_reopened;paragraphs={len(reopened.paragraphs)};even_footer=enabled",
    }


def import_openpyxl() -> Any:
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError("XLSX 组件未安装") from exc
    return openpyxl


def import_pypdf() -> Any:
    try:
        import pypdf
    except ImportError as exc:
        raise RuntimeError("PDF 组件未安装") from exc
    return pypdf


def iter_docx_paragraphs(document: Any) -> Iterable[Any]:
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in document.sections:
        # 访问一个尚不存在且“链接到前一节”的页眉/页脚会让 python-docx
        # 创建新部件。处理现有材料时只遍历真实存在的独立部件。
        if not section.header.is_linked_to_previous:
            yield from section.header.paragraphs
        if not section.footer.is_linked_to_previous:
            yield from section.footer.paragraphs


def replace_in_runs(paragraph: Any, replacements: list[tuple[str, str]]) -> int:
    """优先逐 run 替换，跨 run 匹配时再回退为段落级替换。"""
    total = 0
    for run in paragraph.runs:
        text = run.text
        for before, after in replacements:
            if before:
                text, count = re.subn(re.escape(before), after, text)
                total += count
        run.text = text
    full_text = paragraph.text
    transformed = full_text
    for before, after in replacements:
        if before:
            transformed, count = re.subn(re.escape(before), after, transformed)
            total += count
    if transformed != full_text and paragraph.runs:
        paragraph.runs[0].text = transformed
        for run in paragraph.runs[1:]:
            run.text = ""
    return total


def normalize_docx(document: Any) -> int:
    changed = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if re.match(r"^[一二三四五六七八九十]+、", text) and paragraph.style.name != "Heading 1":
            try:
                paragraph.style = document.styles["Heading 1"]
                changed += 1
            except KeyError:
                pass
        paragraph.paragraph_format.space_after = None
    return changed


def normalize_gongwen_attachment_lists(paragraphs: list[Any], preserve_bold: bool = True) -> int:
    """Align explicit attachment lists without confusing them with body headings.

    Normalize only consecutive numbered items introduced by 附件：1. .  Clear
    character-based indents as well as point indents; Word gives the former
    priority. Replace padding spaces with a tab and hanging indentation so
    wrapped names remain aligned. Unrelated numbered paragraphs stay untouched.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from copy import deepcopy
    from gongwen_layout import complex_paragraph

    def remove_inherited_indent(paragraph: Any) -> None:
        # Merely deleting direct firstLineChars leaves the same attribute in
        # Normal Indent (or another base style) active. Flatten this paragraph's
        # style into a private derivative, preserving the other style settings.
        original = paragraph.style
        styles = paragraph.part.document.styles
        name = "zhai附件说明-" + original.style_id
        if original.name.startswith("zhai附件说明-"):
            return
        if name not in styles:
            derived = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            chain = []
            seen = set()
            current = original
            while current is not None and current.style_id not in seen:
                seen.add(current.style_id)
                chain.append(current)
                current = current.base_style
            for source_style in reversed(chain):
                for part in ('pPr', 'rPr'):
                    source = source_style.element.find(qn('w:' + part))
                    if source is None:
                        continue
                    destination = getattr(derived.element, 'get_or_add_' + part)()
                    for child in source:
                        if part == 'pPr' and child.tag in (qn('w:ind'), qn('w:tabs'), qn('w:numPr')):
                            continue
                        existing = destination.find(child.tag)
                        if existing is not None:
                            destination.remove(existing)
                        destination.append(deepcopy(child))
        paragraph.style = styles[name]

    first_pattern = re.compile(r"^\s*(附件\s*[：:]\s*)(1)([.．、])[ \t\u3000]*(?=\S)")
    next_pattern = re.compile(r"^[ \t\u3000]*(\d{1,3})([.．、])[ \t\u3000]*(?=\S)")
    changed = 0
    index = 0
    while index < len(paragraphs):
        first = first_pattern.match(paragraphs[index].text)
        if first is None:
            index += 1
            continue
        items = [(paragraphs[index], first, "附件" + first.group(1).strip()[-1] + "1" + first.group(3))]
        end = index + 1
        while end < len(paragraphs):
            paragraph = paragraphs[end]
            if paragraphs[end - 1]._p.getnext() is not paragraph._p:
                break
            match = next_pattern.match(paragraph.text)
            if match is None or int(match.group(1)) != len(items) + 1:
                break
            items.append((paragraph, match, match.group(1) + match.group(2)))
            end += 1
        # Fields, links, automatic numbering and manual line breaks need their
        # own structural handling. Do not flatten them to fix a plain-text list.
        if any(
            complex_paragraph(p) or "\n" in p.text or p._p.xpath('./w:pPr/w:numPr')
            or ''.join(r.text for r in p.runs) != p.text
            or p._p.xpath('.//w:drawing | .//w:pict | .//w:fldChar')
            for p, _, _ in items
        ):
            index = end
            continue
        name_offset = 80 + max(16, (len(str(len(items))) + 1) * 8)
        for position, (paragraph, match, prefix) in enumerate(items):
            remove_inherited_indent(paragraph)
            # Remove only the old prefix, retaining all content runs and styles.
            remaining = match.end()
            inserted = False
            for run in paragraph.runs:
                if remaining <= 0:
                    break
                old = run.text
                consumed = min(remaining, len(old))
                run.text = (prefix + "\t" if not inserted else "") + old[consumed:]
                inserted = True
                remaining -= consumed
            ppr = paragraph._p.get_or_add_pPr()
            for ind in list(ppr.findall(qn("w:ind"))):
                ppr.remove(ind)
            for tabs in list(ppr.findall(qn("w:tabs"))):
                ppr.remove(tabs)
            fmt = paragraph.paragraph_format
            fmt.left_indent = Pt(name_offset)
            fmt.right_indent = Pt(0)
            fmt.first_line_indent = Pt((32 if position == 0 else 80) - name_offset)
            fmt.tab_stops.add_tab_stop(Pt(name_offset))
            fmt.keep_together = len(paragraph.text) <= 120
            fmt.keep_with_next = (len(items) <= 4 and sum(len(p.text) for p, _, _ in items) <= 240 and position < len(items) - 1)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if not preserve_bold:
                for run in paragraph.runs:
                    run.bold = False
            changed += 1
        index = end
    return changed


def format_existing_docx_as_gongwen(
    document: Any,
    parameters: dict[str, Any] | None = None,
) -> tuple[int, list[str]]:
    from gongwen_layout import apply_layout
    # Font notices are collected from the saved package by process_file,
    # including inherited styles, headers and WordArt, not a second partial scan.
    return apply_layout(document, parameters, _set_run_font, normalize_gongwen_attachment_lists)


def clear_docx_metadata(document: Any) -> None:
    properties = document.core_properties
    properties.author = ""
    properties.last_modified_by = ""
    properties.comments = ""
    properties.keywords = ""
    properties.subject = ""


def normalize_whitespace_text(value: str) -> tuple[str, int]:
    updated, changes = re.subn(r"[ \t]{2,}", " ", value)
    updated, chinese_changes = re.subn(
        r"(?<=[\u3400-\u9fff]) (?=[\u3400-\u9fff])",
        "",
        updated,
    )
    return updated, changes + chinese_changes


def remove_paragraph(paragraph: Any) -> None:
    element = paragraph._element  # noqa: SLF001 - python-docx has no public delete API
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def wps_bridge_path() -> Path | None:
    candidates: list[Path] = []
    configured = os.environ.get("MATERIAL_ASSISTANT_WPS_BRIDGE")
    if configured:
        candidates.append(Path(configured))
    executable_directory = Path(sys.executable).resolve().parent
    candidates.extend(
        [
            executable_directory / "material-wps-bridge-aarch64-apple-darwin",
            executable_directory.parent
            / "Resources/binaries/material-wps-bridge-aarch64-apple-darwin",
        ]
    )
    if sys.platform.startswith("linux"):
        candidates.extend(
            [
                executable_directory / "material-wps-bridge-aarch64-unknown-linux-gnu",
                Path("/usr/lib/zhai/material-wps-bridge"),
                Path("/usr/local/libexec/zhai/material-wps-bridge"),
                Path.home() / ".local/libexec/zhai/material-wps-bridge",
            ]
        )
        resolved = shutil.which("material-wps-bridge")
        if resolved:
            candidates.append(Path(resolved))
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def libreoffice_path() -> Path | None:
    candidates: list[Path] = []
    configured = os.environ.get("MATERIAL_WORKSHOP_LIBREOFFICE")
    if configured:
        candidates.append(Path(configured))
    if sys.platform == "darwin":
        candidates.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    elif os.name == "nt":
        for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
            if os.environ.get(variable):
                candidates.append(Path(os.environ[variable]) / "LibreOffice/program/soffice.exe")
    for name in ("libreoffice", "soffice"):
        resolved = shutil.which(name)
        if resolved:
            candidates.append(Path(resolved))
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def cjk_character_count(value: str) -> int:
    return sum("\u3400" <= character <= "\u9fff" for character in value)


def validate_pdf_conversion(source: Path, target: Path) -> tuple[int, str]:
    reader = import_pypdf().PdfReader(target)
    page_count = len(reader.pages)
    if page_count == 0:
        raise RuntimeError("PDF 转换结果没有页面")
    extracted_pages = [(page.extract_text() or "").strip() for page in reader.pages]
    source_text = extract_text_preview(source)
    source_cjk = cjk_character_count(source_text)
    output_cjk = cjk_character_count("\n".join(extracted_pages))
    extracted_text = "\n".join(extracted_pages)
    if "�" in extracted_text:
        raise RuntimeError("PDF 文字校验失败：发现 Unicode 替换字符，可能存在字体或编码问题")
    if source.suffix.lower() == ".docx" and source_cjk >= 20:
        retention = output_cjk / source_cjk
        if retention < 0.80:
            raise RuntimeError(
                "PDF 中文保真校验失败：可提取中文仅保留 "
                f"{retention:.0%}，已拒绝该转换结果"
            )
    nonempty_pages = sum(bool(value) for value in extracted_pages)
    if source.suffix.lower() == ".docx" and source_text.strip() and nonempty_pages == 0:
        raise RuntimeError("PDF 视觉前置校验失败：所有页面均未提取到正文")
    return page_count, f"cjk={output_cjk}/{source_cjk};nonempty={nonempty_pages}/{page_count}"


def convert_office_with_wps(source: Path, target: Path, action: str) -> dict[str, Any]:
    bridge = wps_bridge_path()
    if bridge is None:
        raise RuntimeError("WPS 已安装，但材料助手的官方 RPC/加载项桥接尚未安装或验证")
    temporary_root = target.parent / f".convert-{uuid.uuid4().hex}"
    temporary_root.mkdir(parents=True, exist_ok=False)
    try:
        generated = temporary_root / target.name
        command = [str(bridge), "convert", "--action", action, "--input", str(source), "--output", str(generated)]
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
        if completed.returncode != 0 or not generated.is_file():
            detail = (completed.stderr or completed.stdout or "WPS 桥接未生成结果").strip()
            raise RuntimeError(f"WPS 转换失败：{detail.splitlines()[0]}")
        if action in {"docx-to-pdf", "xlsx-to-pdf"}:
            page_count, verification = validate_pdf_conversion(source, generated)
            notes = [f"WPS 导出 PDF 共 {page_count} 页；中文与空白页校验通过"]
        elif action == "doc-to-docx":
            document = import_docx().Document(generated)
            verification = f"docx_reopened;paragraphs={len(document.paragraphs)}"
            notes = ["WPS DOC→DOCX 转换后已重新打开校验"]
        elif action in {"xls-to-xlsx", "xlsx-recalculate"}:
            workbook = import_openpyxl().load_workbook(generated, read_only=True, data_only=False)
            verification = f"xlsx_reopened;sheets={len(workbook.sheetnames)}"
            workbook.close()
            notes = ["WPS 工作簿输出已重新打开校验"]
        else:
            raise RuntimeError(f"不支持的 WPS 桥接动作：{action}")
        atomic_writer(target, lambda temporary: shutil.copy2(generated, temporary))
        return {
            "changes": 1,
            "notes": notes,
            "verification": verification,
        }
    finally:
        shutil.rmtree(temporary_root, ignore_errors=True)


def convert_office_with_libreoffice(source: Path, target: Path, action: str) -> dict[str, Any]:
    executable = libreoffice_path()
    if executable is None:
        raise RuntimeError("未检测到 LibreOffice；请先安装或在设置中指定 soffice 路径")
    filters = {
        "doc-to-docx": ("docx:Office Open XML Text", ".docx"),
        "doc-to-pdf": ("pdf:writer_pdf_Export", ".pdf"),
        "xls-to-xlsx": ("xlsx:Calc MS Excel 2007 XML", ".xlsx"),
        "xlsx-recalculate": ("xlsx:Calc MS Excel 2007 XML", ".xlsx"),
        "docx-to-pdf": ("pdf:writer_pdf_Export", ".pdf"),
        "xlsx-to-pdf": ("pdf:calc_pdf_Export", ".pdf"),
    }
    if action not in filters:
        raise RuntimeError(f"LibreOffice 不支持该转换动作：{action}")
    filter_name, suffix = filters[action]
    temporary_root = target.parent / f".libreoffice-{uuid.uuid4().hex}"
    input_root = temporary_root / "input"
    output_root = temporary_root / "output"
    profile_root = temporary_root / "profile"
    input_root.mkdir(parents=True, exist_ok=False)
    output_root.mkdir()
    profile_root.mkdir()
    staged_source = input_root / source.name
    shutil.copy2(source, staged_source)
    try:
        command = [
            str(executable),
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            f"-env:UserInstallation={profile_root.resolve().as_uri()}",
            "--convert-to",
            filter_name,
            "--outdir",
            str(output_root),
            str(staged_source),
        ]
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=180,
        )
        generated = output_root / f"{staged_source.stem}{suffix}"
        if completed.returncode != 0 or not generated.is_file():
            detail = (completed.stderr or completed.stdout or "LibreOffice 未生成结果").strip()
            raise RuntimeError(f"LibreOffice 转换失败：{detail.splitlines()[0] if detail else '未生成结果'}")
        if action in {"doc-to-pdf", "docx-to-pdf", "xlsx-to-pdf"}:
            page_count, verification = validate_pdf_conversion(source, generated)
            notes = [f"LibreOffice 导出 PDF 共 {page_count} 页；中文文本与空白页校验通过"]
            missing_fonts = missing_document_fonts(source)
            if missing_fonts:
                notes.append("转换前检测到缺失字体，LibreOffice 可能使用替代字体：" + "、".join(missing_fonts))
        elif action == "doc-to-docx":
            document = import_docx().Document(generated)
            verification = f"docx_reopened;paragraphs={len(document.paragraphs)}"
            notes = ["LibreOffice DOC→DOCX 转换后已重新打开校验"]
        else:
            workbook = import_openpyxl().load_workbook(generated, read_only=True, data_only=False)
            verification = f"xlsx_reopened;sheets={len(workbook.sheetnames)}"
            workbook.close()
            notes = ["LibreOffice 工作簿输出已重新打开校验"]
        atomic_writer(target, lambda temporary: shutil.copy2(generated, temporary))
        return {"changes": 1, "notes": notes, "verification": verification}
    finally:
        shutil.rmtree(temporary_root, ignore_errors=True)


def convert_office(
    source: Path,
    target: Path,
    action: str,
    parameters: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Convert with a verified WPS bridge first and LibreOffice as an explicit fallback."""
    parameters = parameters or {}
    preference = str(parameters.get("adapterPreference") or "auto").lower()
    if preference not in {"auto", "wps", "libreoffice"}:
        raise RuntimeError(f"未知的办公转换器偏好：{preference}")

    if preference == "wps":
        return convert_office_with_wps(source, target, action)
    if preference == "libreoffice":
        return convert_office_with_libreoffice(source, target, action)

    failures: list[str] = []
    if wps_bridge_path() is not None:
        try:
            return convert_office_with_wps(source, target, action)
        except Exception as exc:  # noqa: BLE001 - retain adapter diagnostics for fallback
            failures.append(f"WPS：{exc}")
    if libreoffice_path() is not None:
        result = convert_office_with_libreoffice(source, target, action)
        if failures:
            result.setdefault("notes", []).insert(
                0,
                "WPS 优先转换未通过，已明确回退 LibreOffice：" + failures[-1],
            )
        else:
            result.setdefault("notes", []).insert(
                0,
                "当前未检测到已验证的 WPS 桥接，已使用 LibreOffice",
            )
        return result
    if failures:
        raise RuntimeError("WPS 转换失败且没有 LibreOffice 回退：" + failures[-1])
    raise RuntimeError("没有可用的办公转换器；请配置已验证的 WPS 桥接，或安装 LibreOffice 作为回退")


def process_docx(
    source: Path,
    target: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> dict[str, Any]:
    docx = import_docx()
    document = docx.Document(source)
    layout_only = not ai_reviews and operation_kinds <= {"normalize_docx_styles", "normalize_whitespace", "clean_metadata"}
    def content_signature(doc: Any) -> tuple[Any, ...]:
        from lxml import etree
        texts = tuple(re.sub(r'\s+', '', ''.join(e.itertext())) for e in doc._element.xpath('.//w:t'))
        # Compare text as a stream: harmless run splits and attachment tabs may change.
        objects = tuple(etree.tostring(e, method='c14n') for e in doc._element.xpath('.//w:drawing | .//w:pict | .//w:hyperlink | .//w:fldSimple | .//w:bookmarkStart | .//w:bookmarkEnd'))
        rels = tuple(sorted((r.rId, r.reltype, r.target_ref) for r in doc.part.rels.values() if not r.reltype.endswith('/footer')))
        return ''.join(texts), objects, rels, len(doc.tables)
    original_content = content_signature(document) if layout_only else None
    changes = 0
    notes: list[str] = []
    paragraphs = list(iter_docx_paragraphs(document))
    paragraph_contexts: dict[int, str] = {}

    def collect_table_contexts(table: Any) -> None:
        for row in table.rows:
            for column_index, cell in enumerate(row.cells):
                context_parts: list[str] = []
                if column_index > 0:
                    context_parts.append(row.cells[column_index - 1].text)
                for paragraph in cell.paragraphs:
                    paragraph_contexts[id(paragraph._element)] = " ".join(context_parts)  # noqa: SLF001
                for nested in cell.tables:
                    collect_table_contexts(nested)

    for table in document.tables:
        collect_table_contexts(table)

    for review in ai_reviews:
        locator = review.get("locator") or {}
        if locator.get("kind") != "docx_paragraph":
            continue
        index = int(locator.get("index", -1))
        if index < 0 or index >= len(paragraphs):
            notes.append(f"AI 建议定位失效：段落 {index}")
            continue
        before = str(review.get("before", ""))
        after = str(review.get("after", ""))
        if before and before in paragraphs[index].text:
            changes += replace_in_runs(paragraphs[index], [(before, after)])
        else:
            notes.append(f"AI 建议未应用：段落 {index} 的原文已变化")

    if "replace_text" in operation_kinds or "update_dates" in operation_kinds:
        params = operation_parameters.get("replace_text", {})
        replacements = [
            (str(item.get("before", "")), str(item.get("after", "")))
            for item in params.get("replacements", [])
            if item.get("before")
        ]
        for paragraph in paragraphs:
            changes += replace_in_runs(paragraph, replacements)
        if not replacements:
            notes.append("未配置具体替换对，文件仅生成安全副本")

    if "redact_content" in operation_kinds:
        for paragraph in paragraphs:
            context = paragraph_contexts.get(id(paragraph._element), "")  # noqa: SLF001
            replacements = redaction_replacements(paragraph.text, context)
            changes += replace_in_runs(paragraph, replacements)

    if "normalize_whitespace" in operation_kinds:
        params = operation_parameters.get("normalize_whitespace", {})
        from gongwen_layout import clean_whitespace
        cleaned, clean_notes = clean_whitespace(document, params.get("collapseSpaces", True), params.get("removeEmptyParagraphs", True))
        changes += cleaned
        notes.extend(clean_notes)

    if "remove_line_breaks" in operation_kinds:
        params = operation_parameters.get("remove_line_breaks", {})
        for paragraph in list(iter_docx_paragraphs(document)):
            for run in paragraph.runs:
                updated, count = re.subn(r"[\r\n]+", "", run.text)
                if count:
                    run.text = updated
                    changes += count
        if params.get("joinParagraphs", False):
            body_paragraphs = list(document.paragraphs)
            for paragraph in body_paragraphs[1:]:
                previous = paragraph._element.getprevious()  # noqa: SLF001
                if previous is not None and previous.tag.endswith("}p"):
                    previous_paragraph = next(
                        (item for item in document.paragraphs if item._element is previous),  # noqa: SLF001
                        None,
                    )
                    if previous_paragraph is not None and paragraph.text:
                        previous_paragraph.add_run(paragraph.text)
                        remove_paragraph(paragraph)
                        changes += 1

    if "normalize_docx_styles" in operation_kinds:
        style_parameters = operation_parameters.get("normalize_docx_styles", {})
        if style_parameters.get("preset") == "gongwen":
            formatted, format_notes = format_existing_docx_as_gongwen(
                document,
                style_parameters,
            )
            changes += formatted
            notes.extend(format_notes)
        else:
            changes += normalize_docx(document)

    clear_docx_metadata(document)
    def save_checked(temporary: Path) -> None:
        document.save(temporary)
        candidate = docx.Document(temporary)
        if original_content is not None and content_signature(candidate) != original_content:
            raise RuntimeError("排版内容完整性检查失败，未发布成果；请保留原件并检查对象或关系变化")
    atomic_writer(target, save_checked)
    reopened = docx.Document(target)
    return {
        "changes": changes,
        "notes": notes,
        "verification": f"structure_reopened;paragraphs={len(reopened.paragraphs)};sections={len(reopened.sections)}",
    }


def process_xlsx(
    source: Path,
    target: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> dict[str, Any]:
    openpyxl = import_openpyxl()
    keep_vba = source.suffix.lower() == ".xlsm"
    workbook = openpyxl.load_workbook(source, data_only=False, keep_vba=keep_vba)
    changes = 0
    notes: list[str] = []
    replacements = [
        (str(item.get("before", "")), str(item.get("after", "")))
        for item in operation_parameters.get("replace_text", {}).get("replacements", [])
        if item.get("before")
    ]
    for review in ai_reviews:
        locator = review.get("locator") or {}
        if locator.get("kind") != "xlsx_cell":
            continue
        sheet_name = str(locator.get("sheet", ""))
        coordinate = str(locator.get("cell", ""))
        if sheet_name not in workbook.sheetnames or not coordinate:
            notes.append(f"AI 建议定位失效：{sheet_name}!{coordinate}")
            continue
        cell = workbook[sheet_name][coordinate]
        before = str(review.get("before", ""))
        if str(cell.value) == before:
            cell.value = str(review.get("after", ""))
            changes += 1
        else:
            notes.append(f"AI 建议未应用：{sheet_name}!{coordinate} 的原文已变化")
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for column_index, cell in enumerate(row):
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                value = cell.value
                if "replace_text" in operation_kinds or "update_dates" in operation_kinds:
                    for before, after in replacements:
                        value, count = re.subn(re.escape(before), after, value)
                        changes += count
                if "redact_content" in operation_kinds:
                    context_parts: list[str] = []
                    if column_index > 0:
                        context_parts.append(str(row[column_index - 1].value or ""))
                    if cell.row > 1:
                        context_parts.append(str(sheet.cell(1, cell.column).value or ""))
                    value, count = redact_text(value, " ".join(context_parts))
                    changes += count
                if "normalize_whitespace" in operation_kinds:
                    value, count = normalize_whitespace_text(value)
                    changes += count
                if "remove_line_breaks" in operation_kinds:
                    value, count = re.subn(r"[\r\n]+", "", value)
                    changes += count
                cell.value = value
    if "set_xlsx_cell" in operation_kinds:
        params = operation_parameters.get("set_xlsx_cell", {})
        sheet_name = str(params.get("sheet", ""))
        coordinate = str(params.get("cell", "")).upper()
        if sheet_name not in workbook.sheetnames:
            raise RuntimeError(f"工作表不存在：{sheet_name}")
        if not re.fullmatch(r"[A-Z]{1,3}[1-9]\d*", coordinate):
            raise RuntimeError(f"单元格地址无效：{coordinate}")
        workbook[sheet_name][coordinate].value = params.get("value", "")
        changes += 1
    if "delete_blank_rows" in operation_kinds:
        for sheet in workbook.worksheets:
            for row_index in range(sheet.max_row, 0, -1):
                if all(sheet.cell(row_index, column).value is None for column in range(1, sheet.max_column + 1)):
                    sheet.delete_rows(row_index)
                    changes += 1
    if "replace_text" in operation_kinds and not replacements:
        notes.append("未配置具体替换对，公式和单元格保持原样")
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    try:
        workbook.properties.creator = ""
        workbook.properties.lastModifiedBy = ""
        workbook.properties.title = ""
        workbook.properties.subject = ""
        workbook.properties.keywords = ""
    except AttributeError:
        pass
    atomic_writer(target, lambda temporary: workbook.save(temporary))
    reopened = openpyxl.load_workbook(target, read_only=True, data_only=False)
    sheet_count = len(reopened.sheetnames)
    reopened.close()
    workbook.close()
    return {
        "changes": changes,
        "notes": notes,
        "verification": f"{sheet_count}_sheets_reopened",
    }


def csv_encoding(path: Path) -> str:
    sample = path.read_bytes()[:128_000]
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            sample.decode(encoding)
            return encoding
        except UnicodeDecodeError:
            continue
    return "utf-8"


def process_csv(
    source: Path,
    target: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> dict[str, Any]:
    encoding = csv_encoding(source)
    replacements = [
        (str(item.get("before", "")), str(item.get("after", "")))
        for item in operation_parameters.get("replace_text", {}).get("replacements", [])
        if item.get("before")
    ]
    changes = 0

    def write(temporary: Path) -> None:
        nonlocal changes
        with source.open("r", encoding=encoding, newline="") as input_stream:
            sample = input_stream.read(8192)
            input_stream.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel
            rows = csv.reader(input_stream, dialect)
            with temporary.open("w", encoding="utf-8-sig", newline="") as output_stream:
                writer = csv.writer(output_stream, dialect)
                headers: list[str] = []
                for row_index, row in enumerate(rows, start=1):
                    if row_index == 1:
                        headers = list(row)
                    output_row = []
                    for column_index, value in enumerate(row):
                        for review in ai_reviews:
                            locator = review.get("locator") or {}
                            if (
                                locator.get("kind") == "csv_cell"
                                and int(locator.get("row", -1)) == row_index
                                and str(locator.get("column", "")) == column_name(column_index)
                                and value == str(review.get("before", ""))
                            ):
                                value = str(review.get("after", ""))
                                changes += 1
                        if "replace_text" in operation_kinds or "update_dates" in operation_kinds:
                            for before, after in replacements:
                                value, count = re.subn(re.escape(before), after, value)
                                changes += count
                        if "redact_content" in operation_kinds:
                            context_parts = []
                            if column_index > 0:
                                context_parts.append(row[column_index - 1])
                            if column_index < len(headers):
                                context_parts.append(headers[column_index])
                            value, count = redact_text(value, " ".join(context_parts))
                            changes += count
                        if "normalize_whitespace" in operation_kinds:
                            value, count = normalize_whitespace_text(value)
                            changes += count
                        if "remove_line_breaks" in operation_kinds:
                            value, count = re.subn(r"[\r\n]+", "", value)
                            changes += count
                        output_row.append(value)
                    writer.writerow(output_row)

    atomic_writer(target, write)
    return {
        "changes": changes,
        "notes": [f"源编码：{encoding}；输出编码：UTF-8 with BOM"],
        "verification": "csv_reopened",
    }


def process_pdf(
    source: Path,
    target: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> dict[str, Any]:
    validate_pdf_operations(operation_kinds)
    if "ocr_pdf" in operation_kinds or operation_parameters.get("review_content", {}).get("ocr"):
        return process_pdf_actions(source, target, operation_kinds | {"ocr_pdf"}, operation_parameters)
    pypdf = import_pypdf()
    reader = pypdf.PdfReader(source)
    writer = pypdf.PdfWriter()
    writer.clone_document_from_reader(reader)
    try:
        writer.metadata = None
    except AttributeError:
        writer.add_metadata({"/Producer": "材料助手"})
    notes: list[str] = []
    changes = 0
    if "convert_to_pdf" in operation_kinds:
        notes.append("源文件已为 PDF，无需格式转换")
    if "redact_content" in operation_kinds:
        notes.append(
            "PDF 正文默认不直接重写；本次已清理文档元数据并保留页面，正文疑似敏感项写入报告供人工确认"
        )
    if ai_reviews:
        notes.append(f"{len(ai_reviews)} 条 PDF AI 建议仅写入审校报告，未直接重写正文")
    atomic_writer(target, lambda temporary: _write_pdf(writer, temporary))
    reopened = pypdf.PdfReader(target)
    verification = f"{len(reopened.pages)}_pages_reopened"

    return {
        "changes": changes,
        "notes": notes,
        "verification": verification,
    }


def _write_pdf(writer: Any, target: Path) -> None:
    with target.open("wb") as stream:
        writer.write(stream)


def pdf_page_texts(pages: list[Any], *, ocr: bool, languages: list[str] | None = None) -> tuple[list[str], dict[int, float], list[str]]:
    """Read each text layer once; OCR only pages without extractable text.

    OCR page numbers refer to the private subset PDF, never the source PDF.
    Keep this mapping explicit so mixed documents cannot silently lose pages.
    """
    texts = [(page.extract_text() or "").strip() for page in pages]
    missing = [index for index, text in enumerate(texts) if not text]
    confidences: dict[int, float] = {}
    notes: list[str] = []
    if not ocr or not missing:
        if ocr:
            notes.append("全部页面已有可提取文字，直接导出，未调用 OCR")
        return texts, confidences, notes
    ocr_binary = os.environ.get("MATERIAL_WORKSHOP_VISION_OCR") or os.environ.get("MATERIAL_WORKSHOP_RAPID_OCR")
    if not ocr_binary:
        raise RuntimeError(f"第 {','.join(str(index + 1) for index in missing)} 页无可提取文字；本机 OCR 组件不可用，未交付不完整文本")
    if languages is None:
        languages = ["zh-Hans", "en-US"]
    if not isinstance(languages, list) or not languages or any(not isinstance(value, str) or not value.strip() or "," in value for value in languages):
        raise ValueError("OCR languages 必须是非空语言标识数组")
    pypdf = import_pypdf()
    with tempfile.TemporaryDirectory(prefix="zhai-ocr-pages-") as directory:
        subset = Path(directory) / "pages.pdf"
        writer = pypdf.PdfWriter()
        for index in missing:
            writer.add_page(pages[index])
        _write_pdf(writer, subset)
        completed = subprocess.run(
            [ocr_binary, "--input", str(subset), "--languages", ",".join(languages)],
            capture_output=True, text=True, timeout=300, check=False,
        )
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or "本地 OCR 执行失败")
    try:
        payload = json.loads(completed.stdout)
        recognized = payload["pages"]
        if not isinstance(recognized, list) or len(recognized) != len(missing):
            raise ValueError("页数不匹配")
        by_page: dict[int, Any] = {}
        for page in recognized:
            number = page["page"]
            confidence = page["averageConfidence"]
            lines = page["lines"]
            if type(number) is not int or number < 1 or number > len(missing) or number in by_page:
                raise ValueError("页号重复或越界")
            if type(confidence) not in (float, int) or not math.isfinite(confidence) or not 0 <= confidence <= 1:
                raise ValueError("置信度无效")
            if not isinstance(lines, list) or any(not isinstance(line, dict) or not isinstance(line.get("text"), str) for line in lines):
                raise ValueError("文字行无效")
            by_page[number] = page
        for number, index in enumerate(missing, start=1):
            page = by_page[number]
            texts[index] = "\n".join(line["text"] for line in page["lines"]).strip()
            confidences[index] = float(page["averageConfidence"])
            if not texts[index]:
                notes.append(f"第 {index + 1} 页未识别到文字，请核对是否为空白页或识别失败")
            elif confidences[index] < 0.75:
                notes.append(f"第 {index + 1} 页 OCR 平均置信度低于 0.75，必须人工确认")
    except (ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(f"OCR 返回结果不完整或无效，未交付文本：{exc}") from exc
    if not any(texts):
        raise RuntimeError("OCR 未识别到任何文字，请检查扫描清晰度；未交付空文本")
    notes.append(f"保留 {len(pages) - len(missing)} 页原有文字，仅对 {len(missing)} 页调用 OCR；识别结果须复核")
    return texts, confidences, notes


def extract_text_preview(source: Path) -> str:
    extension = source.suffix.lower()
    if extension == ".doc":
        return extract_legacy_doc_direct(source)[:20_000]
    if extension == ".docx":
        document = import_docx().Document(source)
        return "\n".join(paragraph.text for paragraph in iter_docx_paragraphs(document))[:20_000]
    if extension == ".xlsx":
        workbook = import_openpyxl().load_workbook(source, read_only=True, data_only=False)
        values: list[str] = []
        for sheet in workbook.worksheets:
            values.append(f"[{sheet.title}]")
            for row in sheet.iter_rows(values_only=True):
                values.append("\t".join("" if value is None else str(value) for value in row))
                if sum(len(item) for item in values) > 20_000:
                    break
        workbook.close()
        return "\n".join(values)[:20_000]
    if extension == ".pdf":
        reader = import_pypdf().PdfReader(source)
        return "\n".join((page.extract_text() or "") for page in reader.pages)[:20_000]
    if extension == ".csv":
        return source.read_text(encoding=csv_encoding(source), errors="replace")[:20_000]
    if extension in {".md", ".txt"}:
        return source.read_text(encoding="utf-8-sig", errors="replace")[:20_000]
    return ""


def extract_legacy_doc_direct(source: Path) -> str:
    """Read binary DOC text without converting it to DOCX."""
    if sys.platform == "darwin":
        command = ["/usr/bin/textutil", "-convert", "txt", "-stdout", str(source)]
        environment = None
    else:
        antiword = (
            os.environ.get("MATERIAL_WORKSHOP_ANTIWORD")
            or shutil.which("zhai-antiword")
            or shutil.which("antiword")
        )
        if not antiword:
            raise RuntimeError("应用内置的旧版 DOC 读取器不可用，请重新安装应用")
        command = [antiword, str(source)]
        environment = os.environ.copy()
        configured_home = os.environ.get("MATERIAL_WORKSHOP_ANTIWORD_HOME")
        bundled_home = Path("/usr/lib/zhai/antiword")
        if configured_home:
            environment["ANTIWORDHOME"] = configured_home
        elif bundled_home.is_dir():
            environment["ANTIWORDHOME"] = str(bundled_home)
        environment["LC_ALL"] = "C.UTF-8"
    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        timeout=60,
        env=environment,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or b"DOC reader failed").decode(
            "utf-8", errors="replace"
        )
        raise RuntimeError(f"无法直接读取旧版 DOC：{detail.splitlines()[0]}")
    return completed.stdout.decode("utf-8", errors="replace").replace("\x00", "").strip()


def column_name(index: int) -> str:
    value = index + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def extract_structured_file(item: dict[str, Any]) -> list[dict[str, Any]]:
    source = Path(item["path"])
    extension = source.suffix.lower()
    snippets: list[dict[str, Any]] = []
    if extension == ".doc":
        text = extract_legacy_doc_direct(source)
        for index, paragraph in enumerate(part for part in text.splitlines() if part.strip()):
            snippets.append(
                {
                    "fileId": item["id"],
                    "locator": {"kind": "legacy_doc_paragraph", "index": index},
                    "text": paragraph,
                }
            )
    elif extension == ".docx":
        document = import_docx().Document(source)
        for index, paragraph in enumerate(iter_docx_paragraphs(document)):
            if paragraph.text.strip():
                snippets.append(
                    {
                        "fileId": item["id"],
                        "locator": {"kind": "docx_paragraph", "index": index},
                        "text": paragraph.text,
                    }
                )
    elif extension == ".xlsx":
        workbook = import_openpyxl().load_workbook(source, read_only=True, data_only=False)
        try:
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is not None and str(cell.value).strip():
                            snippets.append(
                                {
                                    "fileId": item["id"],
                                    "locator": {
                                        "kind": "xlsx_cell",
                                        "sheet": sheet.title,
                                        "cell": cell.coordinate,
                                    },
                                    "text": str(cell.value),
                                }
                            )
        finally:
            workbook.close()
    elif extension == ".pdf":
        reader = import_pypdf().PdfReader(source)
        texts, _, _ = pdf_page_texts(list(reader.pages), ocr=bool(os.environ.get("MATERIAL_WORKSHOP_VISION_OCR") or os.environ.get("MATERIAL_WORKSHOP_RAPID_OCR")))
        for page_index, text in enumerate(texts, start=1):
            if text.strip():
                snippets.append(
                    {
                        "fileId": item["id"],
                        "locator": {"kind": "pdf_page", "page": page_index},
                        "text": text,
                    }
                )
    elif extension in {".md", ".txt"}:
        for line_index, line in enumerate(source.read_text(encoding="utf-8-sig", errors="replace").splitlines(), start=1):
            if line.strip():
                snippets.append(
                    {
                        "fileId": item["id"],
                        "locator": {"kind": "text_line", "line": line_index},
                        "text": line,
                    }
                )
    elif extension == ".csv":
        encoding = csv_encoding(source)
        with source.open("r", encoding=encoding, newline="") as stream:
            sample = stream.read(8192)
            stream.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel
            for row_index, row in enumerate(csv.reader(stream, dialect), start=1):
                for column_index, value in enumerate(row):
                    if value.strip():
                        snippets.append(
                            {
                                "fileId": item["id"],
                                "locator": {
                                    "kind": "csv_cell",
                                    "row": row_index,
                                    "column": column_name(column_index),
                                },
                                "text": value,
                            }
                        )
    return snippets


def run_extraction(request_path: Path) -> dict[str, Any]:
    if not request_path.is_absolute():
        raise ValueError("提取请求必须使用绝对路径")
    request = json.loads(request_path.read_text(encoding="utf-8"))
    if request.get("schemaVersion") != 1:
        raise ValueError("不支持的提取协议版本")
    snippets: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    for item in request.get("files", []):
        try:
            snippets.extend(extract_structured_file(item))
        except Exception as exc:  # noqa: BLE001 - extraction remains per-file isolated
            errors.append(
                {
                    "fileId": str(item.get("id", "")),
                    "message": str(exc) or exc.__class__.__name__,
                }
            )
    return {"snippets": snippets, "errors": errors}


def inspect_document(source: Path, target: Path) -> dict[str, Any]:
    extension = source.suffix.lower().lstrip(".")
    payload: dict[str, Any] = {
        "file": source.name,
        "format": extension,
        "bytes": source.stat().st_size,
        "sha256": checksum(source),
    }
    text = extract_text_preview(source)
    payload["textCharactersSampled"] = len(text)
    payload["cjkCharactersSampled"] = cjk_character_count(text)
    if extension == "docx":
        document = import_docx().Document(source)
        payload.update(paragraphs=len(list(iter_docx_paragraphs(document))), tables=len(document.tables), sections=len(document.sections))
    elif extension == "xlsx":
        workbook = import_openpyxl().load_workbook(source, read_only=True, data_only=False)
        payload["sheets"] = [
            {"name": sheet.title, "rows": sheet.max_row, "columns": sheet.max_column}
            for sheet in workbook.worksheets
        ]
        workbook.close()
    elif extension == "pdf":
        reader = import_pypdf().PdfReader(source)
        payload["pages"] = len(reader.pages)
        payload["pagesWithText"] = sum(bool((page.extract_text() or "").strip()) for page in reader.pages)
        payload["encrypted"] = bool(reader.is_encrypted)
    elif extension == "csv":
        encoding = csv_encoding(source)
        with source.open("r", encoding=encoding, newline="") as stream:
            row_count = sum(1 for _ in csv.reader(stream))
        payload.update(encoding=encoding, rows=row_count)
    elif extension == "doc":
        payload["reader"] = "textutil" if sys.platform == "darwin" else "antiword"
        payload["convertedBeforeRead"] = False
    atomic_writer(
        target,
        lambda temporary: temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        ),
    )
    return {"changes": 0, "notes": ["检查结果已生成"], "verification": "inspection_json_written"}


def csv_to_xlsx(source: Path, target: Path) -> dict[str, Any]:
    openpyxl = import_openpyxl()
    encoding = csv_encoding(source)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "数据"
    rows = 0
    with source.open("r", encoding=encoding, newline="") as stream:
        sample = stream.read(8192)
        stream.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        for row in csv.reader(stream, dialect):
            sheet.append(row)
            rows += 1
    atomic_writer(target, lambda temporary: workbook.save(temporary))
    workbook.close()
    reopened = openpyxl.load_workbook(target, read_only=True)
    reopened.close()
    return {"changes": rows, "notes": [f"已读取 {rows} 行；源编码 {encoding}"], "verification": "xlsx_reopened"}


def xlsx_to_csv(source: Path, target: Path, parameters: dict[str, Any]) -> dict[str, Any]:
    workbook = import_openpyxl().load_workbook(source, read_only=True, data_only=False)
    requested = str(parameters.get("sheet", "")).strip()
    if requested and requested not in workbook.sheetnames:
        workbook.close()
        raise RuntimeError(f"工作表不存在：{requested}")
    sheet = workbook[requested] if requested else workbook.worksheets[0]
    row_count = 0

    def write(temporary: Path) -> None:
        nonlocal row_count
        with temporary.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream)
            for row in sheet.iter_rows(values_only=True):
                writer.writerow(["" if value is None else value for value in row])
                row_count += 1

    atomic_writer(target, write)
    workbook.close()
    return {"changes": row_count, "notes": [f"已导出工作表：{sheet.title}"], "verification": "csv_written_utf8_bom"}


def merge_xlsx_files(
    files: list[dict[str, Any]],
    output_root: Path,
    parameters: dict[str, Any],
) -> tuple[str, int, int]:
    """Merge rows from XLSX workbooks, aligning columns by header text."""
    openpyxl = import_openpyxl()
    sources = [
        Path(item["path"])
        for item in files
        if Path(item["path"]).suffix.lower() == ".xlsx"
    ]
    if len(sources) < 2:
        raise RuntimeError("合并表格至少需要 2 个 XLSX 文件")

    requested_sheet = str(parameters.get("sheet", "")).strip()
    try:
        header_row = int(parameters.get("headerRow", 1))
    except (TypeError, ValueError) as exc:
        raise RuntimeError("表头行必须是正整数") from exc
    if header_row < 1:
        raise RuntimeError("表头行必须是正整数")
    include_source = bool(parameters.get("includeSource", True))

    unified_headers: list[str] = []
    all_rows: list[tuple[str, dict[str, Any]]] = []
    source_hashes = {source: checksum(source) for source in sources}

    for source in sources:
        workbook = openpyxl.load_workbook(source, read_only=False, data_only=False)
        try:
            if requested_sheet:
                if requested_sheet not in workbook.sheetnames:
                    raise RuntimeError(f"{source.name} 中不存在工作表“{requested_sheet}”")
                sheet = workbook[requested_sheet]
            else:
                sheet = workbook.worksheets[0]
            if sheet.max_row < header_row:
                raise RuntimeError(f"{source.name} 不存在第 {header_row} 行表头")
            if any(area.min_row <= header_row <= area.max_row for area in sheet.merged_cells.ranges):
                raise RuntimeError(f"{source.name} 的表头包含合并单元格，请指定独立字段名所在行，或设计适用于多层表头的能力包")

            raw_headers = [
                sheet.cell(header_row, column).value
                for column in range(1, sheet.max_column + 1)
            ]
            headers: list[str] = []
            seen: dict[str, int] = {}
            for column, value in enumerate(raw_headers, start=1):
                base = str(value).strip() if value is not None else ""
                if not base:
                    if any(sheet.cell(row, column).value is not None for row in range(header_row + 1, sheet.max_row + 1)):
                        raise RuntimeError(f"{source.name} 第 {column} 列有数据但缺少表头，无法可靠对齐字段")
                    headers.append("")
                    continue
                seen[base] = seen.get(base, 0) + 1
                if seen[base] > 1:
                    raise RuntimeError(f"{source.name} 存在重复表头“{base}”，请先明确字段名，避免汇总错列")
                header = base
                headers.append(header)
                if header not in unified_headers:
                    unified_headers.append(header)

            for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                if not any(value is not None and str(value).strip() for value in values):
                    continue
                if any(isinstance(value, str) and value.startswith("=") for value in values):
                    raise RuntimeError(f"{source.name} 含公式，跨文件汇总可能改变引用；请先另存为数值副本，或设计保留公式的专用能力")
                all_rows.append(
                    (
                        source.name,
                        {
                            header: values[index] if index < len(values) else None
                            for index, header in enumerate(headers)
                            if header
                        },
                    )
                )
        finally:
            workbook.close()

    if not all_rows:
        raise RuntimeError("所选表格中没有可合并的数据行")

    source_header = "来源文件"
    while source_header in unified_headers:
        source_header += "_"
    output_headers = ([source_header] if include_source else []) + unified_headers
    target = unique_output_path(output_root, Path("Excel数据汇总.xlsx"))
    merged_workbook = openpyxl.Workbook()
    merged_sheet = merged_workbook.active
    merged_sheet.title = "汇总"
    merged_sheet.append(output_headers)
    for source_name, row in all_rows:
        values = [row.get(header) for header in unified_headers]
        merged_sheet.append(([source_name] if include_source else []) + values)

    from openpyxl.styles import Alignment, Font, PatternFill

    header_fill = PatternFill("solid", fgColor="DCEEFF")
    for cell in merged_sheet[1]:
        cell.font = Font(bold=True, color="1F4E78")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    merged_sheet.freeze_panes = "A2"
    merged_sheet.auto_filter.ref = merged_sheet.dimensions
    for column_cells in merged_sheet.columns:
        sample = list(column_cells[: min(len(column_cells), 200)])
        width = max(
            (len(str(cell.value)) for cell in sample if cell.value is not None),
            default=8,
        )
        merged_sheet.column_dimensions[column_cells[0].column_letter].width = min(
            max(width + 2, 10), 36
        )
    merged_workbook.properties.creator = ""
    merged_workbook.properties.lastModifiedBy = ""
    atomic_writer(target, lambda temporary: merged_workbook.save(temporary))
    merged_workbook.close()

    for source, before_hash in source_hashes.items():
        if checksum(source) != before_hash:
            raise RuntimeError(f"安全校验失败：合并期间原文件发生变化：{source.name}")
    reopened = openpyxl.load_workbook(target, read_only=True, data_only=False)
    try:
        sheet = reopened["汇总"]
        if sheet.max_row != len(all_rows) + 1 or sheet.max_column != len(output_headers):
            raise RuntimeError("合并结果重新打开后行列数不一致")
    finally:
        reopened.close()
    return str(target), len(all_rows), len(sources)


def parse_page_selection(value: str, total_pages: int) -> list[int]:
    selected: list[int] = []
    for token in (part.strip() for part in value.split(",")):
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            if not start_text.isdigit() or not end_text.isdigit():
                raise RuntimeError(f"页码范围无效：{token}")
            start, end = int(start_text), int(end_text)
            if start > end:
                raise RuntimeError(f"页码范围起点大于终点：{token}")
            selected.extend(range(start, end + 1))
        elif token.isdigit():
            selected.append(int(token))
        else:
            raise RuntimeError(f"页码无效：{token}")
    unique = list(dict.fromkeys(selected))
    if not unique:
        raise RuntimeError("没有指定要提取的页码")
    if min(unique) < 1 or max(unique) > total_pages:
        raise RuntimeError(f"页码超出范围，当前 PDF 共 {total_pages} 页")
    return unique


def process_pdf_actions(
    source: Path,
    target: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    validate_pdf_operations(operation_kinds)
    if "redact_content" in operation_kinds:
        raise ValueError("PDF 正文脱敏不能与页面处理或文字提取同时执行；请使用单独的元数据清理或人工核对正文")
    if operation_kinds & {"ocr_pdf", "extract_pdf_text"} and operation_kinds & {"watermark_pdf", "clean_metadata", "redact_content"}:
        raise ValueError("文本导出/OCR 不能与水印、脱敏或元数据修改同时交付，请分别运行 PDF 修改与文字提取")
    pypdf = import_pypdf()
    reader = pypdf.PdfReader(source)
    writer = pypdf.PdfWriter()
    selected = list(range(1, len(reader.pages) + 1))
    notes: list[str] = []
    if "split_pdf" in operation_kinds:
        selected = parse_page_selection(str(operation_parameters.get("split_pdf", {}).get("pages", "")), len(reader.pages))
        for page_number in selected:
            writer.add_page(reader.pages[page_number - 1])
        notes.append(f"已提取页码：{','.join(map(str, selected))}")
    else:
        writer.clone_document_from_reader(reader)
    if "rotate_pdf" in operation_kinds:
        degrees = int(operation_parameters.get("rotate_pdf", {}).get("degrees", 90))
        if degrees not in {90, 180, 270}:
            raise RuntimeError("旋转角度必须是 90、180 或 270")
        for page in writer.pages:
            page.rotate(degrees)
        notes.append(f"所选页面已顺时针旋转 {degrees}°")
    if "watermark_pdf" in operation_kinds:
        watermark_text = str(operation_parameters.get("watermark_pdf", {}).get("text", "")).strip()
        if not watermark_text:
            raise RuntimeError("水印文字不能为空")
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen import canvas

        try:
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        except Exception:  # noqa: BLE001 - duplicate registration is harmless
            pass
        for page in writer.pages:
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            buffer = io.BytesIO()
            overlay = canvas.Canvas(buffer, pagesize=(width, height))
            overlay.saveState()
            overlay.setFillAlpha(0.16)
            overlay.setFillColorRGB(0.25, 0.25, 0.25)
            overlay.setFont("STSong-Light", 34)
            overlay.translate(width / 2, height / 2)
            overlay.rotate(35)
            overlay.drawCentredString(0, 0, watermark_text)
            overlay.restoreState()
            overlay.save()
            buffer.seek(0)
            page.merge_page(pypdf.PdfReader(buffer).pages[0])
        notes.append(f"已添加水印：{watermark_text}")
    if "clean_metadata" in operation_kinds:
        writer.metadata = None
        notes.append("已清理 PDF 元数据")
    order = [name for kind, name in [("split_pdf", "提取页码"), ("rotate_pdf", "旋转"), ("watermark_pdf", "水印"), ("clean_metadata", "清理元数据"), ("ocr_pdf", "按页提取及 OCR"), ("extract_pdf_text", "导出文本")] if kind in operation_kinds]
    if len(order) > 1:
        notes.append("处理顺序：" + " → ".join(order))
    if operation_kinds & {"extract_pdf_text", "ocr_pdf"}:
        use_ocr = "ocr_pdf" in operation_kinds
        parameters = operation_parameters.get("ocr_pdf", operation_parameters.get("review_content", {}))
        texts, confidences, text_notes = pdf_page_texts(list(writer.pages), ocr=use_ocr, languages=parameters.get("languages"))
        notes.extend(text_notes)
        empty_pages = [selected[index] for index, text in enumerate(texts) if not text]
        if empty_pages:
            notes.append(f"原文第 {','.join(map(str, empty_pages))} 页无可提取文字；空白或扫描内容须人工核对" + ("" if use_ocr else "，本次未执行 OCR"))
        if not any(texts):
            raise RuntimeError("PDF 没有可提取文字，未交付空文本；扫描件需要可用的本地 OCR")
        blocks = []
        for index, text in enumerate(texts):
            method = f" · OCR 平均置信度 {confidences[index]:.3f}" if index in confidences else " · 原有文本层"
            blocks.append(f"--- 第 {selected[index]} 页{method} ---\n{text or '[未提取到文字，请核对原页]'}")
        content = "\n\n".join(blocks) + "\n"
        def write_text(temporary: Path) -> None:
            temporary.write_text(content, encoding="utf-8")
            if temporary.read_text(encoding="utf-8") != content:
                raise RuntimeError("PDF 文本重新打开校验失败，拒绝交付")
        atomic_writer(target, write_text)
        notes.append("已按原页序导出 UTF-8 文本；未生成可搜索 PDF")
        return {"changes": len(texts), "notes": notes, "verification": "text_utf8_reopened"}
    def write_pdf(temporary: Path) -> None:
        _write_pdf(writer, temporary)
        if len(pypdf.PdfReader(temporary).pages) != len(selected):
            raise RuntimeError("PDF 页数重新打开校验失败，拒绝交付")
    atomic_writer(target, write_pdf)
    return {"changes": len(selected), "notes": notes, "verification": f"{len(selected)}_pages_reopened"}


@dataclass
class FileResult:
    file_id: str
    source: str
    output: str | None
    status: str
    changes: int = 0
    notes: list[str] = field(default_factory=list)
    verification: str | None = None
    source_sha256_before: str | None = None
    source_sha256_after: str | None = None
    error: str | None = None


def process_file(
    item: dict[str, Any],
    output_root: Path,
    operation_kinds: set[str],
    operation_parameters: dict[str, dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> FileResult:
    source = Path(item["path"])
    if not source.is_absolute() or not source.is_file():
        raise ValueError("源文件不存在或不是普通文件")
    extension = source.suffix.lower().lstrip(".")
    if extension == "pdf":
        validate_pdf_operations(operation_kinds)
    legacy_conversion = (extension == "doc" and bool(operation_kinds)) or (
        extension == "xls" and "xls_to_xlsx" in operation_kinds
    )
    if extension in LEGACY_EXTENSIONS and not legacy_conversion:
        raise RuntimeError("旧版 DOC/XLS 必须先通过 LibreOffice 或已验证的办公软件桥接转换")
    if extension not in SUPPORTED_EXTENSIONS and extension not in LEGACY_EXTENSIONS:
        raise RuntimeError(f"暂不支持 {extension.upper()} 格式")
    before_hash = checksum(source)
    relative = safe_relative_path(item.get("relativePath", source.name), source.name)
    converting_to_pdf = "convert_to_pdf" in operation_kinds and extension in {"doc", "docx", "xlsx"}
    if "apply_redhead" in operation_kinds:
        target_relative = relative.with_name(relative.stem + '_套红.docx')
    elif "inspect_document" in operation_kinds:
        target_relative = relative.with_suffix(relative.suffix + ".inspection.json")
    elif converting_to_pdf:
        target_relative = relative.with_suffix(".pdf")
    elif "generate_gongwen" in operation_kinds and extension in {"md", "txt"}:
        target_relative = relative.with_suffix(".docx")
    elif "doc_to_docx" in operation_kinds and extension == "doc":
        target_relative = relative.with_suffix(".docx")
    elif extension == "doc":
        target_relative = relative.with_suffix(".docx")
    elif "xls_to_xlsx" in operation_kinds and extension == "xls":
        target_relative = relative.with_suffix(".xlsx")
    elif "csv_to_xlsx" in operation_kinds and extension == "csv":
        target_relative = relative.with_suffix(".xlsx")
    elif "xlsx_to_csv" in operation_kinds and extension == "xlsx":
        target_relative = relative.with_suffix(".csv")
    elif extension == "pdf" and ("ocr_pdf" in operation_kinds or operation_parameters.get("review_content", {}).get("ocr")):
        target_relative = relative.with_suffix(".ocr.txt")
    elif "extract_pdf_text" in operation_kinds and extension == "pdf":
        target_relative = relative.with_suffix(".txt")
    else:
        target_relative = relative
    target = unique_output_path(output_root, target_relative)

    if "apply_redhead" in operation_kinds:
        import redhead
        if extension not in {'doc', 'docx'}:
            raise ValueError('套红仅支持 DOC/DOCX 文件')
        detail = redhead.compose(source, target, operation_parameters.get('apply_redhead', {}), sys.modules[__name__])
    elif "inspect_document" in operation_kinds:
        detail = inspect_document(source, target)
    elif "generate_gongwen" in operation_kinds and extension in {"md", "txt"}:
        detail = generate_gongwen_docx(
            source,
            target,
            operation_parameters.get("generate_gongwen", {}),
        )
    elif "doc_to_docx" in operation_kinds and extension == "doc":
        detail = convert_office(
            source,
            target,
            "doc-to-docx",
            operation_parameters.get("doc_to_docx", {}),
        )
    elif "xls_to_xlsx" in operation_kinds and extension == "xls":
        detail = convert_office(
            source,
            target,
            "xls-to-xlsx",
            operation_parameters.get("xls_to_xlsx", {}),
        )
    elif "xlsx_recalculate" in operation_kinds and extension == "xlsx":
        detail = convert_office(
            source,
            target,
            "xlsx-recalculate",
            operation_parameters.get("xlsx_recalculate", {}),
        )
    elif "csv_to_xlsx" in operation_kinds and extension == "csv":
        detail = csv_to_xlsx(source, target)
    elif "xlsx_to_csv" in operation_kinds and extension == "xlsx":
        detail = xlsx_to_csv(source, target, operation_parameters.get("xlsx_to_csv", {}))
    elif converting_to_pdf:
        content_operations = operation_kinds - {"convert_to_pdf", "merge_or_split"}
        if content_operations:
            with tempfile.TemporaryDirectory(
                prefix="material-assistant-preconvert-",
                dir=output_root,
            ) as temporary_directory:
                intermediate = Path(temporary_directory) / source.with_suffix(".docx").name
                if extension == "doc":
                    converted_source = Path(temporary_directory) / f"converted-{source.stem}.docx"
                    convert_office(
                        source,
                        converted_source,
                        "doc-to-docx",
                        operation_parameters.get("convert_to_pdf", {}),
                    )
                    prepared = process_docx(
                        converted_source,
                        intermediate,
                        content_operations,
                        operation_parameters,
                        ai_reviews,
                    )
                elif extension == "docx":
                    prepared = process_docx(
                        source,
                        intermediate,
                        content_operations,
                        operation_parameters,
                        ai_reviews,
                    )
                else:
                    prepared = process_xlsx(
                        source,
                        intermediate,
                        content_operations,
                        operation_parameters,
                        ai_reviews,
                    )
                action = "docx-to-pdf" if extension in {"doc", "docx"} else "xlsx-to-pdf"
                converted = convert_office(
                    intermediate,
                    target,
                    action,
                    operation_parameters.get("convert_to_pdf", {}),
                )
                detail = {
                    "changes": int(prepared.get("changes", 0))
                    + int(converted.get("changes", 0)),
                    "notes": list(prepared.get("notes", []))
                    + list(converted.get("notes", [])),
                    "verification": (
                        f"{prepared.get('verification', 'prepared')};"
                        f"{converted.get('verification', 'converted')}"
                    ),
                }
        else:
            action = "doc-to-pdf" if extension == "doc" else "docx-to-pdf" if extension == "docx" else "xlsx-to-pdf"
            detail = convert_office(
                source,
                target,
                action,
                operation_parameters.get("convert_to_pdf", {}),
            )
    elif extension == "doc":
        with tempfile.TemporaryDirectory(
            prefix="material-assistant-doc-edit-",
            dir=output_root,
        ) as temporary_directory:
            converted_source = Path(temporary_directory) / f"{source.stem}.docx"
            convert_office(source, converted_source, "doc-to-docx")
            detail = process_docx(
                converted_source,
                target,
                operation_kinds,
                operation_parameters,
                ai_reviews,
            )
            detail.setdefault("notes", []).insert(
                0,
                "旧版 DOC 正文使用直接读取器识别；修改阶段转换为 DOCX 并生成新文件",
            )
    elif extension == "docx":
        detail = process_docx(source, target, operation_kinds, operation_parameters, ai_reviews)
    elif extension == "xlsx":
        detail = process_xlsx(source, target, operation_kinds, operation_parameters, ai_reviews)
    elif extension == "pdf":
        pdf_actions = {"split_pdf", "rotate_pdf", "watermark_pdf", "extract_pdf_text", "clean_metadata", "ocr_pdf"}
        if operation_kinds & pdf_actions:
            effective_kinds = operation_kinds | {"ocr_pdf"} if operation_parameters.get("review_content", {}).get("ocr") else operation_kinds
            detail = process_pdf_actions(source, target, effective_kinds, operation_parameters)
        else:
            detail = process_pdf(source, target, operation_kinds, operation_parameters, ai_reviews)
    elif extension == "csv":
        detail = process_csv(source, target, operation_kinds, operation_parameters, ai_reviews)
    else:
        copy_atomic(source, target)
        detail = {"changes": 0, "notes": [], "verification": "copied"}

    after_hash = checksum(source)
    if target.suffix.lower() in {'.docx', '.xlsx'} or converting_to_pdf:
        detail.setdefault('notes', []).extend(font_notices(target if target.suffix.lower() in {'.docx', '.xlsx'} else source))
        detail['notes'] = list(dict.fromkeys(detail['notes']))
    if before_hash != after_hash:
        raise RuntimeError("安全校验失败：处理期间原文件发生变化")
    return FileResult(
        file_id=item["id"],
        source=str(source),
        output=str(target),
        status="completed",
        changes=int(detail.get("changes", 0)),
        notes=list(detail.get("notes", [])),
        verification=detail.get("verification"),
        source_sha256_before=before_hash,
        source_sha256_after=after_hash,
    )


def merge_pdfs(
    files: list[dict[str, Any]],
    output_root: Path,
    operation_kinds: set[str],
) -> str | None:
    if "merge_or_split" not in operation_kinds:
        return None
    # Snapshot paths contain random UUIDs. Order by the authorized display name,
    # including folder-relative paths, and preserve input order for equal names.
    ordered = sorted(files, key=lambda item: str(item.get("relativePath") or item.get("name") or Path(item["path"]).name).casefold())
    pdf_paths = [Path(item["path"]) for item in ordered]
    if any(path.suffix.lower() != ".pdf" for path in pdf_paths):
        raise ValueError("PDF 合并只能接收 PDF 文件，请先完成其他格式转换")
    if not pdf_paths:
        raise ValueError("没有可合并的 PDF")
    pypdf = import_pypdf()
    writer = pypdf.PdfWriter()
    hashes = {path: checksum(path) for path in pdf_paths}
    expected_pages = 0
    for path in pdf_paths:
        reader = pypdf.PdfReader(path)
        if not len(reader.pages):
            raise ValueError("不能合并零页 PDF")
        expected_pages += len(reader.pages)
        writer.append(reader)
    target = unique_output_path(output_root, Path("合并材料.pdf"))
    def write(temporary: Path) -> None:
        _write_pdf(writer, temporary)
        if len(pypdf.PdfReader(temporary).pages) != expected_pages:
            raise ValueError("合并 PDF 页数校验失败，拒绝交付")
        if any(checksum(path) != before for path, before in hashes.items()):
            raise ValueError("合并期间原文件发生变化，拒绝交付")
    atomic_writer(target, write)
    return str(target)


def consistency_findings(files: list[dict[str, Any]]) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    date_pattern = re.compile(r"(?:19|20)\d{2}[年./-]\d{1,2}[月./-]\d{1,2}日?")
    for item in files:
        path = Path(item["path"])
        if path.suffix.lower().lstrip(".") not in SUPPORTED_EXTENSIONS:
            continue
        try:
            text = extract_text_preview(path)
        except Exception as exc:  # noqa: BLE001 - per-file isolation is intentional
            findings.append({"file": str(path), "issue": f"无法提取文本：{exc}"})
            continue
        dates = sorted(set(date_pattern.findall(text)))
        if len(dates) > 1:
            findings.append({"file": str(path), "issue": "发现多个日期值", "values": dates[:20]})
    return findings


def write_report(
    report_root: Path,
    job_id: str,
    results: list[FileResult],
    merged: str | None,
    findings: list[dict[str, Any]],
    skips: list[dict[str, Any]],
    ai_reviews: list[dict[str, Any]],
) -> tuple[Path, Path]:
    payload = {
        "schemaVersion": 1,
        "jobId": job_id,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "total": len(results),
            "completed": sum(item.status == "completed" for item in results),
            "failed": sum(item.status == "failed" for item in results),
            "skipped": len(skips),
            "changes": sum(item.changes for item in results),
        },
        "mergedOutput": merged,
        "consistencyFindings": findings,
        "files": [item.__dict__ for item in results],
        "skips": skips,
        "acceptedAiReviews": ai_reviews,
    }
    json_path = report_root / "processing-report.json"
    atomic_writer(
        json_path,
        lambda temporary: temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        ),
    )
    rows = []
    for item in results:
        rows.append(
            "<tr>"
            f"<td>{html.escape(Path(item.source).name)}</td>"
            f"<td>{'成功' if item.status == 'completed' else '失败'}</td>"
            f"<td>{item.changes}</td>"
            f"<td>{html.escape(item.error or '；'.join(item.notes) or '已通过结构检查')}</td>"
            "</tr>"
        )
    for item in skips:
        rows.append(
            "<tr>"
            f"<td>{html.escape(Path(str(item.get('path', ''))).name)}</td>"
            "<td>跳过</td><td>0</td>"
            f"<td>{html.escape(str(item.get('reason', '旧版格式')))}</td>"
            "</tr>"
        )
    ai_rows = "".join(
        "<tr>"
        f"<td>{html.escape(str(item.get('fileId', '')))}</td>"
        f"<td>{html.escape(str(item.get('before', '')))}</td>"
        f"<td>{html.escape(str(item.get('after', '')))}</td>"
        f"<td>{html.escape(str(item.get('reason', '')))}</td>"
        "</tr>"
        for item in ai_reviews
    )
    ai_section = (
        "<h2>已接受的 AI 建议</h2>"
        "<table><thead><tr><th>文件 ID</th><th>修改前</th><th>修改后</th><th>理由</th></tr></thead>"
        f"<tbody>{ai_rows}</tbody></table>"
        if ai_reviews
        else ""
    )
    html_document = f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>材料助手处理报告</title>
<style>
body{{font:14px/1.6 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;color:#202321;margin:40px}}
h1{{font-size:24px}} .summary{{padding:16px;background:#f3f5f2;border-left:4px solid #1f5a42}}
table{{width:100%;border-collapse:collapse;margin-top:24px}}th,td{{padding:10px;border-bottom:1px solid #ddd;text-align:left}}
th{{background:#f7f7f5}} .muted{{color:#69706b}}
</style></head><body><h1>材料助手处理报告</h1>
<p class="muted">任务 {html.escape(job_id)} · {html.escape(payload["generatedAt"])}</p>
<div class="summary">共 {payload["summary"]["total"]} 个文件，成功 {payload["summary"]["completed"]} 个，
失败 {payload["summary"]["failed"]} 个，跳过 {payload["summary"]["skipped"]} 个，
记录修改 {payload["summary"]["changes"]} 处。</div>
<table><thead><tr><th>文件</th><th>状态</th><th>修改</th><th>检查结果</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table>{ai_section}</body></html>"""
    html_path = report_root / "处理报告.html"
    atomic_writer(
        html_path,
        lambda temporary: temporary.write_text(html_document, encoding="utf-8"),
    )
    return json_path, html_path


def run(request: dict[str, Any]) -> dict[str, Any]:
    job_id = str(request["jobId"])
    files = list(request["files"])
    output_root = Path(request["outputDirectory"])
    report_root = Path(request["reportDirectory"])
    output_root.mkdir(parents=True, exist_ok=True)
    report_root.mkdir(parents=True, exist_ok=True)
    operations = [item for item in request["operations"] if item.get("available", True)]
    operation_kinds = {str(item.get("kind", "")) for item in operations}
    pure_pdf_merge = operation_kinds == {"merge_or_split"} and not request.get("aiReviews")
    if "merge_or_split" in operation_kinds and (operation_kinds & {"ocr_pdf", "extract_pdf_text"} or any(item.get("kind") == "review_content" and (item.get("parameters") or {}).get("ocr") for item in operations)):
        raise ValueError("PDF 合并与文字提取/OCR 不能在同一步交付，请先合并 PDF 再提取文字")
    # Per-file PDF copies are intermediates, including when the legacy job
    # entry point opens outputDirectory directly rather than publishing outputs.
    processing_root = report_root / "merge-inputs" if "merge_or_split" in operation_kinds else output_root
    processing_root.mkdir(parents=True, exist_ok=True)
    operation_parameters: dict[str, dict[str, Any]] = {}
    for item in operations:
        kind = str(item.get("kind", ""))
        parameters = dict(item.get("parameters") or {})
        if kind in {"replace_text", "update_dates"}:
            existing = operation_parameters.setdefault("replace_text", {"replacements": []})
            existing.setdefault("replacements", []).extend(parameters.get("replacements") or [])
        else:
            operation_parameters[kind] = parameters
    ai_reviews = list(request.get("aiReviews") or [])
    if 'apply_redhead' in operation_kinds:
        params = operation_parameters['apply_redhead']
        params.pop('_templatePath', None)  # Only resolved authorized inputs can become templates.
        template_name = str(params.get('templateName') or '').strip()
        if template_name:
            def original_snapshot_name(name: str) -> str:
                # The desktop takes one or more UUID-prefixed read-only snapshots.
                # Strip only that exact prefix, never arbitrary suffix-match paths.
                return re.sub(r'^(?:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}-)+', '', name)
            matches = [f for f in files if any(n == template_name or original_snapshot_name(n) == template_name for n in (f.get('name', ''), Path(f['path']).name))]
            if len(matches) != 1:
                raise ValueError('红头模板名称必须精确匹配一个已上传的 Word 文件')
            template = matches[0]
            params['_templatePath'] = template['path']
            files = [f for f in files if f is not template]
        if not files:
            raise ValueError('请在红头模板以外添加需要套红的正文文件')
        redhead_input_hashes = {f['path']: checksum(Path(f['path'])) for f in request['files']}
    skips = list(request.get("skips") or [])
    results: list[FileResult] = []

    if "merge_xlsx" in operation_kinds:
        try:
            merged_path, row_count, source_count = merge_xlsx_files(
                files,
                output_root,
                operation_parameters.get("merge_xlsx", {}),
            )
            results.append(
                FileResult(
                    file_id="__merge_xlsx__",
                    source=f"{source_count} 个 Excel 文件",
                    output=merged_path,
                    status="completed",
                    changes=row_count,
                    notes=[f"已合并 {source_count} 个文件、{row_count} 行数据"],
                    verification="xlsx_reopened_and_dimensions_verified",
                )
            )
            emit(
                {
                    "type": "file_progress",
                    "fileId": "__batch__",
                    "progress": 100,
                    "message": f"已合并 {source_count} 个表格，共 {row_count} 行数据",
                }
            )
        except Exception as exc:  # noqa: BLE001
            message = str(exc) or exc.__class__.__name__
            results.append(
                FileResult(
                    file_id="__merge_xlsx__",
                    source="Excel 合并任务",
                    output=None,
                    status="failed",
                    error=message,
                )
            )
            emit({"type": "file_failed", "fileId": "__merge_xlsx__", "message": message})
    else:
        for index, item in enumerate(files):
            emit(
                {
                    "type": "file_progress",
                    "fileId": item["id"],
                    "progress": 5,
                    "message": f"正在处理 {item['name']}",
                }
            )
            try:
                file_reviews = [
                    review for review in ai_reviews if review.get("fileId") == item.get("id")
                ]
                if pure_pdf_merge:
                    source = Path(item["path"])
                    if not source.is_absolute() or not source.is_file() or source.suffix.lower() != ".pdf":
                        raise ValueError("PDF 合并需要有效的 PDF 原件")
                    result = FileResult(file_id=item["id"], source=str(source), output=None, status="pending", verification="merge_pending")
                else:
                    result = process_file(
                        item,
                        processing_root,
                        operation_kinds,
                        operation_parameters,
                        file_reviews,
                    )
                results.append(result)
                emit(
                    {
                        "type": "file_progress",
                        "fileId": item["id"],
                        "progress": 100,
                        "message": "已准备合并输入" if pure_pdf_merge else "处理并校验完成",
                    }
                )
            except Exception as exc:  # noqa: BLE001 - failures must not stop the batch
                message = str(exc) or exc.__class__.__name__
                results.append(
                    FileResult(
                        file_id=item["id"],
                        source=item["path"],
                        output=None,
                        status="failed",
                        error=message,
                    )
                )
                emit({"type": "file_failed", "fileId": item["id"], "message": message})
            total_progress = int(((index + 1) / max(len(files), 1)) * 100)
            emit(
                {
                    "type": "file_progress",
                    "fileId": "__batch__",
                    "progress": total_progress,
                    "message": f"已完成 {index + 1}/{len(files)}",
                }
            )

    merged = None
    if 'apply_redhead' in operation_kinds:
        for path, before in redhead_input_hashes.items():
            if checksum(Path(path)) != before:
                raise RuntimeError('套红原件哈希发生变化，拒绝交付')
    try:
        if "merge_or_split" in operation_kinds and any(item.status == "failed" for item in results):
            raise ValueError("部分输入处理失败，未交付不完整的合并 PDF")
        original_inputs = {item["id"]: item for item in files}
        merge_inputs = files if pure_pdf_merge else [
            {"path": item.output, "name": original_inputs.get(item.file_id, {}).get("name"), "relativePath": original_inputs.get(item.file_id, {}).get("relativePath")}
            for item in results
            if item.status == "completed"
            and item.output
            and Path(item.output).suffix.lower() == ".pdf"
        ]
        merged = merge_pdfs(merge_inputs, output_root, operation_kinds)
        if pure_pdf_merge and merged:
            for item in results:
                item.status = "completed"
                item.verification = "source_hash_unchanged_and_merged_pdf_reopened"
    except Exception as exc:  # noqa: BLE001
        if pure_pdf_merge:
            for item in results:
                if item.status == "pending":
                    item.status = "skipped"
                    item.notes.append("合并未完成，未交付任何最终 PDF")
        results.append(
            FileResult(
                file_id="__merge__",
                source="PDF 合并任务",
                output=None,
                status="failed",
                error=str(exc),
            )
        )

    findings = (
        consistency_findings(files)
        if "cross_file_consistency_check" in operation_kinds
        else []
    )
    _, html_report = write_report(
        report_root,
        job_id,
        results,
        merged,
        findings,
        skips,
        ai_reviews,
    )
    failures = [
        {"fileId": item.file_id, "path": item.source, "message": item.error or "未知错误"}
        for item in results
        if item.status == "failed"
    ]
    completed = sum(item.status == "completed" for item in results)
    status = "completed" if not failures else ("partial" if completed else "failed")
    outcome = {
        "jobId": job_id,
        "status": status,
        "outputDirectory": str(output_root.parent),
        "reportPath": str(html_report),
        "completedFiles": completed,
        "failedFiles": len(failures),
        "failures": failures,
        "skippedFiles": len(skips),
        "skips": skips,
        "outputs": ([{"fileId": "__merge__", "path": merged, "notes": ["已合并全部 PDF 并重新打开核对页数"]}] if merged else []) if "merge_or_split" in operation_kinds else [
            {"fileId": item.file_id, "path": item.output, "notes": item.notes}
            for item in results
            if item.status == "completed" and item.output
        ],
    }
    emit({"type": "completed", "reportPath": str(html_report)})
    emit({"type": "outcome", "outcome": outcome})
    return outcome


class SkillContext(dict[str, Any]):
    """稳定的技能运行接口。

    新技能可使用 context.inputs / context.output_dir / context.parameters；
    旧技能仍可使用 context["input_dir"] 这类字典访问。
    """

    def __init__(self, input_dir: Path, output_dir: Path, parameters: dict[str, Any], resources_dir: Path | None = None) -> None:
        sample_limit = parameters.pop("__material_workbench_sample_limit", None)
        super().__init__(
            input_dir=str(input_dir),
            output_dir=str(output_dir),
            parameters=parameters,
        )
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.parameters = parameters
        self.resources_dir = resources_dir
        self["resources_dir"] = str(resources_dir) if resources_dir else None
        inputs = sorted(path for path in input_dir.rglob("*") if path.is_file())
        if isinstance(sample_limit, int) and sample_limit > 0:
            inputs = inputs[:sample_limit]
        self.inputs = inputs
        # Skills that support both mappings and attribute contexts select the
        # mapping branch because SkillContext inherits dict. Keep that public
        # input list present too, with the same ordering and sample limit.
        self["inputs"] = [str(path) for path in inputs]
        self._progress: list[dict[str, Any]] = []

    def progress(
        self,
        message: str,
        *,
        current: int | None = None,
        total: int | None = None,
        file: str | Path | None = None,
    ) -> None:
        """记录可供工作台展示的处理进度，不向标准输出混入日志。"""
        event: dict[str, Any] = {"message": str(message)[0:500]}
        if current is not None:
            event["current"] = max(0, int(current))
        if total is not None:
            event["total"] = max(0, int(total))
        if file is not None:
            event["file"] = Path(file).name
        self._progress.append(event)

    def result(
        self,
        outputs: list[str | Path],
        summary: str = "技能已运行完成",
    ) -> dict[str, Any]:
        """构造统一的技能返回值。"""
        return {"summary": summary, "outputs": [str(path) for path in outputs]}


def run_agent_script(request_path: Path) -> dict[str, Any]:
    if sys.platform.startswith("linux"):
        import resource
        # Hard limits apply only to temporary/custom code, never mature engine
        # operations. Descendants inherit them and scripts cannot raise them.
        for kind, ceiling in [(resource.RLIMIT_AS, 2 * 1024**3), (resource.RLIMIT_FSIZE, 512 * 1024**2), (resource.RLIMIT_CPU, 300), (resource.RLIMIT_NOFILE, 256), (resource.RLIMIT_NPROC, 256), (resource.RLIMIT_CORE, 0)]:
            _, hard = resource.getrlimit(kind)
            limit = ceiling if hard == resource.RLIM_INFINITY else min(ceiling, hard)
            resource.setrlimit(kind, (limit, limit))
    if not request_path.is_absolute():
        raise ValueError("请求文件必须使用绝对路径")
    request = json.loads(request_path.read_text(encoding="utf-8"))
    if request.get("schemaVersion") != 1:
        raise ValueError("不支持的临时脚本请求协议版本")
    if request.get("neverOverwrite") is not True:
        raise ValueError("安全策略要求 neverOverwrite=true")
    for key in ("scriptPath", "inputDirectory", "outputDirectory"):
        if key not in request:
            raise ValueError(f"临时脚本请求缺少字段：{key}")
    script_path = Path(str(request["scriptPath"])).resolve()
    input_dir = Path(str(request["inputDirectory"])).resolve()
    output_dir = Path(str(request["outputDirectory"])).resolve()
    if not script_path.is_file() or not input_dir.is_dir():
        raise ValueError("临时脚本或输入材料目录不存在")
    output_dir.mkdir(parents=True, exist_ok=True)
    script = script_path.read_text(encoding="utf-8")
    if len(script) > 200_000:
        raise ValueError("临时脚本超过 200 KB")
    context = SkillContext(
        input_dir=input_dir,
        output_dir=output_dir,
        parameters=dict(request.get("parameters") or {}),
        resources_dir=Path(request["resourcesDirectory"]).resolve() if request.get("resourcesDirectory") else None,
    )
    if context.resources_dir:
        if not context.resources_dir.is_dir():
            raise ValueError("能力包资源目录不存在")
        helper_dir = context.resources_dir / "scripts"
        if helper_dir.is_dir():
            sys.path.insert(0, str(helper_dir))
    namespace: dict[str, Any] = {
        "__name__": "material_workbench_agent_script",
        "__file__": str(script_path),
        "CONTEXT": context,
    }
    stdout = io.StringIO()
    stderr = io.StringIO()
    with redirect_stdout(stdout), redirect_stderr(stderr):
        exec(compile(script, str(script_path), "exec"), namespace, namespace)  # noqa: S102 - OS sandbox is the security boundary
        if callable(namespace.get("run")):
            run_function = namespace["run"]
            positional = [
                parameter
                for parameter in inspect.signature(run_function).parameters.values()
                if parameter.kind
                in {parameter.POSITIONAL_ONLY, parameter.POSITIONAL_OR_KEYWORD}
            ]
            if len(positional) >= 2:
                result = run_function(context, context.parameters)
            else:
                result = run_function(context)
        else:
            result = namespace.get("RESULT")
    if result is None:
        result = {}
    if not isinstance(result, dict):
        raise ValueError("脚本必须通过 run(context) 返回字典")
    requested_outputs = result.get("outputs") or []
    if not isinstance(requested_outputs, list):
        raise ValueError("outputs 必须是路径数组")
    outputs: list[str] = []
    for value in requested_outputs:
        path = Path(str(value))
        if not path.is_absolute():
            path = output_dir / path
        path = path.resolve()
        if path.is_file() and path.is_relative_to(output_dir):
            outputs.append(str(path))
    # Explicit outputs=[] is a diagnostic request, never an invitation to
    # publish every JSON/log a script happened to write. Only legacy scripts
    # omitting the field use discovery, with canonical containment enforced.
    if "outputs" not in result:
        outputs = [str(path.resolve()) for path in output_dir.rglob("*") if path.is_file() and not path.is_symlink() and path.resolve().is_relative_to(output_dir)]
    return {
        "summary": str(result.get("summary") or "临时脚本已完成")[0:4000],
        "outputs": outputs,
        "progress": context._progress[-100:],
        "stdout": stdout.getvalue()[-8000:],
        "stderr": stderr.getvalue()[-8000:],
    }


def validate_artifact(path: Path) -> dict[str, Any]:
    """Reopen a generated artifact before it is exposed to the conversation."""
    if not path.is_absolute() or not path.is_file():
        raise ValueError("成果路径必须是存在的绝对文件路径")
    if path.stat().st_size == 0:
        raise ValueError("成果文件为空")
    extension = path.suffix.lower()
    detail: dict[str, Any] = {"format": extension.lstrip("."), "bytes": path.stat().st_size}
    if extension == ".docx":
        document = import_docx().Document(path)
        detail.update(
            paragraphs=len(document.paragraphs),
            tables=len(document.tables),
            sections=len(document.sections),
        )
    elif extension == ".xlsx":
        workbook = import_openpyxl().load_workbook(path, read_only=True, data_only=False)
        try:
            if not workbook.sheetnames:
                raise ValueError("XLSX 没有工作表")
            row_counts: dict[str, int] = {}
            non_empty_counts: dict[str, int] = {}
            column_non_empty: dict[str, dict[str, int]] = {}
            for sheet in workbook.worksheets[:20]:
                row_counts[sheet.title] = sheet.max_row
                non_empty_counts[sheet.title] = sum(
                    1 for row in sheet.iter_rows() if any(cell.value not in (None, "") for cell in row)
                )
                headers = next(
                    sheet.iter_rows(min_row=1, max_row=1, values_only=True),
                    (),
                )
                sheet_columns: dict[str, int] = {}
                for column_index, header in enumerate(headers, start=1):
                    header_text = str(header or "").strip()
                    if not header_text:
                        continue
                    sheet_columns[header_text] = sum(
                        1
                        for (value,) in sheet.iter_rows(
                            min_row=2,
                            min_col=column_index,
                            max_col=column_index,
                            values_only=True,
                        )
                        if value not in (None, "")
                    )
                column_non_empty[sheet.title] = sheet_columns
            detail.update(
                sheets=len(workbook.sheetnames),
                sheetNames=workbook.sheetnames[:20],
                rows=row_counts,
                nonEmptyRows=non_empty_counts,
                maxNonEmptyRows=max(non_empty_counts.values(), default=0),
                columnNonEmpty=column_non_empty,
            )
        finally:
            workbook.close()
    elif extension == ".pdf":
        reader = import_pypdf().PdfReader(path)
        if not reader.pages:
            raise ValueError("PDF 没有页面")
        detail.update(pages=len(reader.pages))
    elif extension == ".csv":
        with path.open("r", encoding=csv_encoding(path), errors="replace", newline="") as stream:
            rows = list(csv.reader(stream))
        if not rows:
            raise ValueError("CSV 没有数据行")
        detail.update(rows=len(rows), columns=max((len(row) for row in rows), default=0))
    elif extension == ".json":
        value = json.loads(path.read_text(encoding="utf-8"))
        detail.update(rootType=type(value).__name__)
    elif extension in {".txt", ".md", ".log", ".html"}:
        value = path.read_text(encoding="utf-8-sig", errors="strict")
        detail.update(characters=len(value))
    elif extension in {".png", ".jpg", ".jpeg"}:
        from PIL import Image

        with Image.open(path) as image:
            image.verify()
            detail.update(imageFormat=image.format)
    else:
        raise ValueError(f"尚未配置 {extension or '无扩展名'} 成果验收器")
    return {"valid": True, "detail": detail}


def main() -> int:
    parser = argparse.ArgumentParser(add_help=True)
    request_group = parser.add_mutually_exclusive_group(required=True)
    request_group.add_argument("--request", type=Path)
    request_group.add_argument("--extract-request", type=Path)
    request_group.add_argument("--run-agent-script", type=Path)
    request_group.add_argument("--validate-artifact", type=Path)
    request_group.add_argument("--word-request", type=Path)
    args = parser.parse_args()
    try:
        if args.extract_request:
            print(
                json.dumps(
                    run_extraction(args.extract_request),
                    ensure_ascii=False,
                    separators=(",", ":"),
                )
            )
        elif args.run_agent_script:
            print(json.dumps(run_agent_script(args.run_agent_script), ensure_ascii=False, separators=(",", ":")))
        elif args.word_request:
            from word_structure import run_request
            print(json.dumps(run_request(args.word_request), ensure_ascii=False, separators=(",", ":")))
        elif args.validate_artifact:
            print(json.dumps(validate_artifact(args.validate_artifact), ensure_ascii=False, separators=(",", ":")))
        else:
            request = read_request(args.request)
            run(request)
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"{exc.__class__.__name__}: {exc}", file=sys.stderr)
        if os.environ.get("MATERIAL_WORKSHOP_DEBUG") == "1":
            traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
