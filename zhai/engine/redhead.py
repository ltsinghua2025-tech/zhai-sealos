#!/usr/bin/env python3
"""Offline letterhead composition. Template artwork remains in its original package.

GB/T 9704—2012 sections 5.2, 7.2–7.5. Templates may intentionally differ.
"""
from __future__ import annotations

import re
from copy import deepcopy
from datetime import date
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.text.paragraph import Paragraph

DEFAULT_UNIT = "中共珠海市xxx"
DEFAULT_DATE = "202x年x月x日印发"


def imprint_values(params: dict) -> tuple[str, str]:
    unit = str(params.get("imprintUnit") or params.get("unitName") or DEFAULT_UNIT).strip()
    office = str(params.get("officeName") or "").strip()
    if office and not unit.endswith(office):
        unit += office
    stamp = str(params.get("printDate") or DEFAULT_DATE).strip()
    if stamp != DEFAULT_DATE:
        m = re.fullmatch(r"(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})日?(?:印发)?", stamp)
        if not m:
            raise ValueError("印发日期请填写有效年月日，例如 2026年9月15日")
        year, month, day = map(int, m.groups())
        date(year, month, day)
        stamp = f"{year}年{month}月{day}日印发"
    if len(unit) > 60:
        raise ValueError("印发机关名称超过 60 字，请确认名称")
    return unit, stamp


def _text(element) -> str:
    return "".join(t.text or "" for t in element.iter(qn('w:t')))


def _red(element) -> bool:
    xml = element.xml.lower()
    return any(token in xml for token in ('ff0000', 'fillcolor="red"', 'color="red"', 'c00000'))


def _plain_replace(paragraph, value):
    # Keep the first run's typography, but remove old field instructions too.
    props = deepcopy(paragraph.runs[0]._r.rPr) if paragraph.runs and paragraph.runs[0]._r.rPr is not None else None
    paragraph.clear()
    r = paragraph.add_run(value)
    if props is not None:
        r._r.insert(0, props)


def _border(paragraph, edge="bottom", color="000000", size=8):
    pp = paragraph._p.get_or_add_pPr()
    borders = pp.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        pp.append(borders)
    el = OxmlElement(f"w:{edge}")
    for k, v in {"val": "single", "sz": str(size), "color": color, "space": "0"}.items():
        el.set(qn(f"w:{k}"), v)
    borders.append(el)


def _prepare_body(source: Path, scratch: Path, engine, params):
    notes = []
    if source.suffix.lower() == '.doc':
        engine.extract_legacy_doc_direct(source)
        converted = scratch / 'body.docx'
        notes += engine.convert_office(source, converted, 'doc-to-docx')['notes']
        source = converted
    doc = Document(source)
    if len(doc.sections) > 1:
        raise ValueError("正文包含多个分节，请先拆分为单节文件；避免套红时改变横向页或页眉")
    if doc.element.xpath('.//w:ins | .//w:del'):
        raise ValueError("正文含未接受的修订，请先确认修订后再套红")
    title = next((p for p in doc.paragraphs if p.text.strip()), None)
    if title is None:
        raise ValueError("未找到正文标题")
    if _red(title._p):
        raise ValueError("正文首段已有红色文头，请使用未套红正文，避免重复套红")
    _, format_notes = engine.format_existing_docx_as_gongwen(doc, {**params, 'applyPageNumbers': True, 'addTitleBlankLine': False})
    # Inspect the composed file once, including the retained template fonts.
    # Main addressee has no first-line indent. Title and addressee stay with body.
    nonempty = [p for p in doc.paragraphs if p.text.strip()]
    title.paragraph_format.keep_with_next = True
    title.paragraph_format.space_after = Pt(28)
    # Suppress inherited decorative title borders from source styles.
    borders = OxmlElement('w:pBdr')
    for edge in ('top','left','bottom','right','between','bar'):
        el = OxmlElement('w:'+edge); el.set(qn('w:val'),'nil'); borders.append(el)
    title._p.get_or_add_pPr().append(borders)
    if len(nonempty) > 1 and nonempty[1].text.rstrip().endswith(('：', ':')):
        nonempty[1].paragraph_format.first_line_indent = Pt(0)
        nonempty[1].paragraph_format.keep_with_next = True
    # Preserve signature text, but place date and immediately preceding unit at right.
    for i, p in enumerate(nonempty):
        if re.fullmatch(r'\d{4}年\d{1,2}月\d{1,2}日', p.text.strip()):
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.right_indent = Pt(64)
            if i > 1:
                nonempty[i-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                nonempty[i-1].paragraph_format.first_line_indent = Pt(0)
                nonempty[i-1].paragraph_format.keep_with_next = True
        elif '\n' in p.text and re.search(r'\d{4}年\d{1,2}月\d{1,2}日$', p.text.strip()):
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.right_indent = Pt(32)
    return doc, notes


def _template(path, scratch, engine, params):
    notes = []
    if path.suffix.lower() == '.doc':
        engine.extract_legacy_doc_direct(path)
        converted = scratch / 'template.docx'
        notes += engine.convert_office(path, converted, 'doc-to-docx')['notes']
        path = converted
    doc = Document(path)
    if len(doc.sections) != 1:
        raise ValueError("红头模板须为单节；请先保留需要使用的文头")
    elements = list(doc.element.body)[:-1]
    slots = [i for i, e in enumerate(elements) if '{{正文}}' in _text(e)]
    # Explicit slot takes priority; otherwise detect the contiguous red masthead.
    red_indices = [i for i, e in enumerate(elements[:10]) if _red(e)]
    header_red = any(_red(h._element) for h in (doc.sections[0].header, doc.sections[0].first_page_header))
    if not slots and not red_indices and not header_red:
        raise ValueError("无法识别红头边界：请在模板正文起点单独一段写 {{正文}}，或使用含红色文头的模板")
    end = slots[0] if slots else (max(red_indices)+1 if red_indices else 0)
    # Keep document-number paragraph between masthead and body, excluding instructions.
    if not slots:
        for i in range(end, min(end+4, len(elements))):
            if re.search(r'〔.*〕.*号', _text(elements[i])):
                end = i+1
                break
    imprint = None
    for e in elements[end:]:
        if '印发' not in _text(e):
            continue
        # Retain real group/table/paragraph, excluding instruction callouts in same anchor.
        candidate = deepcopy(e)
        if candidate.tag == qn('w:p'):
            for r in list(candidate):
                if r.tag != qn('w:r'):
                    continue
                for child in list(r):
                    if child.tag != qn('w:rPr') and '印发' not in _text(child):
                        r.remove(child)
        imprint = candidate
        break
    for e in elements[end:]:
        doc.element.body.remove(e)
    for p in doc.paragraphs:
        if p.text.strip() in ('（空一行）', '(空一行)'):
            p.clear()
        if re.search(r'〔.*〕.*号', p.text):
            _plain_replace(p, str(params.get('documentNumber') or ''))
    # Existing footer page fields remain; comments and teaching callouts aren't body.
    notes.append('已复用上传模板的原始文头对象、红线及页面尺寸')
    return doc, imprint, notes


def _new_head(body, engine, params):
    doc = Document()
    engine.format_existing_docx_as_gongwen(doc, {'applyPageNumbers': True})
    unit = str(params.get('unitName') or DEFAULT_UNIT).strip()
    if not unit or len(unit) > 60:
        raise ValueError('请填写不超过 60 字的文头单位名称')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Mm(35)
    p.paragraph_format.space_after = Pt(56)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.15
    # Font size adapts to one or two lines without horizontal stretching.
    size = 34 if len(unit) <= 13 else 28 if len(unit) <= 16 else 26
    r = p.add_run(unit)
    engine._set_run_font(r, '方正小标宋简体', size)
    r.font.color.rgb = RGBColor.from_string('FF0000')
    p = doc.add_paragraph(str(params.get('documentNumber') or ''))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_after = Mm(4)
    for r in p.runs:
        engine._set_run_font(r, '仿宋_GB2312', 16)
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = Pt(1)
    p.paragraph_format.space_after = Pt(56)
    p.paragraph_format.keep_with_next = True
    _border(p, color='FF0000', size=6)
    return doc


def _imprint(doc, element, engine, params):
    unit, stamp = imprint_values(params)
    # Reuse typography and rule weight, laying out the imprint as a Word table.
    # Legacy floating drawing groups can detach text from their rules on repagination.
    imprint_runs = []
    rule_size = 8
    if element is not None:
        for p in element.iter(qn('w:p')):
            if _text(p).strip():
                runs = list(p.iter(qn('w:rPr')))
                if runs: imprint_runs.append(deepcopy(runs[0]))
        strokes = list(element.iter('{urn:schemas-microsoft-com:vml}stroke'))
        if strokes:
            weight = strokes[0].get('weight','')
            if weight.isdigit(): rule_size = max(2, round(int(weight)/12700*8))
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    width = doc.sections[-1].page_width-doc.sections[-1].left_margin-doc.sections[-1].right_margin
    table.columns[0].width = int(width * .60)
    table.columns[1].width = int(width * .40)
    for i, value in enumerate((unit, stamp)):
        cell = table.cell(0, i)
        cell.width = table.columns[i].width
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = Pt(21)
        engine._set_run_font(p.add_run(value), '仿宋_GB2312', 14)
        # The supplied redhead remains exact; the new imprint uses the standard
        # 14pt font consistently even when legacy samples mixed 16pt and empty runs.
        cell_borders = OxmlElement('w:tcBorders')
        for name in ('top','bottom'):
            edge = OxmlElement('w:'+name)
            for k,v in {'val':'single','sz':str(rule_size),'color':'000000'}.items(): edge.set(qn('w:'+k),v)
            cell_borders.append(edge)
        cell._tc.get_or_add_tcPr().append(cell_borders)
    pr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for name in ('top', 'bottom'):
        edge = OxmlElement(f'w:{name}')
        for k,v in {'val':'single','sz':str(rule_size),'color':'000000'}.items(): edge.set(qn(f'w:{k}'),v)
        borders.append(edge)
    pr.insert_element_before(borders, 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook')
    margin = OxmlElement('w:tblCellMar')
    for name in ('left','right'):
        el = OxmlElement(f'w:{name}')
        el.set(qn('w:w'),'280'); el.set(qn('w:type'),'dxa'); margin.append(el)
    pr.insert_element_before(margin, 'w:tblLook')
    floating = OxmlElement('w:tblpPr')
    for k,v in {'vertAnchor':'margin','horzAnchor':'margin','tblpX':'280','tblpYSpec':'bottom','topFromText':'160','bottomFromText':'0','leftFromText':'0','rightFromText':'0'}.items():
        floating.set(qn('w:'+k),v)
    pr.insert(0, floating)


def _copy_header_footer(source, target):
    """Keep template header/footer content when enabling a distinct first page."""
    for child in list(target._element):
        target._element.remove(child)
    for child in source._element:
        copied = deepcopy(child)
        for node in copied.iter():
            for attribute, value in list(node.attrib.items()):
                if attribute in (qn('r:id'), qn('r:embed'), qn('r:link')):
                    relation = source.part.rels.get(value)
                    if relation is not None:
                        node.set(attribute, target.part.relate_to(
                            relation.target_part if not relation.is_external else relation.target_ref,
                            relation.reltype, relation.is_external,
                        ))
        target._element.append(copied)


def _remove_page_fields(footer):
    for paragraph in footer._element.xpath('.//w:p'):
        for field in list(paragraph.xpath('.//w:fldSimple')):
            if 'PAGE' in (field.get(qn('w:instr')) or '').upper():
                field.getparent().remove(field)
        runs = list(paragraph.xpath('./w:r'))
        begin = None
        for index, run in enumerate(runs):
            types = [node.get(qn('w:fldCharType')) for node in run.xpath('./w:fldChar')]
            if 'begin' in types:
                begin = index
            if 'end' in types and begin is not None:
                instruction = ''.join(
                    node.text or '' for item in runs[begin:index+1]
                    for node in item.xpath('./w:instrText')
                )
                if 'PAGE' in instruction.upper():
                    for item in runs[begin:index+1]:
                        if item.getparent() is paragraph:
                            paragraph.remove(item)
                begin = None
        text = ''.join(node.text or '' for node in paragraph.xpath('.//w:t')).strip()
        if text.replace('—', '').replace('-', '').strip() == '':
            for run in list(paragraph.xpath('./w:r')):
                paragraph.remove(run)


def _hide_first_page_number(document):
    section = document.sections[0]
    if not section.different_first_page_header_footer:
        _copy_header_footer(section.header, section.first_page_header)
        _copy_header_footer(section.footer, section.first_page_footer)
        section.different_first_page_header_footer = True
    _remove_page_fields(section.first_page_footer)


def _ensure_later_page_numbers(document, engine):
    """Fill only missing ordinary-page number fields; never restart numbering."""
    section = document.sections[0]
    settings = document.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        settings.append(OxmlElement('w:evenAndOddHeaders'))
    for footer, align in ((section.footer, WD_ALIGN_PARAGRAPH.RIGHT),
                          (section.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT)):
        if 'PAGE' in footer._element.xml.upper():
            continue
        paragraph = footer.add_paragraph() if any(p.text.strip() for p in footer.paragraphs) else footer.paragraphs[0]
        paragraph.alignment = align
        paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_after = Pt(0)
        engine._set_run_font(paragraph.add_run('— '), '宋体', 14)
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')
        run = OxmlElement('w:r')
        value = OxmlElement('w:t')
        value.text = '1'
        run.append(value)
        field.append(run)
        paragraph._p.append(field)
        engine._set_run_font(paragraph.add_run(' —'), '宋体', 14)


def compose(source: Path, target: Path, params: dict, engine) -> dict:
    import tempfile
    from docxcompose.composer import Composer
    imprint_values(params)  # Validate before any file conversion.
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.redhead-', dir=target.parent) as temp:
        scratch = Path(temp)
        body, notes = _prepare_body(source, scratch, engine, params)
        original_text = [p.text for p in body.paragraphs if p.text.strip()]
        template = params.get('_templatePath')
        if template:
            head, imprint, extra = _template(Path(template), scratch, engine, params)
            notes += extra
        else:
            head, imprint = _new_head(body, engine, params), None
        # Template headers and relationships are retained by composing into the template.
        composer = Composer(head)
        composer.append(body)
        _imprint(head, imprint, engine, params)
        _ensure_later_page_numbers(head, engine)
        _hide_first_page_number(head)
        settings = head.settings.element
        update = OxmlElement('w:updateFields'); update.set(qn('w:val'), 'true'); settings.append(update)
        engine.atomic_writer(target, lambda p: composer.save(p))
        reopened = Document(target)
        actual = [p.text for p in reopened.paragraphs]
        if not all(text in actual for text in original_text):
            target.unlink(missing_ok=True)
            raise ValueError('套红后正文完整性检查失败')
        notes += engine.font_notices(target)
        if not params.get('printDate'):
            notes.append('印发日期未提供，已保留占位日期，请正式印发前补全')
        return {'changes':1, 'notes':list(dict.fromkeys(notes)), 'verification':'redhead_docx_reopened;body_preserved;template_readonly;GB_T_9704_2012'}
