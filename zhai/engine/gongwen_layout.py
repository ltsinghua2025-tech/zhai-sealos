"""Conservative, offline government-document layout rules (profile 0.23).

Uncertain structures are protected rather than flattened into ordinary text.
No document content is invented, renumbered or sent to a model.
"""
from copy import deepcopy
import math
import re

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


DATE = re.compile(r'^(?:[0-9xX〇零一二三四五六七八九十]{4}年[0-9xX一二三四五六七八九十]{1,3}月[0-9xX一二三四五六七八九十]{1,3}日|\{\{(?:成文日期|日期)\}\})$')
ORG = re.compile(r'(?:单位|机关|部门|街道|办事处|政府|委员会|办公室|中心|学校|社区|财政局|[\u4e00-\u9fff]{1,12}局|公司|乡|镇)$')
TITLE_END = re.compile(r'(?:通知|通报|公告|通告|请示|报告|批复|决定|决议|意见|方案|计划|总结|纪要|规定|办法|细则|函)$')
HEADING = re.compile(r'^(?:[一二三四五六七八九十百]+、|[（(][一二三四五六七八九十百]+[）)]|\d+[.．、]|[（(]\d+[）)])')
ATTACHMENT = re.compile(r'^附件\s*[：:]')
TEXT_TAGS = {qn(t) for t in ('w:p', 'w:pPr', 'w:r', 'w:rPr', 'w:t', 'w:tab', 'w:br', 'w:cr')}


def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def complex_paragraph(p):
    # Hyperlinks, fields, drawings, comments, revisions and numbering have their
    # own semantics; even apparently empty paragraphs may carry important objects.
    if p._p.xpath('./w:pPr/w:numPr | ./w:pPr/w:sectPr'):
        return True
    if any(s.element.xpath('./w:pPr/w:numPr') for s in style_chain(p.style)):
        return True
    for child in p._p:
        if child.tag == qn('w:pPr'):
            continue
        for node in child.iter():
            if node.tag in (qn('w:rPr'),) or any(a.tag == qn('w:rPr') for a in node.iterancestors()):
                continue
            if node.tag not in TEXT_TAGS:
                return True
            if node.tag == qn('w:br') and node.get(qn('w:type'), 'textWrapping') != 'textWrapping':
                return True
    return False


def plain_empty(p):
    return not p.text.strip() and not complex_paragraph(p) and not p._p.xpath('.//w:br | .//w:cr')


def org_name(text):
    text = text.strip()
    return bool(0 < len(text) <= 60 and not re.search(r'[。！？：:；;\n]|请|报送|提交|联系|送至', text)
                and (ORG.search(text) or text in ('某单位', '{{主送单位}}', '{{发文机关}}')))


def recipient_piece(text):
    parts = [p for p in re.split(r'[、，,]', text.rstrip('：:、，,')) if p]
    return bool(parts) and all(org_name(part) for part in parts)


def layout_guard(document):
    if document.styles.element.xpath('./w:docDefaults/w:pPrDefault/w:pPr/w:ind'):
        return '文档默认样式含全局缩进，保留原版式，请先确认默认样式'
    if len(document.sections) > 1:
        return '多节文档保留原版式，尚未自动重排；请按正文与附件分别处理后核对'
    section = document.sections[0]
    if section.page_width > section.page_height or section._sectPr.xpath('./w:cols[@w:num != "1"]'):
        return '横向或分栏文档保留原版式，未执行竖向公文重排'
    if document._element.xpath('.//w:ins | .//w:del | .//w:sdt | .//w:object | .//w:pict'):
        return '文档含修订、内容控件或复杂对象，保留原版式，请确认结构后再排版'
    # Floating stamps/letterheads must never move through a global reflow.
    if document._element.xpath('.//wp:anchor'):
        return '文档含浮动印章或定位对象，保留原版式，未自动重排'
    for p in document.paragraphs[:8]:
        if re.search(r'[〔\[][0-9xX]{4}[〕\]]|签发人\s*[：:]', p.text) or (p.text.strip().endswith('文件') and org_name(p.text.strip()[:-2])):
            return '已有红头/发文字号结构，保留原版式；正文排版应在套红之前执行'
        for s in [p._p, *(v.element for v in style_chain(p.style))]:
            for color in s.xpath('.//w:color'):
                value = color.get(qn('w:val'), '')
                if re.fullmatch(r'[0-9A-Fa-f]{6}', value):
                    r, g, b = (int(value[i:i+2], 16) for i in (0, 2, 4))
                    if r > 150 and r > g * 1.5 and r > b * 1.5:
                        return '已有红头或红色标识，保留原版式；正文排版应在套红之前执行'
    return None


def classify(document, recipient_mode='auto'):
    ps = list(document.paragraphs)
    roles = ['empty' if plain_empty(p) else 'protected' if complex_paragraph(p) else 'body' for p in ps]
    notes = []
    nonempty = [i for i, p in enumerate(ps) if p.text.strip()]
    title = None
    if nonempty:
        i = nonempty[0]
        text = ps[i].text.strip()
        if roles[i] != 'protected' and not re.search(r'[。！？：:；;]', text) and len(text) <= 120:
            if TITLE_END.search(text) or ps[i].style.name in ('Title', '标题'):
                title = i
                roles[i] = 'title'
            else:
                roles[i] = 'unknown'
                notes.append(f'第{i+1}段标题身份不确定，已保留原格式')
    if title is not None:
        candidates = []
        for i in nonempty[1:5]:
            t = ps[i].text.strip()
            if roles[i] == 'protected' or not recipient_piece(t):
                break
            if candidates and ps[candidates[-1]]._p.getnext() is not ps[i]._p:
                break
            candidates.append(i)
            if t.endswith(('：', ':')):
                for j in candidates:
                    roles[j] = 'recipient' if recipient_mode == 'auto' else 'unknown'
                if recipient_mode == 'none':
                    notes.append('选择无主送，但发现疑似主送，已保留原段落而未改作正文')
                break
            if not t.endswith(('、', '，', ',')):
                for j in candidates:
                    roles[j] = 'unknown'
                notes.append('疑似主送缺少明确边界，已保留原格式')
                break
        else:
            for j in candidates:
                roles[j] = 'unknown'
        # Never treat unresolved recipient continuations as ordinary paragraphs.
        for j in candidates:
            if roles[j] == 'body':
                roles[j] = 'unknown'
                notes.append(f'第{j+1}段疑似主送，已保留原格式')

    # Locate dates near the tail; only an adjacent organization forms a signature.
    signature = []
    date_index = None
    for i in nonempty[-6:]:
        if not DATE.fullmatch(ps[i].text.strip()) or roles[i] == 'protected':
            continue
        roles[i] = 'date'
        date_index = i
        prev = [j for j in nonempty if j < i]
        for j in reversed(prev):
            if roles[j] in ('protected', 'title', 'recipient') or not org_name(ps[j].text):
                break
            between = ps[j]._p.getnext()
            if between is not ps[signature[-1] if signature else i]._p:
                break
            signature.append(j)
            roles[j] = 'signature'
        break
    signature.reverse()
    if len(signature) > 1:
        notes.append('疑似联合行文落款，保留原布局，不套用单一机关规则')
    for i, p in enumerate(ps):
        t = p.text.strip()
        if roles[i] != 'body':
            continue
        if ATTACHMENT.match(t):
            roles[i] = 'attachment'
        elif p.alignment in (WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER) or (i in nonempty[-3:] and org_name(t)):
            roles[i] = 'unknown'
            notes.append(f'第{i+1}段为疑似落款或特殊对齐，保留原格式')
        elif HEADING.match(t) and len(t) <= 40 and not re.search('[。！？；;]', t):
            # Arabic numbering alone is insufficient to distinguish a list from a heading.
            if re.match(r'^[一二三四五六七八九十百]+、', t):
                roles[i] = 'heading1'
            elif re.match(r'^[（(][一二三四五六七八九十百]+[）)]', t):
                roles[i] = 'heading2'
            elif any(s.name.startswith('Heading') for s in style_chain(p.style)):
                roles[i] = 'heading3'
    # Protect all attachment continuations before any body rules are applied.
    for i in range(len(ps)):
        if roles[i] != 'attachment':
            continue
        j = i + 1
        while j < len(ps) and ps[j-1]._p.getnext() is ps[j]._p and re.match(r'^\s*\d+[.．、]', ps[j].text):
            roles[j] = 'attachment'
            j += 1
    return ps, roles, title, signature, date_index, notes


def clean_whitespace(document, collapse=True, remove_empty=True):
    """Only remove redundant empty paragraphs, never structural/role boundaries."""
    guard = layout_guard(document)
    if guard:
        return 0, ['空白清理未执行：' + guard]
    ps, roles, _, _, _, notes = classify(document)
    changed = 0
    for i, p in enumerate(ps):
        if roles[i] not in ('body', 'empty'):
            continue
        if collapse and roles[i] == 'body' and not complex_paragraph(p):
            for node in p._p.xpath('./w:r/w:t'):
                # Text nodes only: assigning run.text destroys drawings and fields.
                text, count = re.subn(r' {2,}', ' ', node.text or '')
                if count:
                    node.text = text
                    changed += count
        if not remove_empty or roles[i] != 'empty' or not plain_empty(p):
            continue
        prev = next((roles[j] for j in range(i-1, -1, -1) if roles[j] != 'empty'), None)
        nxt = next((roles[j] for j in range(i+1, len(ps)) if roles[j] != 'empty'), None)
        body_roles = {'body', 'heading1', 'heading2', 'heading3'}
        if prev in body_roles and nxt in body_roles and p._p.getprevious() is not None and p._p.getnext() is not None:
            if p._p.getprevious().tag == qn('w:p') and p._p.getnext().tag == qn('w:p'):
                p._p.getparent().remove(p._p)
                changed += 1
    return changed, notes


def effective_bold(p, run):
    if run.bold is not None:
        return run.bold
    for s in list(style_chain(run.style)) + list(style_chain(p.style)):
        if s.font.bold is not None:
            return s.font.bold
    return None


def reset_indents(p):
    original = p.style
    styles = p.part.document.styles
    if not original.name.startswith('zhai公文-'):
        name = 'zhai公文-' + original.style_id
        if name not in styles:
            derived = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            for s in reversed(list(style_chain(original))):
                for part in ('pPr', 'rPr'):
                    src = s.element.find(qn('w:' + part))
                    if src is None:
                        continue
                    dst = getattr(derived.element, 'get_or_add_' + part)()
                    for child in src:
                        if part == 'pPr' and child.tag in (qn('w:ind'), qn('w:tabs'), qn('w:numPr')):
                            continue
                        existing = dst.find(child.tag)
                        if existing is not None:
                            dst.remove(existing)
                        dst.append(deepcopy(child))
        p.style = styles[name]
    pr = p._p.get_or_add_pPr()
    for node in list(pr):
        if node.tag in (qn('w:ind'), qn('w:tabs')):
            pr.remove(node)
    f = p.paragraph_format
    f.left_indent = f.right_indent = f.first_line_indent = Pt(0)


def strip_padding(p):
    remaining = len(p.text) - len(p.text.lstrip(' \t\u3000'))
    for node in p._p.xpath('./w:r/w:t | ./w:r/w:tab'):
        if not remaining:
            break
        if node.tag == qn('w:tab'):
            node.getparent().remove(node)
            remaining -= 1
        else:
            value = node.text or ''
            n = min(remaining, len(value))
            node.text = value[n:]
            remaining -= n


def add_page_numbers(document, set_font):
    # An existing footer is authoritative, including sdt-wrapped fields and
    # first-page/odd-even settings. Do not create a second, possibly overlapping footer.
    refs = document._element.xpath('.//w:footerReference')
    if refs:
        return ['已有页脚或页码，完整保留；未覆盖、未重复添加，请核对奇偶页设置']
    sec = document.sections[0]
    if sec.different_first_page_header_footer:
        return ['首页不同设置已保留，未自动增加页码']
    if document.settings.element.find(qn('w:evenAndOddHeaders')) is None:
        document.settings.element.append(OxmlElement('w:evenAndOddHeaders'))
    sec.footer_distance = Cm(1.75)
    for footer, right in ((sec.footer, True), (sec.even_page_footer, False)):
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if right else WD_ALIGN_PARAGRAPH.LEFT
        f = p.paragraph_format
        f.space_before = f.space_after = Pt(0)
        f.line_spacing = Pt(28)
        f.right_indent = Pt(14) if right else Pt(0)
        f.left_indent = Pt(0) if right else Pt(14)
        set_font(p.add_run('— '), '宋体', 14)
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = '1'
        r.append(t)
        field.append(r)
        p._p.append(field)
        from docx.text.run import Run
        set_font(Run(r, p), '宋体', 14)
        set_font(p.add_run(' —'), '宋体', 14)
    return []


def apply_layout(document, parameters, set_font, normalize_attachments):
    params = parameters or {}
    recipient = params.get('recipientMode', 'auto')
    signature_mode = params.get('signatureMode', 'preserve')
    if recipient not in ('auto', 'none') or signature_mode not in ('preserve', 'unsigned'):
        raise ValueError('未知的主送或落款模式')
    try:
        line = float(params.get('lineSpacingPt', 28))
    except (ValueError, TypeError) as exc:
        raise ValueError('行距必须为12—60之间的有限数值') from exc
    if not math.isfinite(line) or not 12 <= line <= 60:
        raise ValueError('行距必须为12—60之间的有限数值')
    guard = layout_guard(document)
    if guard:
        return 0, ['排版提示：' + guard, '原件未覆盖；本次仅生成保留版式的副本，未完成公文重排']
    cleaned, clean_notes = clean_whitespace(document, collapse=False, remove_empty=True)
    ps, roles, title, signatures, date, notes = classify(document, recipient)
    notes.extend(clean_notes)
    fonts = {k: str(params.get(k) or v).strip() for k, v in {
        'titleFont': '方正小标宋简体', 'bodyFont': '仿宋_GB2312',
        'firstHeadingFont': '黑体', 'secondHeadingFont': '楷体_GB2312'}.items()}
    sec = document.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(3.7), Cm(3.5)
    sec.left_margin, sec.right_margin = Cm(2.8), Cm(2.6)
    preserve = bool(params.get('preserveBold', True))
    changed = cleaned
    for i, p in enumerate(ps):
        role = roles[i]
        if role in ('empty', 'unknown', 'protected', 'signature', 'date', 'attachment'):
            if role == 'protected':
                notes.append(f'第{i+1}段含编号/域/链接等结构，已保留原格式')
            continue
        original_bold = [effective_bold(p, r) for r in p.runs]
        reset_indents(p)
        strip_padding(p)
        f = p.paragraph_format
        f.first_line_indent = Pt(0 if role in ('title', 'recipient') else 32)
        f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        f.line_spacing = Pt(line)
        f.space_before = f.space_after = Pt(0)
        f.widow_control = True
        f.keep_with_next = role in ('title', 'heading1', 'heading2', 'heading3', 'recipient')
        has_soft_break = bool(p._p.xpath('.//w:br | .//w:cr'))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if role == 'title' else WD_ALIGN_PARAGRAPH.LEFT if role == 'recipient' or has_soft_break else WD_ALIGN_PARAGRAPH.JUSTIFY
        if has_soft_break:
            notes.append(f'第{i+1}段含手动换行，保留换行并左对齐，未自动拆分标题、正文或落款')
        font = fonts[{'title': 'titleFont', 'heading1': 'firstHeadingFont', 'heading2': 'secondHeadingFont'}.get(role, 'bodyFont')]
        size = 22 if role == 'title' else 16
        for r, bold in zip(p.runs, original_bold):
            set_font(r, font, size, False)
            r.font.color.rgb = RGBColor(0, 0, 0)
            r.bold = bold if preserve else False
        changed += 1
    if title is not None and params.get('addTitleBlankLine', True):
        p = ps[title]
        following = p._p.getnext()
        if following is not None and title + 1 < len(ps) and following is ps[title+1]._p and plain_empty(ps[title+1]):
            while title + 2 < len(ps) and following.getnext() is ps[title+2]._p and plain_empty(ps[title+2]):
                following.getparent().remove(ps[title+2]._p)
                ps.pop(title+2)
                changed += 1
        if following is not None and not (title + 1 < len(ps) and plain_empty(ps[title+1]) and following is ps[title+1]._p):
            spacer = OxmlElement('w:p')
            p._p.addnext(spacer)
            from docx.text.paragraph import Paragraph
            f = Paragraph(spacer, p._parent).paragraph_format
            f.line_spacing = Pt(line)
            f.space_before = f.space_after = Pt(0)
            f.keep_with_next = True
    if signatures or date is not None:
        if signature_mode == 'unsigned' and len(signatures) == 1:
            sig, dp = ps[signatures[0]], ps[date]
            # Point estimates follow the specified Latin/CJK fonts. Exact glyph
            # widths still require checking in the target Office renderer.
            date_width = sum(0.5 if c.isascii() else 1 for c in dp.text.strip()) * 16
            org_width = len(sig.text.strip()) * 16
            date_right = 32 if date_width > org_width else max(0, org_width - date_width)
            sig_right = 32 + max(0, (date_width - org_width) / 2)
            for p, right in ((sig, sig_right), (dp, date_right)):
                bolds = [effective_bold(p, r) for r in p.runs]
                reset_indents(p)
                strip_padding(p)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                f = p.paragraph_format
                f.right_indent = Pt(right)
                f.line_spacing = Pt(line)
                f.space_before = f.space_after = Pt(0)
                f.keep_together = True
                f.keep_with_next = p is sig
                for r, bold in zip(p.runs, bolds):
                    set_font(r, fonts['bodyFont'], 16, False)
                    r.bold = bold if preserve else False
            prev = signatures[0] - 1
            if prev < 0 or not plain_empty(ps[prev]) or ps[prev]._p.getnext() is not sig._p:
                sig.paragraph_format.space_before = Pt(line)
            else:
                ps[prev].paragraph_format.line_spacing = Pt(line)
                ps[prev].paragraph_format.space_before = ps[prev].paragraph_format.space_after = Pt(0)
                ps[prev].paragraph_format.keep_with_next = True
            notes.append('明确无章单一机关：已调整落款间隔、署名日期布局；内容未改写')
        elif len(signatures) <= 1:
            notes.append('落款盖章方式或署名结构未明确，已保留原布局；需补空行时选择“无章单一机关”')
    # Existing numbered implementation only receives safe contiguous groups.
    for i, p in enumerate(ps):
        if roles[i] != 'attachment':
            continue
        if complex_paragraph(p):
            notes.append(f'第{i+1}段附件含复杂结构，已保留')
            continue
        t = p.text.strip()
        if ATTACHMENT.match(t) and not re.match(r'^附件\s*[：:]\s*\d+[.．、]', t):
            if re.match(r'^附件\s*[：:]\s*\d', t):
                notes.append('附件编号边界不明确，保留并请人工核对')
                continue
            bolds = [effective_bold(p, r) for r in p.runs]
            reset_indents(p)
            strip_padding(p)
            f = p.paragraph_format
            f.left_indent, f.first_line_indent = Pt(80), Pt(-48)
            f.line_spacing = Pt(line)
            f.space_before = Pt(0 if i and plain_empty(ps[i-1]) else line)
            f.space_after = Pt(0)
            f.keep_together = len(t) <= 120
            f.keep_with_next = False
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r, bold in zip(p.runs, bolds):
                set_font(r, fonts['bodyFont'], 16, False)
                r.bold = bold if preserve else False
            changed += 1
    changed += normalize_attachments(ps, preserve)
    for i, p in enumerate(ps):
        if roles[i] == 'attachment' and p.style.name.startswith('zhai附件说明-'):
            f = p.paragraph_format
            f.line_spacing = Pt(line)
            f.space_before = Pt(line if ATTACHMENT.match(p.text.strip()) and (not i or not plain_empty(ps[i-1])) else 0)
            f.space_after = Pt(0)
            for r in p.runs:
                bold = effective_bold(p, r)
                set_font(r, fonts['bodyFont'], 16, False)
                r.bold = bold if preserve else False
    if params.get('formatTables', False):
        for t in document.tables:
            # Never reformat fixed-height or layout/banji tables.
            if t._tbl.xpath('.//w:trHeight | ./w:tblPr/w:tblpPr | .//w:drawing | .//w:pict') or re.search('印发|抄送', ''.join(c.text for r in t.rows for c in r.cells)):
                notes.append('定位/版记/固定行高表格已保留，未套正文规则')
                continue
            seen = set()
            for row in t.rows:
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    for p in cell.paragraphs:
                        if complex_paragraph(p):
                            continue
                        for r in p.runs:
                            # Font family only, preserving size, spacing and emphasis.
                            r.font.name = fonts['bodyFont']
                            r._r.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fonts['bodyFont'])
                            if not preserve:
                                r.bold = False
        notes.append('表格仅规范安全区域字体名，保留原字号、行距和几何布局')
    elif document.tables:
        notes.append('表格原样保留，未强制套用正文三号字/固定行距')
    if params.get('applyPageNumbers', True):
        notes.extend(add_page_numbers(document, set_font))
    notes.insert(0, f'公文安全排版0.23：调整{changed}段；保留原有加粗={"是" if preserve else "否"}；原件未覆盖')
    notes.append('已做结构保护；实际分页和字体替代仍需在目标办公软件核对，不等同于逐页视觉验收')
    notes = list(dict.fromkeys(notes))
    return changed, [notes[0], *('排版提示：' + n for n in notes[1:])]
