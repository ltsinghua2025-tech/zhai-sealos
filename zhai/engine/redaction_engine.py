#!/usr/bin/env python3
"""Lightweight, offline Chinese PII detection without NLP models."""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class RedactionHit:
    start: int
    end: int
    entity_type: str
    replacement: str
    confidence: float
    reason: str


EMAIL = re.compile(r"(?<![A-Z0-9._%+-])[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}(?![A-Z0-9_-])", re.IGNORECASE)
MOBILE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
LANDLINE = re.compile(r"(?<!\d)(?:0\d{2,3}[-－— ]?)?\d{7,8}(?:[-－— ]\d{1,6})?(?!\d)")
ID_CARD = re.compile(r"(?<![0-9A-Za-z])\d{17}[0-9Xx](?![0-9A-Za-z])")
BANK_CARD = re.compile(r"(?<!\d)(?:\d[ -]?){15,18}\d(?!\d)")
CREDIT_CODE = re.compile(r"(?<![0-9A-Z])[0-9A-HJ-NPQRTUWXY]{18}(?![0-9A-Z])", re.IGNORECASE)
PASSPORT = re.compile(r"(?<![0-9A-Za-z])(?:[EeGgDdSsPp]\d{7,8}|[HhMm]\d{8,10})(?![0-9A-Za-z])")
PLATE = re.compile(r"(?<![\u3400-\u9fffA-Z0-9])[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-Z][A-HJ-NP-Z0-9]{5,6}(?![A-Z0-9])", re.IGNORECASE)
IPV4 = re.compile(r"(?<!\d)(?:\d{1,3}\.){3}\d{1,3}(?!\d)")

PERSON_LABEL = re.compile(r"(?:姓名|联系人|经办人|负责人|申请人|被申请人|甲方|乙方|填表人)(?:\s*[：:]\s*|\s+)([\u3400-\u9fff·]{2,8})")
PERSON_COMRADE = re.compile(r"(?<![\u3400-\u9fff])([\u3400-\u9fff]{2,4})(?=同志)")
ADDRESS_LABEL = re.compile(r"(?:家庭住址|居住地址|通讯地址|联系地址|注册地址|办公地址|住所地?|住址|地址)(?:\s*[：:]\s*|\s+)([^\n；;,，。]{4,80})")
ORGANIZATION_LABEL = re.compile(r"(?:工作单位|所在单位|任职单位|单位名称|机构名称|单位)(?:\s*[：:]\s*|\s+)([^\n；;,，。]{2,80})")

CONTEXT_TYPES = {
    "person": ("姓名", "联系人", "经办人", "负责人", "申请人", "填表人"),
    "address": ("住址", "地址", "住所", "居住地", "通讯地址"),
    "organization": ("工作单位", "所在单位", "任职单位", "单位名称", "机构名称"),
    "mobile": ("手机号", "手机", "联系电话", "联系方式", "电话"),
    "id_card": ("身份证", "身份证号", "公民身份号码", "证件号码"),
    "bank_card": ("银行卡", "银行账号", "卡号", "账户"),
    "email": ("邮箱", "电子邮箱", "邮件地址"),
    "credit_code": ("统一社会信用代码", "信用代码"),
    "passport": ("护照", "护照号"),
    "plate": ("车牌", "车牌号"),
}


def looks_like_chinese_id(value: str) -> bool:
    value = value.upper()
    if not re.fullmatch(r"\d{17}[0-9X]", value):
        return False
    try:
        year, month, day = int(value[6:10]), int(value[10:12]), int(value[12:14])
        from datetime import date

        date(year, month, day)
    except ValueError:
        return False
    return True


def valid_chinese_id(value: str) -> bool:
    value = value.upper()
    if not looks_like_chinese_id(value):
        return False
    weights = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
    checks = "10X98765432"
    return checks[sum(int(value[i]) * weights[i] for i in range(17)) % 11] == value[-1]


def valid_luhn(value: str) -> bool:
    digits = re.sub(r"\D", "", value)
    if not 16 <= len(digits) <= 19 or len(set(digits)) == 1:
        return False
    total = 0
    parity = len(digits) % 2
    for index, character in enumerate(digits):
        number = int(character)
        if index % 2 == parity:
            number *= 2
            if number > 9:
                number -= 9
        total += number
    return total % 10 == 0


def valid_credit_code(value: str) -> bool:
    value = value.upper()
    alphabet = "0123456789ABCDEFGHJKLMNPQRTUWXY"
    if len(value) != 18 or any(character not in alphabet for character in value):
        return False
    weights = (1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28)
    check_index = (31 - sum(alphabet.index(value[i]) * weights[i] for i in range(17)) % 31) % 31
    return alphabet[check_index] == value[-1]


def valid_ipv4(value: str) -> bool:
    try:
        return all(0 <= int(part) <= 255 for part in value.split("."))
    except ValueError:
        return False


def context_type(context: str | None) -> str | None:
    normalized = re.sub(r"[\s：:（）()]+", "", context or "")
    for entity_type, labels in CONTEXT_TYPES.items():
        if normalized in labels:
            return entity_type
    return None


def _add(hits: list[RedactionHit], match: re.Match[str], entity_type: str, replacement: str, confidence: float, reason: str, group: int = 0) -> None:
    start, end = match.span(group)
    if start != end:
        hits.append(RedactionHit(start, end, entity_type, replacement, confidence, reason))


def detect_pii(text: str, context: str | None = None) -> list[RedactionHit]:
    hits: list[RedactionHit] = []
    for match in EMAIL.finditer(text):
        _add(hits, match, "email", "[邮箱已脱敏]", 0.99, "邮箱格式")
    for match in MOBILE.finditer(text):
        _add(hits, match, "mobile", "[手机号已脱敏]", 0.98, "中国大陆手机号格式")
    for match in LANDLINE.finditer(text):
        # Seven/eight-digit sequences collide easily with dates and document
        # numbers, so only accept them beside an explicit telephone label.
        nearby = text[max(0, match.start() - 10) : match.start()]
        if any(label in nearby for label in CONTEXT_TYPES["mobile"]):
            _add(hits, match, "landline", "[电话已脱敏]", 0.9, "电话标签与固定电话格式")
    for match in ID_CARD.finditer(text):
        if valid_chinese_id(match.group()):
            _add(hits, match, "id_card", "[身份证号已脱敏]", 1.0, "身份证日期和GB 11643校验位通过")
    for match in BANK_CARD.finditer(text):
        compact = re.sub(r"\D", "", match.group())
        if not looks_like_chinese_id(compact) and valid_luhn(match.group()):
            _add(hits, match, "bank_card", "[银行卡号已脱敏]", 0.99, "银行卡Luhn校验通过")
    for match in CREDIT_CODE.finditer(text):
        if not looks_like_chinese_id(match.group()) and valid_credit_code(match.group()):
            _add(hits, match, "credit_code", "[信用代码已脱敏]", 1.0, "统一社会信用代码校验位通过")
    for match in PASSPORT.finditer(text):
        _add(hits, match, "passport", "[护照号已脱敏]", 0.88, "中国护照/通行证格式")
    for match in PLATE.finditer(text):
        _add(hits, match, "plate", "[车牌号已脱敏]", 0.9, "中国大陆车牌格式")
    for match in IPV4.finditer(text):
        if valid_ipv4(match.group()):
            _add(hits, match, "ipv4", "[IP地址已脱敏]", 0.95, "IPv4格式")

    for match in PERSON_LABEL.finditer(text):
        _add(hits, match, "person", "[姓名已脱敏]", 0.9, "姓名标签上下文", group=1)
    for match in PERSON_COMRADE.finditer(text):
        _add(hits, match, "person", "[姓名已脱敏]", 0.72, "“同志”上下文", group=1)
    for match in ADDRESS_LABEL.finditer(text):
        _add(hits, match, "address", "[地址已脱敏]", 0.88, "地址标签上下文", group=1)
    for match in ORGANIZATION_LABEL.finditer(text):
        _add(hits, match, "organization", "[单位已脱敏]", 0.82, "单位标签上下文", group=1)

    kind = context_type(context)
    stripped = text.strip()
    if stripped and kind:
        span = (text.index(stripped), text.index(stripped) + len(stripped))
        valid = False
        replacement = "[敏感信息已脱敏]"
        confidence = 0.9
        if kind == "person" and re.fullmatch(r"[\u3400-\u9fff·]{2,8}", stripped):
            valid, replacement = True, "[姓名已脱敏]"
        elif kind == "address" and 4 <= len(stripped) <= 100:
            valid, replacement = True, "[地址已脱敏]"
        elif kind == "organization" and 2 <= len(stripped) <= 100:
            valid, replacement = True, "[单位已脱敏]"
        elif kind == "mobile" and MOBILE.fullmatch(stripped):
            valid, replacement = True, "[手机号已脱敏]"
        elif kind == "id_card" and valid_chinese_id(stripped):
            valid, replacement = True, "[身份证号已脱敏]"
        elif kind == "bank_card" and valid_luhn(stripped):
            valid, replacement = True, "[银行卡号已脱敏]"
        elif kind == "email" and EMAIL.fullmatch(stripped):
            valid, replacement = True, "[邮箱已脱敏]"
        elif kind == "credit_code" and valid_credit_code(stripped):
            valid, replacement = True, "[信用代码已脱敏]"
        elif kind == "passport" and PASSPORT.fullmatch(stripped):
            valid, replacement = True, "[护照号已脱敏]"
        elif kind == "plate" and PLATE.fullmatch(stripped):
            valid, replacement = True, "[车牌号已脱敏]"
        if valid:
            hits.append(RedactionHit(span[0], span[1], kind, replacement, confidence, f"相邻字段“{context}”"))

    # Prefer wider/high-confidence hits when recognizers overlap.
    selected: list[RedactionHit] = []
    for hit in sorted(hits, key=lambda item: (-item.confidence, -(item.end - item.start), item.start)):
        if not any(hit.start < item.end and item.start < hit.end for item in selected):
            selected.append(hit)
    return sorted(selected, key=lambda item: item.start)


def redact_text(text: str, context: str | None = None) -> tuple[str, int]:
    hits = detect_pii(text, context)
    result = text
    for hit in reversed(hits):
        result = result[: hit.start] + hit.replacement + result[hit.end :]
    return result, len(hits)


def redaction_replacements(text: str, context: str | None = None) -> list[tuple[str, str]]:
    return [(text[hit.start : hit.end], hit.replacement) for hit in detect_pii(text, context)]
