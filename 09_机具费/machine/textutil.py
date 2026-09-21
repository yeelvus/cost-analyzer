# -*- coding: utf-8 -*-
from __future__ import annotations
import re
import unicodedata

MONTH_LIKE_TEXT_PATTERNS = [
    re.compile(r"^\d{4}[-/.年]\s*\d{1,2}(?:月|月份)?$"),
    re.compile(r"^\d{4}(?:0[1-9]|1[0-2])$"),
    re.compile(r"^(?:20\d{2})\s*(?:年)?\s*(?:0?[1-9]|1[0-2])\s*(?:月|月份)$"),
    re.compile(r"^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)$", re.I),
]

def is_month_like_text(value):
    text = str(value).strip() if value is not None else ""
    if not text or text == "nan":
        return False
    for pattern in MONTH_LIKE_TEXT_PATTERNS:
        if re.match(pattern, text):
            return True
    return False


def normalize_lookup_text(value):
    """将机具种类/型号等映射键归一化，减少空格、全半角和大小写差异。"""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip()
    if not text or text.lower() in {"nan", "none", "nat", "<na>"}:
        return ""
    text = re.sub(r"\s+", "", text)
    text = text.replace("—", "-").replace("–", "-").replace("－", "-")
    return text.lower()


def normalize_column_name(col_name):
    """标准化列名：去除空格、换行符、全角符号等"""
    if not col_name:
        return ""
    col_name = str(col_name)
    # 去除所有空白字符（空格、换行、制表符等）
    col_name = re.sub(r'\s+', '', col_name)
    # 替换全角符号为半角
    replacements = {
        '（': '(', '）': ')', '：': ':', '％': '%',
        '，': ',', '＋': '+', '／': '/'
    }
    for old, new in replacements.items():
        col_name = col_name.replace(old, new)
    return col_name


def safe_to_float(value):
    """安全地将值转换为浮点数"""
    if value is None or value == "" or str(value).strip() == "":
        return 0.0
    try:
        # 去除逗号和空格
        val_str = str(value).replace(",", "").replace(" ", "").strip()
        if val_str == "" or val_str == "nan":
            return 0.0
        return float(val_str)
    except:
        return 0.0


def normalize_year_month_text(value):
    """将常见年月文本归一为 YYYY-MM。"""
    text = str(value).strip() if value is not None else ""
    if not text or text.lower() in {"nan", "none", "nat", "<na>"}:
        return ""

    patterns = [
        re.compile(r"(?<!\d)(20\d{2})[-/.年]\s*(1[0-2]|0?[1-9])(?:月|月份)?"),
        re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)"),
    ]

    for pattern in patterns:
        match = pattern.search(text)
        if not match:
            continue
        year = int(match.group(1))
        month = int(match.group(2))
        if 2020 <= year <= 2035:
            return f"{year:04d}-{month:02d}"

    return ""
