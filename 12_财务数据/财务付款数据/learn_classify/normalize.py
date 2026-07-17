# -*- coding: utf-8 -*-
"""文本规范化：用于记忆键、关键词匹配。"""

from __future__ import annotations

import re
import unicodedata
from typing import Any


def normalize_text(text: Any) -> str:
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).strip().lower()
    s = re.sub(r"[\s\t\r\n]+", "", s)
    s = re.sub(r"[()（）\[\]【】:：_\-·•,/，、]", "", s)
    return s


def memory_key(subject: Any, summary: Any, vendor: Any = None) -> str:
    """精确记忆主键：科目|摘要（可选供应商）。"""
    parts = [normalize_text(subject), normalize_text(summary)]
    if vendor is not None and str(vendor).strip():
        parts.append(normalize_text(vendor))
    return "|".join(parts)


def soft_key(summary: Any) -> str:
    """仅摘要键，用于弱匹配兜底。"""
    return normalize_text(summary)
