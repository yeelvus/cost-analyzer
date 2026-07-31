# -*- coding: utf-8 -*-
"""文本规范化：用于记忆键、关键词匹配、相似泛化。"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, FrozenSet, Set

# 过宽核心词：单独出现时不写入 core 记忆，避免误伤
_CORE_STOP = frozenset(
    {
        "费用",
        "付款",
        "支付",
        "报销",
        "工程",
        "项目",
        "合同",
        "其他",
        "暂估",
        "预付",
        "结算",
        "结算款",
        "货款",
        "材料",
        "劳务",
        "服务",
        "管理",
        "办公",
        "差旅",
        "交通",
        "采购",
        "购买",
        "支出",
        "收入",
        "转账",
        "汇款",
        "现金",
        "支票",
        "发票",
        "税",
        "税金",
        "税款",
        "月",
        "月份",
        "年",
        "日",
        "号",
        "等",
        "及",
        "和",
        "与",
        "the",
        "and",
        "for",
        "fee",
        "pay",
        "payment",
    }
)

# 相似匹配时降权的通用片段
_SIM_STOP = _CORE_STOP | frozenset(
    {
        "公司",
        "有限",
        "有限公司",
        "ltd",
        "co",
        "thailand",
        "thai",
        "股份",
        "集团",
        "中心",
        "单位",
        "部门",
    }
)


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


def core_summary(summary: Any) -> str:
    """
    去日期/数字/标点后的摘要核心，用于跨月份/数量变体的泛化键。
    例：'01月保安费用' / '02月保安费用' → '保安费用'
    注意：不剥非日期语境中的「月」字（如「测试月」「满月」）。
    """
    n = normalize_text(summary)
    if not n:
        return ""
    # 先剥明确日期模式，再剥残留数字
    n = re.sub(r"20\d{2}年", "", n)
    n = re.sub(r"\d{1,2}月份?", "", n)
    n = re.sub(r"\d{1,2}日", "", n)
    n = re.sub(r"\d{1,2}号", "", n)
    n = re.sub(r"\d+(\.\d+)?", "", n)
    # 尾部无意义「等」
    n = re.sub(r"等+$", "", n)
    n = n.strip()
    if len(n) < 2 or n in _CORE_STOP:
        return ""
    return n


def core_memory_key(summary: Any) -> str:
    """core 记忆键：|core:保安费用"""
    c = core_summary(summary)
    if not c:
        return ""
    return f"|core:{c}"


def summary_bigrams(text: str) -> FrozenSet[str]:
    """字符 bigram 集合（已 normalize 的文本）。"""
    if not text or len(text) < 2:
        return frozenset()
    return frozenset(text[i : i + 2] for i in range(len(text) - 1))


def summary_features(summary: Any) -> Set[str]:
    """
    用于相似匹配的特征：core、中文 2~4 字片段、英文词。
    过滤过宽 stop。
    """
    n = normalize_text(summary)
    core = core_summary(summary)
    feats: Set[str] = set()
    if core:
        feats.add(core)
    # 去数字后的串上抽 ngram
    pure = core_summary(summary) or re.sub(r"\d+(\.\d+)?", "", n)
    for length in (4, 3, 2):
        if len(pure) < length:
            continue
        for i in range(len(pure) - length + 1):
            tok = pure[i : i + length]
            if tok in _SIM_STOP:
                continue
            if re.fullmatch(r"[a-z]+", tok) and len(tok) < 3:
                continue
            feats.add(tok)
    for m in re.findall(r"[a-z]{3,12}", n):
        if m not in _SIM_STOP:
            feats.add(m)
    return feats


def jaccard(a: Set[str] | FrozenSet[str], b: Set[str] | FrozenSet[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    if inter == 0:
        return 0.0
    return inter / len(a | b)
