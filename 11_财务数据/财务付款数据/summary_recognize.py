# -*- coding: utf-8 -*-
"""摘要识别（与科目分类分开）。

输入原始摘要，以及可选的已匹配项目名称。
只产出三样东西：

- 业务摘要：去掉项目名前缀后的付款描述，分类和记忆都用它
- 项目提示：摘要里带出来的项目名片段（不写回「项目名称」列）
- 记忆键：业务摘要的规范化键，与 learn_classify.normalize.memory_key 一致

不查科目树，不判一级/二级/三级。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import pandas as pd

from learn_classify.normalize import memory_key

_DASH_CHARS = ("－", "—", "–", "﹣", "−")
_DASH_PREFIX_RE = re.compile(r"^[\s\-－—–﹣−]+")
_DATE_PREFIX_RE = re.compile(r"^\d{1,2}([./月]\d{1,2})")
_NOISE_PREFIX_RE = re.compile(
    r"^(内部|银行|订正|转账|未知|资金|手续费|利息|税费|工资|社保)"
)
# 前缀里出现这些，视为业务动作，不当作项目名
_ACTION_IN_PREFIX_RE = re.compile(
    r"(支付|采购|发放|报销|付款|转账|收款|工资|费用|取现|保证金|退回|外售|"
    r"客户名|合同款|投标|退押金|存入|调拨|手续费|代付|代垫|代扣|暂估|预付|结算|"
    r"利息|税费|社保)"
)
_CJK_OR_LETTER_RE = re.compile(r"[\u4e00-\u9fffA-Za-z]")


def _cell(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def unify_dashes(text: str) -> str:
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    return text


def strip_project_suffix(name: str) -> str:
    name = name.strip()
    if name.endswith("项目") and len(name) > 2:
        return name[:-2].strip()
    return name


def extract_summary_prefix(summary: Any) -> str:
    """摘要里第一个连接符之前的文字。没有连接符则返回空。"""
    text = unify_dashes(_cell(summary))
    if "-" not in text:
        return ""
    return text.split("-", 1)[0].strip()


def is_plausible_project_name(prefix: str) -> bool:
    """像项目简称，而不是日期、业务动作或长描述。"""
    if not prefix:
        return False
    if len(prefix) < 2 or len(prefix) > 16:
        return False
    if _DATE_PREFIX_RE.match(prefix):
        return False
    if _NOISE_PREFIX_RE.search(prefix):
        return False
    if _ACTION_IN_PREFIX_RE.search(prefix):
        return False
    if not _CJK_OR_LETTER_RE.search(prefix):
        return False
    if prefix.count("(") + prefix.count("（") >= 1 and len(prefix) > 10:
        return False
    digit_ratio = sum(ch.isdigit() for ch in prefix) / max(len(prefix), 1)
    if digit_ratio > 0.4:
        return False
    return True


def _project_name_candidates(project_name: str) -> list[str]:
    name = _cell(project_name)
    if not name:
        return []
    found = {name}
    short = strip_project_suffix(name)
    if short:
        found.add(short)
    if not name.endswith("项目"):
        found.add(name + "项目")
    return sorted(found, key=len, reverse=True)


def _strip_named_prefix(project_name: str, text: str) -> tuple[str, str]:
    """摘要以已知项目名开头时，返回 (前缀, 剩余业务描述)。"""
    for cand in _project_name_candidates(project_name):
        if text[: len(cand)].casefold() != cand.casefold():
            continue
        rest = _DASH_PREFIX_RE.sub("", text[len(cand) :]).strip()
        if rest:
            return text[: len(cand)], rest
    return "", text


@dataclass(frozen=True)
class SummaryRecognition:
    raw_summary: str
    business_summary: str
    project_hint: str
    memory_key: str


def _key_for(business: str) -> str:
    if not business:
        return ""
    return memory_key("", business, None)


def recognize_summary(
    summary: Any,
    project_name: Any = None,
) -> SummaryRecognition:
    """从原始摘要识别业务摘要和项目提示。"""
    raw = _cell(summary)
    text = unify_dashes(raw)
    hint = ""
    business = text

    if text:
        named_hint, named_rest = _strip_named_prefix(_cell(project_name), text)
        if named_hint:
            hint, business = named_hint, named_rest
        elif "-" in text:
            left, right = text.split("-", 1)
            left, right = left.strip(), right.strip()
            if right and is_plausible_project_name(left):
                hint, business = left, right

    if not business:
        business = text
    return SummaryRecognition(
        raw_summary=raw,
        business_summary=business,
        project_hint=hint,
        memory_key=_key_for(business),
    )


def resolve_business(
    summary: Any,
    project_name: Any = None,
    business_summary: Any = None,
) -> SummaryRecognition:
    """已有业务摘要时直接采用，不再剥第二次。"""
    given = _cell(business_summary)
    if not given:
        return recognize_summary(summary, project_name)
    hinted = recognize_summary(summary, project_name)
    return SummaryRecognition(
        raw_summary=hinted.raw_summary or given,
        business_summary=given,
        project_hint=hinted.project_hint,
        memory_key=_key_for(given),
    )


def _demo() -> None:
    cases = [
        ("欣旺达-材料款", "欣旺达项目", "材料款", "欣旺达"),
        ("欣旺达项目材料款", "欣旺达", "材料款", "欣旺达项目"),
        ("支付-材料款", "", "支付-材料款", ""),
        ("01月保安费用", "", "01月保安费用", ""),
        ("材料款", "欣旺达", "材料款", ""),
        ("代付-税款", "", "代付-税款", ""),
        ("", "欣旺达", "", ""),
    ]
    for summary, project, business, hint in cases:
        got = recognize_summary(summary, project)
        if got.business_summary != business or got.project_hint != hint:
            raise SystemExit(
                f"识别不符: {summary!r} / {project!r} → "
                f"{got.business_summary!r} / {got.project_hint!r}，"
                f"期望 {business!r} / {hint!r}"
            )
        if business and not got.memory_key:
            raise SystemExit(f"缺少记忆键: {summary!r}")
    kept = resolve_business("欣旺达-材料款", "欣旺达", "材料款")
    if kept.business_summary != "材料款":
        raise SystemExit("已有业务摘要被再次剥离")
    print("summary_recognize: ok")


if __name__ == "__main__":
    _demo()
