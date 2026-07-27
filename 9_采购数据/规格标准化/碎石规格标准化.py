#!/usr/bin/env python3
"""
碎石规格标准化 + 材料归类。

物品名称含：碎石、石子、级配碎石、卵石、瓜子片、石屑 等
排除：尾款、运费、采购合同字样过多时可仍归碎石

规格统一：
  A级 / A级碎石 / 纯碎石 → A级
  B级 → B级
  3/4寸 / 3/4 / 3*4 → 3/4寸
  空规格保持空（不强制）

材料归类 → 碎石
"""
from __future__ import annotations

import re

NAME_INCLUDE = ("碎石", "石子", "卵石", "瓜子片", "石屑", "砂石", "级配碎石", "毛石")
NAME_EXCLUDE = ("运费", "刀", "机")

TARGET_CATEGORY = "碎石"

# 粒径
SIZE_PATTERNS = [
    (re.compile(r"3\s*[/\*xX×]\s*4\s*寸?", re.I), "3/4寸"),
    (re.compile(r"1\s*[/\*xX×]\s*2\s*寸?", re.I), "1/2寸"),
    (re.compile(r"5\s*[~\-－]\s*10\s*(mm)?", re.I), "5-10mm"),
    (re.compile(r"10\s*[~\-－]\s*20\s*(mm)?", re.I), "10-20mm"),
    (re.compile(r"16\s*[~\-－]\s*31\.?5\s*(mm)?", re.I), "16-31.5mm"),
    (re.compile(r"20\s*[~\-－]\s*40\s*(mm)?", re.I), "20-40mm"),
]


def is_gravel_item(name: object) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if any(k in s for k in NAME_EXCLUDE) and not any(k in s for k in NAME_INCLUDE):
        return False
    return any(k in s for k in NAME_INCLUDE)


def standardize_gravel_spec(spec_value: object) -> str | None:
    if spec_value is None:
        return None
    text = str(spec_value).strip()
    if not text or text.lower() == "nan":
        return None

    t = re.sub(r"\s+", "", text)
    tu = t.upper()

    # 等级
    if "A级" in t or tu in {"A", "A级碎石"} or "A级碎石" in t:
        return "A级"
    if "B级" in t or tu in {"B", "B级碎石"} or "B级碎石" in t:
        return "B级"
    if "C级" in t:
        return "C级"
    if "纯碎石" in t or t == "纯":
        return "A级"  # 现场「纯碎石」并入 A 级口径

    for pat, std in SIZE_PATTERNS:
        if pat.search(t) or pat.search(text):
            return std

    # 已是标准
    if t in {"A级", "B级", "C级", "3/4寸", "1/2寸"}:
        return t

    return None  # 无法识别则不改写


def get_header_map(worksheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet) -> dict:
    header_map = get_header_map(worksheet)
    if "物品名称" not in header_map:
        return {
            "sheet_name": worksheet.title,
            "matched_rows": 0,
            "updated_rows": 0,
            "prefixed_rows": 0,
            "unresolved_rows": [],
            "skipped": True,
        }

    name_col = header_map["物品名称"]
    spec_col = header_map.get("规格")
    cat_col = header_map.get("材料归类")

    matched = updated = cat_upd = 0
    unresolved = []

    for r in range(2, worksheet.max_row + 1):
        name_v = worksheet.cell(r, name_col).value
        if not is_gravel_item(name_v):
            continue
        matched += 1
        name = str(name_v).strip()

        if cat_col is not None:
            cell = worksheet.cell(r, cat_col)
            old = "" if cell.value is None else str(cell.value).strip()
            if old != TARGET_CATEGORY:
                cell.value = TARGET_CATEGORY
                cat_upd += 1

        if spec_col is None:
            continue
        spec_cell = worksheet.cell(r, spec_col)
        old_spec = "" if spec_cell.value is None else str(spec_cell.value).strip()
        if not old_spec:
            continue
        std = standardize_gravel_spec(spec_cell.value)
        if std is None:
            unresolved.append((r, name, old_spec, "gravel"))
            continue
        if old_spec != std:
            spec_cell.value = std
            updated += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched,
        "updated_rows": updated + cat_upd,
        "prefixed_rows": cat_upd,
        "unresolved_rows": unresolved,
        "skipped": False,
    }
