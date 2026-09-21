from __future__ import annotations

"""
规则：物品名称为「钢筋」或「螺纹钢」，且单位为 kg（含 公斤/千克）
     → 材料归类 填入「钢筋」。
"""

import re

TARGET_ITEM_NAMES = {"钢筋", "螺纹钢"}
TARGET_CATEGORY = "钢筋"
# 单位写法：kg / KG / 公斤 / 千克
KG_UNIT_RE = re.compile(r"^(kg|kgs|公斤|千克)$", re.IGNORECASE)


def get_header_map(worksheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def _is_kg_unit(value: object) -> bool:
    if value is None:
        return False
    text = str(value).strip().replace(" ", "")
    if not text:
        return False
    return KG_UNIT_RE.match(text) is not None


def process_worksheet(worksheet) -> dict:
    header_map = get_header_map(worksheet)

    required = {"物品名称", "材料归类"}
    if not required.issubset(header_map):
        return {
            "sheet_name": worksheet.title,
            "matched_rows": 0,
            "updated_rows": 0,
            "prefixed_rows": 0,
            "unresolved_rows": [],
            "skipped": True,
        }

    name_col = header_map["物品名称"]
    category_col = header_map["材料归类"]
    # 兼容「单位」与「标准单位」
    unit_col = header_map.get("单位") or header_map.get("标准单位")

    matched_rows = 0
    updated_rows = 0

    for row_idx in range(2, worksheet.max_row + 1):
        item_name = worksheet.cell(row_idx, name_col).value
        if item_name is None:
            continue
        item_name_str = str(item_name).strip()
        if item_name_str not in TARGET_ITEM_NAMES:
            continue

        if unit_col is not None:
            unit_val = worksheet.cell(row_idx, unit_col).value
            if not _is_kg_unit(unit_val):
                continue
        # 无单位列时：不按 kg 条件过滤（避免误伤）；严格按「有单位且为 kg」时才改
        else:
            continue

        matched_rows += 1
        category_cell = worksheet.cell(row_idx, category_col)
        old_val = "" if category_cell.value is None else str(category_cell.value).strip()
        if old_val != TARGET_CATEGORY:
            category_cell.value = TARGET_CATEGORY
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows,
        "prefixed_rows": 0,
        "unresolved_rows": [],
        "skipped": False,
    }
