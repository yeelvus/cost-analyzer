from __future__ import annotations

"""
规则：物品名称包含"钢筋网片" → 材料归类 改为 "钢筋"（不限标准单位）
"""

TARGET_NAME_KEYWORD = "钢筋网片"
TARGET_CATEGORY = "钢筋"


def get_header_map(worksheet):
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet):
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
    unit_col = header_map.get("标准单位")
    category_col = header_map["材料归类"]

    matched_rows = 0
    updated_rows = 0

    for row_idx in range(2, worksheet.max_row + 1):
        item_name = worksheet.cell(row_idx, name_col).value
        if item_name is None:
            continue
        item_name_str = str(item_name).strip()

        if TARGET_NAME_KEYWORD not in item_name_str:
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
