from __future__ import annotations

import re


TARGET_NAME_KEYWORD = "砂浆"
TARGET_UNIT = "袋"
TARGET_CATEGORY = "砂浆"

# 只匹配纯数字等级（整数或小数），拒绝管径类后缀如 90MM、450MM
SPEC_PATTERN = re.compile(r"^\s*([mM])\s*(\d+(?:\.\d+)?)\s*$")

# 匹配 "197水泥"、"水泥197" 等录入形式
CEMENT_197_PATTERN = re.compile(r"197\s*水泥|水泥\s*197")
# 砂浆/水泥行可接受的标准单位
TARGET_UNITS = {"袋", "包"}
# 水泥规格等级识别（197 / m197 / M197）
CEMENT_GRADE_SPEC_PATTERN = re.compile(r"^\s*[mM]?\s*(\d+)\s*$")


def get_header_map(worksheet):
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def normalize_spec(spec_value: object) -> str | None:
    if spec_value is None:
        return None

    text = str(spec_value).strip()
    if not text or text.lower() == "nan":
        return None

    # 剥离尾部 MM/mm 单位后缀（录入错误，如 M25MM → M25）
    text = re.sub(r"(?i)mm$", "", text).strip()

    m = SPEC_PATTERN.match(text)
    if not m:
        return None

    return f"M{m.group(2)}"


def normalize_unit(unit_value: object) -> str:
    if unit_value is None:
        return ""
    return str(unit_value).strip()


def get_cement_grade(spec_value: object) -> str | None:
    """将 197 / m197 / M197 统一为 M197 格式；其他规格返回 None。"""
    if spec_value is None:
        return None
    text = re.sub(r"(?i)mm$", "", str(spec_value).strip()).strip()
    m = CEMENT_GRADE_SPEC_PATTERN.match(text)
    if m:
        return f"M{m.group(1)}"
    return None


def process_worksheet(worksheet):
    header_map = get_header_map(worksheet)

    required = {"物品名称", "规格", "材料归类"}
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
    spec_col = header_map["规格"]
    category_col = header_map["材料归类"]
    unit_col = header_map.get("标准单位") or header_map.get("单位")

    matched_rows = 0
    updated_rows = 0
    unresolved_rows = []

    for row_idx in range(2, worksheet.max_row + 1):
        name_cell = worksheet.cell(row_idx, name_col)
        spec_cell = worksheet.cell(row_idx, spec_col)
        category_cell = worksheet.cell(row_idx, category_col)

        item_name = "" if name_cell.value is None else str(name_cell.value).strip()

        is_cement_197 = bool(CEMENT_197_PATTERN.search(item_name))
        is_mortar = TARGET_NAME_KEYWORD in item_name

        # 名称精确为"水泥"且规格形如 197 / m197 / M197
        is_plain_cement = False
        plain_cement_grade = None
        if item_name == "水泥" and not is_cement_197:
            plain_cement_grade = get_cement_grade(spec_cell.value)
            if plain_cement_grade is not None:
                is_plain_cement = True

        if not (is_mortar or is_cement_197 or is_plain_cement):
            continue

        unit_value = ""
        if unit_col is not None:
            unit_value = normalize_unit(worksheet.cell(row_idx, unit_col).value)
        # 砂浆行要求单位为袋/包；水泥类行不限制单位
        if is_mortar and not is_cement_197 and not is_plain_cement:
            if unit_value not in TARGET_UNITS:
                continue

        matched_rows += 1

        row_changed = False

        # 水泥197：名称统一为"水泥"，规格强制写入 M197
        if is_cement_197:
            if item_name != "水泥":
                name_cell.value = "水泥"
                row_changed = True
            normalized_spec = "M197"
        elif is_plain_cement:
            # 名称已是"水泥"，只规格标准化（197/m197 → M197）
            normalized_spec = plain_cement_grade
        else:
            normalized_spec = normalize_spec(spec_cell.value)

        if not is_cement_197 and not is_plain_cement and normalized_spec is None:
            unresolved_rows.append(
                (
                    row_idx,
                    item_name,
                    "" if spec_cell.value is None else str(spec_cell.value).strip(),
                )
            )
            continue

        old_spec = "" if spec_cell.value is None else str(spec_cell.value).strip()
        if old_spec != normalized_spec:
            spec_cell.value = normalized_spec
            row_changed = True

        # 砂浆行根据规格推断物品名称（水泥行名称不做推断）
        if not is_cement_197 and not is_plain_cement:
            if normalized_spec.startswith("M3"):
                target_name = "砌筑砂浆"
            elif normalized_spec.startswith("M2"):
                target_name = "抹灰砂浆"
            elif normalized_spec.startswith("M1"):
                target_name = "水泥"
            else:
                target_name = None

            if target_name is not None and item_name != target_name:
                name_cell.value = target_name
                row_changed = True

        # 水泥行不修改材料归类
        if not is_plain_cement:
            old_category = "" if category_cell.value is None else str(category_cell.value).strip()
            if old_category != TARGET_CATEGORY:
                category_cell.value = TARGET_CATEGORY
                row_changed = True

        # 标准单位不在袋/包时统一写入"袋"（水泥行不强制修改单位）
        if not is_plain_cement and unit_col is not None and unit_value not in TARGET_UNITS:
            worksheet.cell(row_idx, unit_col).value = "袋"
            row_changed = True

        if row_changed:
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows,
        "prefixed_rows": 0,
        "unresolved_rows": unresolved_rows,
        "skipped": False,
    }
