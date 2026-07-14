#!/usr/bin/env python3
"""
模板 / 木方 规格标准化 + 材料归类。

物品：
  - 模板类：模板、模版、木模板、红模板、小模板… → 材料归类「模板」
  - 木方类：木方、杂木方、木方条、方木… → 材料归类亦为「模板」（与模板同归）
排除：机具、开孔器、轮、尾款、五金 等

规格统一为毫米（长边优先排序截面时：大面在前）：
  40*80*3000 / 40*80*3m / 40*80*3000mm → 40x80x3000mm
  1.5寸*3寸*3m → 按 1寸=25.4mm → 38x76x3000mm
  2440*1220*13mm / 244*122*1.3cm / 1.22*2.44 → 2440x1220x13mm 或 2440x1220mm
"""
from __future__ import annotations

import math
import re

TEMPLATE_NAME = ("模板", "模版", "木模板", "红模板", "小模板")
WOOD_NAME = ("木方", "方木", "杂木方", "木方条", "木条", "那样木", "样那木", "混合木")
# 名称含这些则跳过
NAME_EXCLUDE = ("机", "开孔", "轮", "尾款", "五金", "卡子", "卡扣", "夹具", "螺丝")

TARGET_CAT_TEMPLATE = "模板"
TARGET_CAT_WOOD = "模板"  # 木方与模板统一归类为「模板」

NUM = r"(\d+(?:\.\d+)?)"
UNIT = r"(mm|cm|m|毫米|厘米|米|寸)?"
# 2 段或 3 段
DIM_RE = re.compile(
    rf"{NUM}\s*{UNIT}\s*[xX×*＊]\s*{NUM}\s*{UNIT}(?:\s*[xX×*＊]\s*{NUM}\s*{UNIT})?",
    re.IGNORECASE,
)


def format_mm(v: float) -> str:
    r = round(v)
    if math.isclose(v, r, abs_tol=0.51):
        return str(int(r))
    return f"{v:.1f}".rstrip("0").rstrip(".")


def _norm_u(u: str | None) -> str | None:
    if not u:
        return None
    t = str(u).strip().lower()
    t = t.replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")
    if t in {"mm", "cm", "m", "寸"}:
        return t
    return None


def _to_mm(v: float, u: str | None, *, default: str) -> float:
    unit = u or default
    if unit == "mm":
        return v
    if unit == "cm":
        return v * 10
    if unit == "m":
        return v * 1000
    if unit == "寸":
        return v * 25.4  # 现场木材常用近似英寸
    return v


def _infer_unit_for_dim(values: list[float], units: list[str | None], idx: int) -> str:
    """
    逐维推断单位，避免「40x80x3m」把 40/80 也当成 m。
    """
    if units[idx]:
        return units[idx]

    n = len(values)
    v = values[idx]
    present = {u for u in units if u}

    # 同行出现 cm，且数值像厘米板材（50~400）→ cm
    if "cm" in present and 50 <= v <= 400:
        return "cm"
    # 1.3cm 厚度旁的 244/122
    if "cm" in present and v > 30 and max(values) < 500:
        return "cm"

    # 三维木方：前两维截面 mm，最后一维若 <=20 且前两维都 <200 → 米
    if n == 3 and idx == 2 and v <= 20 and values[0] < 200 and values[1] < 200:
        # 若最后一维已写 mm 不会进这里；无单位的 3 → m
        if "m" in present or v <= 12:
            return "m"
        return "mm"
    if n == 3 and idx < 2 and values[0] < 200 and values[1] < 200 and max(values) < 500:
        if v < 15:
            return "寸" if max(values[0], values[1]) <= 15 else "mm"
        return "mm"

    # 二维板：两边都 <10 → 米（1.22x2.44）
    if n == 2 and max(values) < 10:
        return "m"

    if v >= 100:
        return "mm"
    if v <= 30:
        return "mm"
    return "mm"


def standardize_template_wood_spec(spec_value: object) -> str | None:
    if spec_value is None:
        return None
    text = str(spec_value).strip()
    if not text or text.lower() == "nan":
        return None
    # 非尺寸描述
    if re.fullmatch(r"[\u4e00-\u9fff]+", text):
        return None  # 如「大板」

    text = re.sub(r"\s+", "", text)
    text = text.replace("×", "x").replace("＊", "x").replace("*", "x")
    text = text.replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")

    m = DIM_RE.search(text)
    if not m:
        return None

    g = m.groups()
    vals = [float(g[0]), float(g[2])]
    units: list[str | None] = [_norm_u(g[1]), _norm_u(g[3])]
    if g[4] is not None:
        vals.append(float(g[4]))
        units.append(_norm_u(g[5]))

    dims = [
        _to_mm(vals[i], units[i], default=_infer_unit_for_dim(vals, units, i))
        for i in range(len(vals))
    ]

    # 合理性
    if any(d <= 0 or d > 20000 for d in dims):
        return None

    # 板材二维：长边在前
    if len(dims) == 2:
        a, b = sorted(dims, reverse=True)
        return f"{format_mm(a)}x{format_mm(b)}mm"

    # 三维木方：两截面 <300 且长度 >=500 → 小x中x长
    sorted_all = sorted(dims)
    if sorted_all[2] >= 500 and sorted_all[0] < 300 and sorted_all[1] < 300:
        return f"{format_mm(sorted_all[0])}x{format_mm(sorted_all[1])}x{format_mm(sorted_all[2])}mm"

    # 模板板：长 x 宽 x 厚（厚通常最小）
    thick = min(dims)
    planes = sorted([d for d in dims if d != thick] or dims[:2], reverse=True)
    if len(planes) == 1:
        planes = [planes[0], planes[0]]
    if thick < 100 and planes[0] >= 500:
        return f"{format_mm(planes[0])}x{format_mm(planes[1])}x{format_mm(thick)}mm"

    a, b, c = sorted(dims, reverse=True)
    return f"{format_mm(a)}x{format_mm(b)}x{format_mm(c)}mm"


def detect_kind(item_name: object) -> str | None:
    """返回 template / wood / None。"""
    if item_name is None:
        return None
    name = str(item_name).strip()
    if not name:
        return None
    if any(k in name for k in NAME_EXCLUDE):
        return None
    # 木方优先（「木方模板」少见）
    if any(k in name for k in WOOD_NAME) or ("木" in name and "方" in name):
        if any(k in name for k in TEMPLATE_NAME):
            # 木模板 → 模板
            return "template"
        return "wood"
    if any(k in name for k in TEMPLATE_NAME):
        return "template"
    return None


def get_header_map(worksheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet) -> dict:
    header_map = get_header_map(worksheet)
    if "物品名称" not in header_map or "规格" not in header_map:
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
    cat_col = header_map.get("材料归类")

    matched = updated = cat_upd = 0
    unresolved = []

    for r in range(2, worksheet.max_row + 1):
        kind = detect_kind(worksheet.cell(r, name_col).value)
        if kind is None:
            continue
        matched += 1
        name = str(worksheet.cell(r, name_col).value).strip()
        target_cat = TARGET_CAT_WOOD if kind == "wood" else TARGET_CAT_TEMPLATE

        if cat_col is not None:
            cell = worksheet.cell(r, cat_col)
            old = "" if cell.value is None else str(cell.value).strip()
            if old != target_cat:
                cell.value = target_cat
                cat_upd += 1

        spec_cell = worksheet.cell(r, spec_col)
        old_spec = "" if spec_cell.value is None else str(spec_cell.value).strip()
        std = standardize_template_wood_spec(spec_cell.value)
        if std is None:
            if old_spec:
                unresolved.append((r, name, old_spec, kind))
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
