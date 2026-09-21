"""
模板 / 木方：张、根、块等 → m²

物品：模板、模版、木模板、木方、方木、杂木方等（排除机具/五金）
单位：张、根、块、pcs、片 等计件单位（已是 m2 则透传）

单件面积（主工作面）：
  取规格中最大的两个边长（mm）换算为米后相乘
  - 板材 2440x1220x13mm → 2.440 × 1.220 = 2.9768 m²/张
  - 木方 40x80x3000mm   → 0.080 × 3.000 = 0.24 m²/根（较大侧面）

计算数量(m²) = 件数 × 单件面积
"""
from __future__ import annotations

import math
import re
from typing import Any

NAME_INCLUDE = (
    "模板", "模版", "木模板", "红模板", "小模板",
    "木方", "方木", "杂木方", "木方条",
)
NAME_EXCLUDE = ("机", "开孔", "轮", "尾款", "五金", "卡子", "卡扣", "夹具", "螺丝")

COUNT_UNITS = {"张", "根", "块", "片", "pcs", "pc", "PCS", "PC", "个"}
AREA_UNITS = {"m2", "m²", "㎡", "平方米", "平米", "平", "M2"}

NUM = r"(\d+(?:\.\d+)?)"
DIM_RE = re.compile(
    rf"{NUM}\s*(?:mm)?\s*[xX×*＊]\s*{NUM}\s*(?:mm)?(?:\s*[xX×*＊]\s*{NUM}\s*(?:mm)?)?",
    re.IGNORECASE,
)


def is_template_or_wood(name: Any) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if any(k in s for k in NAME_EXCLUDE):
        return False
    if any(k in s for k in NAME_INCLUDE):
        return True
    if "木" in s and "方" in s:
        return True
    return False


def _norm_unit(unit: Any) -> str:
    if unit is None:
        return ""
    return str(unit).strip()


def is_count_unit(unit: Any) -> bool:
    u = _norm_unit(unit)
    return u in COUNT_UNITS or u.lower() in {"pcs", "pc"}


def is_area_unit(unit: Any) -> bool:
    u = _norm_unit(unit)
    if not u:
        return False
    if u in AREA_UNITS:
        return True
    ul = u.lower()
    return ul in {"m2", "m²", "㎡"} or "平方" in u


def parse_face_area_m2(spec: Any) -> float | None:
    """从规格解析主面面积 m²（两最大边）。"""
    if spec is None:
        return None
    text = str(spec).strip()
    if not text or text.lower() == "nan":
        return None
    text = re.sub(r"\s+", "", text)
    text = text.replace("×", "x").replace("＊", "x").replace("*", "x")
    text = text.replace("毫米", "mm").replace("mm", "")

    m = DIM_RE.search(text)
    if not m:
        return None
    dims = [float(m.group(1)), float(m.group(2))]
    if m.group(3) is not None:
        dims.append(float(m.group(3)))

    # 已是米制小数的情况（极少）：若 max < 20 当作 m
    if max(dims) < 20:
        dims_mm = [d * 1000 for d in dims]
    else:
        dims_mm = dims

    if any(d <= 0 or d > 50000 for d in dims_mm):
        return None

    dims_mm = sorted(dims_mm, reverse=True)
    # 主面 = 最大两边
    a_m = dims_mm[0] / 1000.0
    b_m = dims_mm[1] / 1000.0
    area = a_m * b_m
    if area <= 0 or not math.isfinite(area):
        return None
    return area


def try_convert(
    *,
    item_name: Any,
    spec: Any,
    unit: Any,
    qty: Any,
) -> dict | None:
    if not is_template_or_wood(item_name):
        return None

    try:
        q = float(str(qty).strip().replace(",", "")) if qty is not None and str(qty).strip() != "" else None
    except ValueError:
        return None
    if q is None or not math.isfinite(q) or q < 0:
        return None

    # 已是面积单位：透传
    if is_area_unit(unit):
        return {
            "转换单位": "m2",
            "计算数量": round(q, 3),
            "规则": "模板木方_m2透传",
        }

    if not is_count_unit(unit):
        return None

    area = parse_face_area_m2(spec)
    if area is None:
        return None

    return {
        "转换单位": "m2",
        "计算数量": round(q * area, 3),
        "单件面积m2": round(area, 6),
        "规则": "模板木方_转m2",
    }
