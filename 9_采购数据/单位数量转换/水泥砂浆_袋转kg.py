"""
水泥 / 砂浆：袋 → kg

- 物品名称含：水泥、砂浆（及常见商品名如砌筑砂浆、抹灰砂浆等）
- 单位为「袋」
- 按每袋 50 kg 折算：计算数量(kg) = 袋数 × 50

排除：明显非袋装材料且名称虽沾边的（如水泥砖、水泥空心砖 —— 走砖转换）
"""
from __future__ import annotations

import math
import re
from typing import Any

KG_PER_BAG = 50.0

# 袋装胶凝/砂浆类
NAME_INCLUDE = (
    "水泥", "砂浆", "沙浆", "砂桨", "干混",
)
# 砌块/砖类：即使名称带「水泥」也不按袋转（走砖转换）
BLOCK_EXCLUDE = ("砖", "砌块")

BAG_UNITS = {"袋", "包"}  # 「包」按袋同口径 50kg


def is_cement_or_mortar(name: Any) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    # 水泥空心砖、加气砌块等不走本规则
    if any(k in s for k in BLOCK_EXCLUDE):
        return False
    return any(k in s for k in NAME_INCLUDE)


def is_bag_unit(unit: Any) -> bool:
    if unit is None:
        return False
    return str(unit).strip() in BAG_UNITS


def try_convert(
    *,
    item_name: Any,
    spec: Any = None,
    unit: Any,
    qty: Any,
) -> dict | None:
    if not is_cement_or_mortar(item_name):
        return None
    if not is_bag_unit(unit):
        return None

    try:
        q = float(str(qty).strip().replace(",", "")) if qty is not None and str(qty).strip() != "" else None
    except ValueError:
        return None
    if q is None or not math.isfinite(q) or q < 0:
        return None

    return {
        "转换单位": "kg",
        "计算数量": round(q * KG_PER_BAG, 3),
        "每袋kg": KG_PER_BAG,
        "规则": "水泥砂浆_袋转kg",
    }
