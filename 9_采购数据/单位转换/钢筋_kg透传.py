"""
钢筋（含螺纹钢、圆钢等棒材）：单位已是 kg 时，转换列直接填入。

- 物品名称含：钢筋、螺纹钢、圆钢（排除网片、运费、套筒、设备等）
- 单位为 kg / 公斤 / 千克 等
- 转换单位 = kg，计算数量 = 原数量（不改值，仅写入对比列）

说明：钢筋网片的 m2→kg 由「钢筋网片_m2转kg」优先处理；
      本规则只处理「本来就是 kg」的棒材，不重复换算。
"""
from __future__ import annotations

import math
import re
from typing import Any

NAME_INCLUDE = ("钢筋", "螺纹钢", "圆钢")
NAME_EXCLUDE = (
    "网片", "运费", "套筒", "弯曲机", "切断机", "调直", "加工棚", "工棚",
    "检查井", "方桩", "管桩", "扎钩", "绑丝", "铁丝",
)

KG_UNITS = {
    "kg", "KG", "Kg", "公斤", "千克", "kgs", "KGS",
}


def is_rebar_item(name: Any) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if any(k in s for k in NAME_EXCLUDE):
        return False
    return any(k in s for k in NAME_INCLUDE)


def is_kg_unit(unit: Any) -> bool:
    if unit is None:
        return False
    u = str(unit).strip()
    if u in KG_UNITS:
        return True
    return u.lower() in {"kg", "kgs"}


def try_convert(
    *,
    item_name: Any,
    spec: Any = None,
    unit: Any,
    qty: Any,
) -> dict | None:
    if not is_rebar_item(item_name):
        return None
    if not is_kg_unit(unit):
        return None

    try:
        q = float(str(qty).strip().replace(",", "")) if qty is not None and str(qty).strip() != "" else None
    except ValueError:
        return None
    if q is None or not math.isfinite(q) or q < 0:
        return None

    return {
        "转换单位": "kg",
        "计算数量": round(q, 3),
        "规则": "钢筋_kg透传",
    }
