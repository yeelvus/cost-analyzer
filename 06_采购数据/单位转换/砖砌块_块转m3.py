"""
砖 / 加气块 / 砌块：块 → m³

- 物品名称含空心砖、实心砖、加气块、砌块等（排除砂浆、胶泥等）
- 单位为「块」
- 规格优先标准格式 390x190x190mm，兼容 190*190*390、19*19*39cm

单块体积：
  V(m³) = (a/1000) * (b/1000) * (c/1000)   # a,b,c 为 mm

计算数量(m³) = 块数 × V
"""
from __future__ import annotations

import math
import re
from typing import Any

NAME_INCLUDE = (
    "空心砖", "实心砖", "加气块", "加气砖", "水泥砖", "红砖", "砌块",
    "多孔砖", "页岩砖", "标砖", "通风砖", "植草砖", "面包砖", "墙砖",
    "水泥空心砖", "实心水泥砖", "人行道砖", "人行横道砖", "地板砖",
)
NAME_EXCLUDE = (
    "砂浆", "胶泥", "砖刀", "抹灰", "粘结", "粘接", "粉刷", "机", "锯",
)

BLOCK_UNITS = {"块"}

NUM = r"(\d+(?:\.\d+)?)"
UNIT = r"(mm|cm|m|毫米|厘米|米)?"
DIM3_RE = re.compile(
    rf"{NUM}\s*{UNIT}\s*[xX×*＊]\s*{NUM}\s*{UNIT}\s*[xX×*＊]\s*{NUM}\s*{UNIT}",
    re.IGNORECASE,
)


def is_brick_item(name: Any) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if any(k in s for k in NAME_EXCLUDE):
        return False
    if any(k in s for k in NAME_INCLUDE):
        return True
    return "砖" in s and "瓷砖胶" not in s


def is_block_unit(unit: Any) -> bool:
    if unit is None:
        return False
    return str(unit).strip() in BLOCK_UNITS


def _norm_u(u: str | None) -> str | None:
    if not u:
        return None
    t = str(u).strip().lower().replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")
    return t if t in {"mm", "cm", "m"} else None


def _to_mm(v: float, u: str) -> float:
    if u == "mm":
        return v
    if u == "cm":
        return v * 10
    if u == "m":
        return v * 1000
    return v


def _infer_units(values: list[float], units: list[str | None]) -> list[str]:
    if any(units):
        present = [u for u in units if u]
        maj = present[0]
        return [u if u else maj for u in units]

    vmax, vmin = max(values), min(values)
    if vmax >= 100 and vmin <= 50:
        out = []
        for v in values:
            if v >= 100:
                out.append("mm")
            elif v < 50:
                out.append("cm")
            else:
                out.append("mm")
        return out
    if vmax >= 100:
        return ["mm", "mm", "mm"]
    if vmax <= 80:
        return ["cm", "cm", "cm"]
    return ["mm", "mm", "mm"]


def parse_dims_mm(spec: Any) -> tuple[float, float, float] | None:
    if spec is None:
        return None
    text = str(spec).strip()
    if not text or text.lower() == "nan":
        return None
    text = re.sub(r"\s+", "", text)
    text = text.replace("×", "x").replace("＊", "x").replace("*", "x")
    text = text.replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")

    m = DIM3_RE.search(text)
    if not m:
        return None
    v1, u1, v2, u2, v3, u3 = m.groups()
    values = [float(v1), float(v2), float(v3)]
    units = _infer_units(
        values,
        [_norm_u(u1), _norm_u(u2), _norm_u(u3)],
    )
    dims = [_to_mm(values[i], units[i]) for i in range(3)]
    if any(d < 15 or d > 1200 for d in dims):
        return None
    return dims[0], dims[1], dims[2]


def volume_m3(dims_mm: tuple[float, float, float]) -> float:
    a, b, c = dims_mm
    return (a / 1000.0) * (b / 1000.0) * (c / 1000.0)


def try_convert(
    *,
    item_name: Any,
    spec: Any,
    unit: Any,
    qty: Any,
) -> dict | None:
    if not is_brick_item(item_name):
        return None
    if not is_block_unit(unit):
        return None

    try:
        q = float(str(qty).strip().replace(",", "")) if qty is not None and str(qty).strip() != "" else None
    except ValueError:
        return None
    if q is None or not math.isfinite(q) or q < 0:
        return None

    dims = parse_dims_mm(spec)
    if dims is None:
        return None
    v = volume_m3(dims)
    if v <= 0 or not math.isfinite(v):
        return None

    return {
        "转换单位": "m3",
        "计算数量": round(q * v, 6),
        "单块体积m3": round(v, 8),
        "规则": "砖砌块_块转m3",
    }
