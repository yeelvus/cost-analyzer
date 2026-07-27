"""
钢筋网片：面积单位(m2) → 重量(kg)

- 仅物品名称含「钢筋网片」或「网片」（排除运费等）
- 仅当单位为 m2/㎡/平方米/平米 等面积单位时转换
- 其他钢筋（单位 kg）不转换

理论单重（双向网片，钢材密度 7.85 g/cm³）：
  kg/m² = 6.165 × d² × (1/s1 + 1/s2)
  d、s1、s2 单位均为 mm（正交等径网片）

数量：计算数量(kg) = 原数量(m²) × kg/m²
"""
from __future__ import annotations

import math
import re
from typing import Any

STEEL_DENSITY_FACTOR = 6.165  # ≈ 7850 * π/4 * 1e-3

MESH_NAME_OK = ("钢筋网片", "网片")
MESH_NAME_EXCLUDE = ("运费", "加工", "安装", "绑扎")

AREA_UNITS = {
    "m2", "m²", "㎡", "平方米", "平米", "平方", "平",
    "M2", "Ｍ2", "平方米",
}

# 标准规格：200x200 Φ8 3x6m
STD_SPACING = re.compile(r"(?<!\d)(\d{2,4})\s*[xX×*]\s*(\d{2,4})(?!\d)")
STD_DIA = re.compile(r"[ΦφØø]\s*(\d+(?:\.\d+)?)")
# 宽松：φ8@200x200 / φ8*20*20
RAW_PHI = re.compile(r"[ΦφØø]\s*(\d+(?:\.\d+)?)")
RAW_AT = re.compile(r"@\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+(?:\.\d+)?)")
RAW_STAR = re.compile(
    r"[ΦφØø]?\s*(\d+(?:\.\d+)?)\s*(?:mm)?\s*[xX×*@]\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+(?:\.\d+)?)",
    re.I,
)


def _norm_unit(unit: Any) -> str:
    if unit is None:
        return ""
    u = str(unit).strip().replace(" ", "")
    return u


def is_area_unit(unit: Any) -> bool:
    u = _norm_unit(unit)
    if not u:
        return False
    if u in AREA_UNITS:
        return True
    ul = u.lower()
    return ul in {"m2", "㎡", "m²"} or "平方" in u or u in {"平米", "平"}


def is_mesh_item(name: Any) -> bool:
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if any(k in s for k in MESH_NAME_EXCLUDE):
        return False
    return any(k in s for k in MESH_NAME_OK)


def _spacing_to_mm(a: float, b: float) -> tuple[float, float] | None:
    # 10~40 按 cm
    if 5 <= a <= 40 and 5 <= b <= 40:
        return a * 10, b * 10
    if 50 <= a <= 500 and 50 <= b <= 500:
        return a, b
    return None


def parse_mesh_params(spec: Any) -> tuple[float, float, float] | None:
    """
    从规格解析 (直径mm, 间距1mm, 间距2mm)。
    优先标准格式 200x200 Φ8 …
    """
    if spec is None:
        return None
    text = str(spec).strip()
    if not text or text.lower() == "nan":
        return None

    t = text.replace("×", "x").replace("＊", "x").replace("*", "x")
    t = t.replace("＠", "@").replace("φ", "Φ").replace("Ø", "Φ")

    dia = None
    s1 = s2 = None

    m = STD_DIA.search(t)
    if m:
        dia = float(m.group(1))

    # 标准：开头 200x200
    m = re.match(r"^\s*(\d{2,4})\s*[xX]\s*(\d{2,4})\b", t)
    if m:
        sp = _spacing_to_mm(float(m.group(1)), float(m.group(2)))
        if sp:
            s1, s2 = sp

    if s1 is None:
        m = RAW_AT.search(t)
        if m:
            sp = _spacing_to_mm(float(m.group(1)), float(m.group(2)))
            if sp:
                s1, s2 = sp

    if dia is None:
        m = RAW_PHI.search(t)
        if m:
            dia = float(m.group(1))

    if (s1 is None or dia is None) and re.search(r"\d", t):
        # φ8x20x20 / 6mmx200x200
        m = re.search(
            r"(?:[Φφ]?\s*)(\d+(?:\.\d+)?)\s*(?:mm)?\s*[xX@]\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)",
            t,
            re.I,
        )
        if m:
            a, b, c = float(m.group(1)), float(m.group(2)), float(m.group(3))
            # 直径在前：d x s x s
            if 2 <= a <= 20:
                dia = dia or a
                sp = _spacing_to_mm(b, c)
                if sp:
                    s1, s2 = sp
            # 间距在前：s x s x d
            elif 2 <= c <= 20 and b >= 50:
                dia = dia or c
                sp = _spacing_to_mm(a, b)
                if sp:
                    s1, s2 = sp

    if dia is None or s1 is None or s2 is None:
        return None
    if not (2 <= dia <= 25):
        return None
    if s1 <= 0 or s2 <= 0:
        return None
    return dia, s1, s2


def kg_per_m2(diameter_mm: float, spacing1_mm: float, spacing2_mm: float) -> float:
    """正交等径钢筋网片理论重量 kg/m²。"""
    return STEEL_DENSITY_FACTOR * (diameter_mm ** 2) * (1.0 / spacing1_mm + 1.0 / spacing2_mm)


def try_convert(
    *,
    item_name: Any,
    spec: Any,
    unit: Any,
    qty: Any,
) -> dict | None:
    """
    成功返回:
      {
        "转换单位": "kg",
        "计算数量": float,
        "单重kg_m2": float,
        "规则": "钢筋网片_m2转kg",
      }
    """
    if not is_mesh_item(item_name):
        return None
    if not is_area_unit(unit):
        return None

    try:
        q = float(str(qty).strip().replace(",", "")) if qty is not None and str(qty).strip() != "" else None
    except ValueError:
        return None
    if q is None or not math.isfinite(q) or q < 0:
        return None

    params = parse_mesh_params(spec)
    if params is None:
        return None
    d, s1, s2 = params
    w = kg_per_m2(d, s1, s2)
    if w <= 0 or not math.isfinite(w):
        return None

    return {
        "转换单位": "kg",
        "计算数量": round(q * w, 3),
        "单重kg_m2": round(w, 4),
        "规则": "钢筋网片_m2转kg",
    }
