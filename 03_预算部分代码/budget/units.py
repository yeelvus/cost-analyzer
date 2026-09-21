# -*- coding: utf-8 -*-
"""
预算材料单位规范化与换算。

规则：
- 钢筋等重量类 → 基准 kg（1 t / 吨 = 1000 kg）
- 混凝土/砂浆/碎石 → 基准 m³
  * 砂浆/碎石在导出中常出现 t、kg：按容重折算为 m³（可配置）
- 未知单位或与材料不匹配：不静默按 1 处理，ok=False，记入异常表
"""

from __future__ import annotations

import re
from dataclasses import dataclass


# 材料 → 汇总基准单位（展示用）
MATERIAL_BASE_UNITS = {
    "钢筋": "kg",
    "混凝土": "m³",
    "砂浆": "m³",
    "碎石": "m³",
    "砖": "块",
    "其他材料": "",
}

# 干粉砂浆 / 散装砂浆近似容重（kg/m³），用于 t、kg → m³
# 行业常用约 1400~1800，取 1500 作默认；可在 config 中覆盖
MORTAR_BULK_DENSITY_KG_PER_M3 = 1500.0
# 碎石松散容重大约 1500~1700 kg/m³
AGGREGATE_BULK_DENSITY_KG_PER_M3 = 1600.0

# 标准码 → 换算到材料基准单位的系数
_TO_KG = {
    "KG": 1.0,
    "TON": 1000.0,
}
_TO_M3_VOLUME_ONLY = {
    "M3": 1.0,
}
_TO_PCS = {
    "PCS": 1.0,
}


def _m3_table_with_weight(density_kg_per_m3: float) -> dict[str, float]:
    """体积基准表：允许 m³，以及按容重把 kg/吨 折成 m³。"""
    d = float(density_kg_per_m3) if density_kg_per_m3 > 0 else 1500.0
    return {
        "M3": 1.0,
        "KG": 1.0 / d,           # kg → m³
        "TON": 1000.0 / d,       # t → m³
    }


# 材料 → 允许的标准码及换算表（动态生成含重量折算的砂浆/碎石）
MATERIAL_UNIT_TABLES: dict[str, dict[str, float]] = {
    "钢筋": _TO_KG,
    "混凝土": _TO_M3_VOLUME_ONLY,  # 混凝土一般不应按吨折体积（易误判），只认 m³
    "砂浆": _m3_table_with_weight(MORTAR_BULK_DENSITY_KG_PER_M3),
    "碎石": _m3_table_with_weight(AGGREGATE_BULK_DENSITY_KG_PER_M3),
    "砖": _TO_PCS,
    "其他材料": {},  # 不强制
}

# 原始单位字符串 → 标准码
_UNIT_ALIASES = {
    "kg": "KG",
    "kgs": "KG",
    "kilogram": "KG",
    "kilograms": "KG",
    "千克": "KG",
    "公斤": "KG",
    "t": "TON",
    "ton": "TON",
    "tons": "TON",
    "tonne": "TON",
    "tonnes": "TON",
    "吨": "TON",
    "噸": "TON",
    "m3": "M3",
    "m³": "M3",
    "m^3": "M3",
    "立方米": "M3",
    "立方": "M3",
    "方": "M3",
    "块": "PCS",
    "块。": "PCS",
    "个": "PCS",
    "pcs": "PCS",
    "pc": "PCS",
}


@dataclass
class QtyConvertResult:
    ok: bool
    quantity_base: float
    unit_code: str  # KG / TON / M3 / PCS / UNKNOWN / EMPTY
    unit_raw: str
    factor: float
    message: str = ""


def set_bulk_densities(
    mortar_kg_per_m3: float | None = None,
    aggregate_kg_per_m3: float | None = None,
) -> None:
    """运行时调整砂浆/碎石容重，并刷新换算表。"""
    global MORTAR_BULK_DENSITY_KG_PER_M3, AGGREGATE_BULK_DENSITY_KG_PER_M3
    if mortar_kg_per_m3 is not None and mortar_kg_per_m3 > 0:
        MORTAR_BULK_DENSITY_KG_PER_M3 = float(mortar_kg_per_m3)
        MATERIAL_UNIT_TABLES["砂浆"] = _m3_table_with_weight(MORTAR_BULK_DENSITY_KG_PER_M3)
    if aggregate_kg_per_m3 is not None and aggregate_kg_per_m3 > 0:
        AGGREGATE_BULK_DENSITY_KG_PER_M3 = float(aggregate_kg_per_m3)
        MATERIAL_UNIT_TABLES["碎石"] = _m3_table_with_weight(AGGREGATE_BULK_DENSITY_KG_PER_M3)


def _clean_unit_text(unit) -> str:
    if unit is None:
        return ""
    try:
        import pandas as pd

        if isinstance(unit, float) and pd.isna(unit):
            return ""
    except Exception:
        pass
    text = str(unit).strip()
    if not text or text.lower() in {"nan", "none", "-", "—", "－", "/", "\\"}:
        return ""
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    return text


def normalize_unit(unit) -> str:
    """将原始单位规范为标准码：KG / TON / M3 / PCS / UNKNOWN / EMPTY"""
    raw = _clean_unit_text(unit)
    if not raw:
        return "EMPTY"

    key = raw.lower()
    if key in _UNIT_ALIASES:
        return _UNIT_ALIASES[key]
    if raw in _UNIT_ALIASES:
        return _UNIT_ALIASES[raw]

    lower = raw.lower()
    if "吨" in raw or "噸" in raw or "ton" in lower:
        return "TON"
    if re.fullmatch(r"t", lower) or re.search(r"\(t\)", lower):
        return "TON"
    if lower == "t":
        return "TON"
    if "千克" in raw or "公斤" in raw or re.search(r"kg", lower):
        return "KG"
    if "立方米" in raw or "立方" in raw or "m³" in raw or "m3" in lower or "m^3" in lower:
        return "M3"
    if raw in {"块", "个"} or lower in {"pcs", "pc"}:
        return "PCS"
    if lower == "kg":
        return "KG"

    return "UNKNOWN"


def unit_matches_material(unit, material_name: str) -> bool:
    """判断单位是否属于该主材允许的类型（用于筛选明细行）。"""
    code = normalize_unit(unit)
    table = MATERIAL_UNIT_TABLES.get(material_name, {})
    if not table:
        return True
    if code in table:
        return True
    if code == "EMPTY":
        return True
    return False


def _base_unit_label(material_name: str) -> str:
    return MATERIAL_BASE_UNITS.get(material_name, "") or "基准单位"


def to_base_quantity(qty, unit, material_name: str = "钢筋") -> QtyConvertResult:
    """
    将数量换算为材料基准单位数量。
    - 识别成功且在换算表内：ok=True
    - 单位字符串无法解析：UNKNOWN
    - 能识别但与材料基准不兼容：ok=False，文案说明「类型不匹配」
    """
    unit_raw = _clean_unit_text(unit)
    try:
        q = float(qty)
    except (TypeError, ValueError):
        return QtyConvertResult(False, 0.0, "UNKNOWN", unit_raw, 0.0, "数量无法解析")

    if q != q:  # NaN
        return QtyConvertResult(False, 0.0, "UNKNOWN", unit_raw, 0.0, "数量为空")
    if q == 0:
        return QtyConvertResult(True, 0.0, normalize_unit(unit_raw), unit_raw, 1.0, "")

    code = normalize_unit(unit_raw)
    table = MATERIAL_UNIT_TABLES.get(material_name)

    # 其他材料：不换算
    if table is not None and len(table) == 0:
        return QtyConvertResult(
            True, q, code if code not in ("EMPTY",) else "UNKNOWN", unit_raw, 1.0, "其他材料不换算"
        )

    if table is None:
        table = _TO_KG if material_name == "钢筋" else {}

    if code == "EMPTY":
        return QtyConvertResult(
            False, 0.0, "EMPTY", unit_raw, 0.0, f"单位为空，未计入{material_name}数量合计"
        )

    if code == "UNKNOWN":
        return QtyConvertResult(
            False,
            0.0,
            "UNKNOWN",
            unit_raw,
            0.0,
            f"无法解析单位「{unit_raw}」，未计入{material_name}数量合计",
        )

    if code not in table:
        base = _base_unit_label(material_name)
        allowed = "、".join(sorted(table.keys()))
        return QtyConvertResult(
            False,
            0.0,
            code,
            unit_raw,
            0.0,
            (
                f"单位「{unit_raw}」已识别为{code}，但{material_name}数量基准为{base}"
                f"（允许：{allowed}），未计入合计"
            ),
        )

    factor = table[code]
    note = ""
    if material_name in ("砂浆", "碎石") and code in ("KG", "TON"):
        dens = (
            MORTAR_BULK_DENSITY_KG_PER_M3
            if material_name == "砂浆"
            else AGGREGATE_BULK_DENSITY_KG_PER_M3
        )
        note = f"按容重{dens:.0f}kg/m³折算为m³"
    return QtyConvertResult(True, q * factor, code, unit_raw, factor, note)


def sum_quantities_to_base(
    qty_list,
    unit_list,
    material_name: str,
    price_list=None,
    require_positive_price: bool = False,
) -> tuple[float, list[dict]]:
    """批量换算并求和。返回 (合计基准数量, 异常行列表)。"""
    total = 0.0
    anomalies: list[dict] = []
    n = len(qty_list)
    for i in range(n):
        q = qty_list[i]
        u = unit_list[i] if i < len(unit_list) else ""
        price = None
        if price_list is not None and i < len(price_list):
            try:
                price = float(price_list[i])
            except (TypeError, ValueError):
                price = None

        if require_positive_price and price is not None and not (price > 0):
            continue

        result = to_base_quantity(q, u, material_name)
        if result.ok:
            total += result.quantity_base
        else:
            if q is not None:
                try:
                    qf = float(q)
                    if qf == qf and qf != 0:
                        anomalies.append(
                            {
                                "材料": material_name,
                                "原始数量": q,
                                "原始单位": result.unit_raw,
                                "标准码": result.unit_code,
                                "原因": result.message,
                            }
                        )
                except (TypeError, ValueError):
                    pass
    return total, anomalies
