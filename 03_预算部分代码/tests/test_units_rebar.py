# -*- coding: utf-8 -*-
"""单位换算单测：python -m budget units-test  或  python tests/test_units_rebar.py"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from budget.units import (  # noqa: E402
    normalize_unit,
    sum_quantities_to_base,
    to_base_quantity,
    unit_matches_material,
)


def test_normalize():
    assert normalize_unit("t") == "TON"
    assert normalize_unit("吨") == "TON"
    assert normalize_unit("kg") == "KG"
    assert normalize_unit("千克") == "KG"
    assert normalize_unit("m³") == "M3"
    print("normalize_unit OK")


def test_rebar_mix():
    total, anomalies = sum_quantities_to_base([10, 500], ["吨", "kg"], "钢筋")
    assert abs(total - 10500) < 1e-6 and not anomalies
    print("10吨+500kg=10500kg OK")


def test_unknown():
    r = to_base_quantity(100, "袋", "钢筋")
    assert not r.ok
    print("unknown unit rejected OK")


def test_filter():
    assert unit_matches_material("吨", "钢筋")
    assert unit_matches_material("kg", "钢筋")
    print("filter OK")


def test_mortar_weight_to_m3():
    """砂浆导出常为 t/kg：按容重折 m³，不再报「未识别」。"""
    from budget.units import MORTAR_BULK_DENSITY_KG_PER_M3

    r = to_base_quantity(1.5, "t", "砂浆")
    assert r.ok, r.message
    expected = 1.5 * 1000 / MORTAR_BULK_DENSITY_KG_PER_M3
    assert abs(r.quantity_base - expected) < 1e-6, r.quantity_base

    r2 = to_base_quantity(1500, "kg", "砂浆")
    assert r2.ok and abs(r2.quantity_base - 1.0) < 1e-6

    total, anomalies = sum_quantities_to_base([1, 1500], ["t", "kg"], "砂浆")
    assert not anomalies
    assert abs(total - (1000 / MORTAR_BULK_DENSITY_KG_PER_M3 + 1.0)) < 1e-6
    print("mortar t/kg → m³ OK")


def test_message_not_false_unknown():
    """混凝土不允许用吨时，应说明类型不匹配，而不是「未识别」。"""
    r = to_base_quantity(1, "t", "混凝土")
    assert not r.ok
    assert "未识别" not in r.message
    assert r.unit_code == "TON"
    print("incompatible message OK")


if __name__ == "__main__":
    test_normalize()
    test_rebar_mix()
    test_unknown()
    test_filter()
    test_mortar_weight_to_m3()
    test_message_not_false_unknown()
    print("\n全部通过")
