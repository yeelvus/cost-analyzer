"""
单位/数量转换统一入口（独立于「物品名称规格标准化」）。

默认规则顺序（先匹配先用）：
  1. 钢筋网片_m2转kg  — 网片面积 → kg
  2. 钢筋_kg透传      — 钢筋/螺纹钢 已是 kg → 原样写入转换列
  3. 水泥砂浆_袋转kg  — 水泥/砂浆 袋 → kg（50kg/袋）
  4. 砖砌块_块转m3    — 砖/加气块 块 → m3
  5. 模板木方_转m2    — 模板/木方 张根块 → m2（主工作面）

每条规则暴露 try_convert(item_name=, spec=, unit=, qty=) -> dict|None
返回至少含：转换单位、计算数量
"""
from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent

DEFAULT_CONVERTERS = [
    "钢筋网片_m2转kg.py",
    "钢筋_kg透传.py",
    "水泥砂浆_袋转kg.py",
    "砖砌块_块转m3.py",
    "模板木方_转m2.py",
]

_LOADED: list[tuple[str, ModuleType]] | None = None


def _load_module(path: Path) -> ModuleType:
    name = f"unit_conv_{path.stem}_{abs(hash(str(path)))}"
    spec = importlib.util.spec_from_file_location(name, str(path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载: {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def load_converters(names: list[str] | None = None) -> list[tuple[str, ModuleType]]:
    global _LOADED
    if names is None and _LOADED is not None:
        return _LOADED

    use = names if names is not None else DEFAULT_CONVERTERS
    out: list[tuple[str, ModuleType]] = []
    for n in use:
        p = SCRIPT_DIR / n
        if not p.exists():
            print(f"  ⚠️ 转换规则不存在: {p.name}")
            continue
        try:
            mod = _load_module(p)
            if not hasattr(mod, "try_convert"):
                print(f"  ⚠️ {p.name} 缺少 try_convert，跳过")
                continue
            out.append((p.name, mod))
        except Exception as exc:
            print(f"  ⚠️ 加载 {p.name} 失败: {exc}")
    if names is None:
        _LOADED = out
    return out


def convert_row(
    *,
    item_name: Any = None,
    spec: Any = None,
    unit: Any = None,
    qty: Any = None,
    converters: list[tuple[str, ModuleType]] | None = None,
) -> dict:
    """
    对单行尝试转换。
    返回固定键：
      转换单位, 计算数量  （无法转换时为空字符串 / None）
      以及规则名等附加信息（若有）
    """
    mods = converters if converters is not None else load_converters()
    for rule_name, mod in mods:
        try:
            result = mod.try_convert(
                item_name=item_name,
                spec=spec,
                unit=unit,
                qty=qty,
            )
        except Exception:
            continue
        if not result:
            continue
        return {
            "转换单位": result.get("转换单位", ""),
            "计算数量": result.get("计算数量", ""),
            "转换规则": result.get("规则", rule_name),
            **{k: v for k, v in result.items() if k not in {"转换单位", "计算数量", "规则"}},
        }
    return {"转换单位": "", "计算数量": "", "转换规则": ""}


def apply_to_dataframe(df, *, name_col="物品名称", spec_col="规格", unit_col="单位", qty_col="数量"):
    """为 DataFrame 增加/覆盖 转换单位、计算数量 列。"""
    import pandas as pd

    mods = load_converters()
    conv_units = []
    conv_qtys = []
    for _, row in df.iterrows():
        r = convert_row(
            item_name=row.get(name_col, ""),
            spec=row.get(spec_col, ""),
            unit=row.get(unit_col, ""),
            qty=row.get(qty_col, ""),
            converters=mods,
        )
        conv_units.append(r.get("转换单位", ""))
        conv_qtys.append(r.get("计算数量", ""))
    out = df.copy()
    out["转换单位"] = conv_units
    out["计算数量"] = conv_qtys
    return out
