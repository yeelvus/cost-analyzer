#!/usr/bin/env python3
"""
砖 / 加气块 型号（规格）标准化。

适用范围：
  - 物品名称含：空心砖、实心砖、加气块、加气砖、水泥砖、红砖、砌块 等
  - 排除：砂浆、胶泥、砖刀、抹灰 等非砌块本体
  - 单位为「块」（仅处理块，不改 m2/袋 等）

标准规格格式（毫米，三维，从大到小排序）：
  390x190x190mm
  200x200x600mm

尺寸识别：
  - 190*190*390 / 190x190x390mm → 毫米
  - 19*19*39 / 19*19*39cm → 厘米转毫米
  - 7.5cm*20cm*60cm → 按标注单位转换

无法解析的规格：
  - 用「同物品名称 + 单价相近(默认±15%)」在本表已成功标准化行中匹配
  - 多候选时优先单价差最小，其次同规格出现次数最多

材料归类：
  - 命中且单位为「块」→ 材料归类 填「砖」（覆盖空/其他/错填）
"""
from __future__ import annotations

import math
import re
from collections import Counter, defaultdict

NAME_INCLUDE = (
    "空心砖", "实心砖", "加气块", "加气砖", "水泥砖", "红砖", "砌块",
    "多孔砖", "页岩砖", "标砖", "通风砖", "植草砖", "面包砖", "墙砖",
    "水泥空心砖", "实心水泥砖", "人行道砖", "人行横道砖", "地板砖",
)
# 名称包含这些则跳过（非砌块本体）
NAME_EXCLUDE = (
    "砂浆", "胶泥", "砖刀", "抹灰", "粘结", "粘接", "粉刷", "机", "锯",
)

BLOCK_UNITS = {"块"}  # 严格只处理「块」
PRICE_TOLERANCE = 0.15  # 单价相对偏差 15%
TARGET_CATEGORY = "砖"
# 尺寸数字：12 / 11.5 / 7.5 等
NUM = r"(\d+(?:\.\d+)?)"
UNIT = r"(mm|cm|m|毫米|厘米|米)?"
# 三段尺寸，分隔符 x * × ＊
DIM3_RE = re.compile(
    rf"{NUM}\s*{UNIT}\s*[xX×*＊]\s*{NUM}\s*{UNIT}\s*[xX×*＊]\s*{NUM}\s*{UNIT}",
    re.IGNORECASE,
)


def format_mm(value: float) -> str:
    rounded = round(value)
    if math.isclose(value, rounded, abs_tol=0.05):
        return str(int(rounded))
    # 保留一位小数（极少见）
    return f"{value:.1f}".rstrip("0").rstrip(".")


def normalize_unit_token(u: str | None) -> str | None:
    if not u:
        return None
    t = str(u).strip().lower()
    t = t.replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")
    if t in {"mm", "cm", "m"}:
        return t
    return None


def to_mm(value: float, unit: str) -> float:
    if unit == "mm":
        return value
    if unit == "cm":
        return value * 10.0
    if unit == "m":
        return value * 1000.0
    return value


def infer_units_per_dim(values: list[float], explicit_units: list[str | None]) -> list[str]:
    """
    为每一维推断单位。

    - 有显式单位：该维用显式；未标注维跟随众数显式单位
    - 无显式单位：
        · 全 ≥100 → 全 mm（190*190*390）
        · 最大 ≤80 → 全 cm（19*19*39、8*9*19）
        · 混合：最大≥100 且 最小≤50 → 大值 mm、小值 cm
          （8*9*190、7.5*200*600）
        · 其余 → mm
    """
    if any(explicit_units):
        present = [u for u in explicit_units if u]
        majority = Counter(present).most_common(1)[0][0]
        return [u if u else majority for u in explicit_units]

    vmax = max(values)
    vmin = min(values)

    if vmax >= 100 and vmin <= 50:
        # 混合尺度：一边像 mm、一边像 cm
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


def parse_brick_dims(spec_value: object) -> tuple[float, float, float] | None:
    """解析为三个毫米尺寸（未排序）。"""
    if spec_value is None:
        return None
    text = str(spec_value).strip()
    if not text or text.lower() == "nan":
        return None

    # 去掉括号杂讯
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    text = text.replace("Ｘ", "x").replace("×", "x").replace("＊", "x").replace("*", "x")
    text = text.replace("毫米", "mm").replace("厘米", "cm").replace("米", "m")

    m = DIM3_RE.search(text)
    if not m:
        return None

    v1, u1, v2, u2, v3, u3 = m.groups()
    values = [float(v1), float(v2), float(v3)]
    units_raw = [normalize_unit_token(u1), normalize_unit_token(u2), normalize_unit_token(u3)]
    units = infer_units_per_dim(values, units_raw)
    dims = [
        to_mm(values[0], units[0]),
        to_mm(values[1], units[1]),
        to_mm(values[2], units[2]),
    ]
    # 合理性：砌块边长约 20~800mm
    if any(d < 15 or d > 1200 for d in dims):
        return None
    return dims[0], dims[1], dims[2]


def format_standard_spec(dims: tuple[float, float, float]) -> str:
    """统一为 大x中x小 + mm。"""
    a, b, c = sorted(dims, reverse=True)
    return f"{format_mm(a)}x{format_mm(b)}x{format_mm(c)}mm"


def standardize_brick_spec(spec_value: object) -> str | None:
    dims = parse_brick_dims(spec_value)
    if dims is None:
        return None
    return format_standard_spec(dims)


def is_brick_item(item_name: object) -> bool:
    if item_name is None:
        return False
    name = str(item_name).strip()
    if not name:
        return False
    if any(k in name for k in NAME_EXCLUDE):
        return False
    return any(k in name for k in NAME_INCLUDE) or (
        "砖" in name and "瓷砖胶" not in name
    )


def is_block_unit(unit_value: object) -> bool:
    if unit_value is None:
        return False
    return str(unit_value).strip() in BLOCK_UNITS


def _to_price(value) -> float | None:
    try:
        if value is None:
            return None
        if isinstance(value, str):
            value = value.strip().replace(",", "")
            if not value:
                return None
        num = float(value)
        if math.isnan(num):
            return None
        return num
    except (TypeError, ValueError):
        return None


def _price_close(a: float, b: float, tol: float = PRICE_TOLERANCE) -> bool:
    if a == 0 and b == 0:
        return True
    base = max(abs(a), abs(b), 1e-9)
    return abs(a - b) / base <= tol


def get_header_map(worksheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet) -> dict:
    header_map = get_header_map(worksheet)
    need = {"物品名称", "规格"}
    if not need.issubset(header_map):
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
    unit_col = header_map.get("单位") or header_map.get("标准单位")
    price_col = header_map.get("单价")
    category_col = header_map.get("材料归类")

    matched_rows = 0
    updated_rows = 0
    category_updated = 0
    unresolved_rows: list[tuple] = []

    # 第一遍：可解析的直接标准化，并建立 (名称, 标准规格) 与价格索引
    resolved: dict[int, str] = {}
    price_index: dict[str, list[tuple[float, str]]] = defaultdict(list)
    spec_freq: dict[str, Counter] = defaultdict(Counter)

    candidates: list[dict] = []

    for row_idx in range(2, worksheet.max_row + 1):
        item_name = worksheet.cell(row_idx, name_col).value
        if not is_brick_item(item_name):
            continue

        unit_val = worksheet.cell(row_idx, unit_col).value if unit_col else None
        if unit_col is not None and not is_block_unit(unit_val):
            continue

        matched_rows += 1
        name_key = str(item_name).strip()
        spec_val = worksheet.cell(row_idx, spec_col).value
        price = _to_price(worksheet.cell(row_idx, price_col).value) if price_col else None
        old_text = "" if spec_val is None else str(spec_val).strip()

        # 材料归类 → 砖（凡命中的块类砌块行都填）
        if category_col is not None:
            cat_cell = worksheet.cell(row_idx, category_col)
            old_cat = "" if cat_cell.value is None else str(cat_cell.value).strip()
            if old_cat != TARGET_CATEGORY:
                cat_cell.value = TARGET_CATEGORY
                category_updated += 1

        std = standardize_brick_spec(spec_val)
        candidates.append(
            {
                "row": row_idx,
                "name": name_key,
                "old": old_text,
                "price": price,
                "std": std,
            }
        )
        if std:
            resolved[row_idx] = std
            spec_freq[name_key][std] += 1
            if price is not None:
                price_index[name_key].append((price, std))

    # 第二遍：无法解析的按 名称 + 单价 匹配
    for item in candidates:
        if item["std"]:
            continue
        name = item["name"]
        price = item["price"]
        matched_spec = None

        if price is not None and name in price_index:
            scored: list[tuple[float, int, str]] = []
            for p, std in price_index[name]:
                if not _price_close(price, p):
                    continue
                gap = abs(price - p)
                pop = spec_freq[name][std]
                scored.append((gap, -pop, std))
            if scored:
                scored.sort()
                matched_spec = scored[0][2]

        if not matched_spec and name in spec_freq and spec_freq[name]:
            matched_spec = spec_freq[name].most_common(1)[0][0]

        if matched_spec:
            item["std"] = matched_spec
            resolved[item["row"]] = matched_spec
            spec_freq[name][matched_spec] += 1
            if price is not None:
                price_index[name].append((price, matched_spec))
        else:
            unresolved_rows.append(
                (item["row"], name, item["old"], "brick")
            )

    # 写回规格
    for item in candidates:
        std = item.get("std")
        if not std:
            continue
        cell = worksheet.cell(item["row"], spec_col)
        if item["old"] != std:
            cell.value = std
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows + category_updated,
        "prefixed_rows": category_updated,  # 复用字段：归类更新数
        "unresolved_rows": unresolved_rows,
        "skipped": False,
    }


# ── 可选独立运行 ──────────────────────────────────────────
def main():
    import argparse
    from openpyxl import load_workbook
    from pathlib import Path

    parser = argparse.ArgumentParser(description="砖/加气块规格标准化")
    parser.add_argument("input_file", help="输入 Excel")
    parser.add_argument("--output-file", default="", help="输出路径，默认加后缀")
    args = parser.parse_args()

    inp = Path(args.input_file).expanduser().resolve()
    out = (
        Path(args.output_file).expanduser().resolve()
        if args.output_file
        else inp.with_name(f"{inp.stem}_砖加气块规格标准化{inp.suffix}")
    )
    wb = load_workbook(inp)
    total_m = total_u = total_r = 0
    for ws in wb.worksheets:
        r = process_worksheet(ws)
        if r.get("skipped"):
            print(f"[{ws.title}] 跳过")
            continue
        total_m += r["matched_rows"]
        total_u += r["updated_rows"]
        total_r += len(r["unresolved_rows"])
        print(
            f"[{ws.title}] 命中 {r['matched_rows']} 更新 {r['updated_rows']} "
            f"未识别 {len(r['unresolved_rows'])}"
        )
        for row_idx, name, spec, _ in r["unresolved_rows"][:10]:
            print(f"  未识别 行{row_idx}: {name} | {spec}")
    wb.save(out)
    print(f"\n完成 命中{total_m} 更新{total_u} 未识别{total_r}")
    print(f"输出: {out}")


if __name__ == "__main__":
    main()
