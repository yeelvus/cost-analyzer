from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

from openpyxl import load_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
SUMMARY_INPUT_ROOT = SCRIPT_DIR / "B_汇总数据"
ANALYSIS_OUTPUT_ROOT = SCRIPT_DIR / "C_数据分析"
INPUT_GLOB = "*_材料采购_merged.xlsx"
OUTPUT_SUFFIX = "_桩规格标准化"
TARGET_KEYWORD = "桩"
TARGET_CATEGORY = "桩"  # 命中桩类行 → 材料归类填「桩」
PILE_COST_PREFIX = "（打桩费）"
PILE_COST_KEYWORDS = ("打桩", "施工")

SPLIT_PATTERN = re.compile(r"[xX×*＊]")
PART_PATTERN = re.compile(r"^(\d+(?:\.\d+)?)\s*(mm|cm|m)?$", re.IGNORECASE)
FALLBACK_PATTERN = re.compile(r"(\d+(?:\.\d+)?)(mm|cm|m)?", re.IGNORECASE)


def normalize_spec_text(spec_text: str) -> str:
    normalized = str(spec_text).strip()
    normalized = normalized.replace("毫米", "mm")
    normalized = normalized.replace("厘米", "cm")
    normalized = normalized.replace("米", "m")
    normalized = normalized.replace("Ｍ", "M")
    normalized = normalized.replace("ｍ", "m")
    normalized = normalized.replace("Ｘ", "x")
    normalized = normalized.replace("×", "x")
    normalized = normalized.replace("*", "x")
    normalized = normalized.replace("＊", "x")
    normalized = re.sub(r"\s+", "", normalized)
    return normalized.lower()


def parse_spec_parts(spec_value: object):
    if spec_value is None:
        return None

    spec_text = str(spec_value).strip()
    if not spec_text or spec_text.lower() == "nan":
        return None

    normalized = normalize_spec_text(spec_text)
    parts = [part for part in SPLIT_PATTERN.split(normalized) if part]

    def build_parsed_parts(raw_parts: list[str]):
        parsed_parts = []
        for part in raw_parts[:3]:
            match = PART_PATTERN.fullmatch(part)
            if not match:
                return None
            parsed_parts.append(
                {
                    "raw": part,
                    "value": float(match.group(1)),
                    "unit": match.group(2).lower() if match.group(2) else None,
                }
            )
        return parsed_parts

    if len(parts) == 3:
        parsed_parts = build_parsed_parts(parts)
        if parsed_parts is not None:
            return parsed_parts

    fallback_matches = FALLBACK_PATTERN.findall(normalized)
    if len(fallback_matches) < 3:
        return None

    parsed_parts = build_parsed_parts(["".join(match) for match in fallback_matches[:3]])
    if parsed_parts is None:
        return None

    return parsed_parts


def infer_side_unit(value: float, unit: str | None, peer_unit: str | None) -> str:
    if unit:
        return unit
    if peer_unit:
        return peer_unit
    if value < 10:
        return "m"
    if value < 100:
        return "cm"
    return "mm"


def infer_length_unit(
    value: float,
    unit: str | None,
    raw_side_values: list[float],
    side_units: list[str],
) -> str:
    if unit:
        return unit
    if value >= 1000:
        return "mm"
    if any(side_value < 10 for side_value in raw_side_values) or "m" in side_units:
        return "m"
    if "cm" in side_units and value >= 100:
        return "cm"
    return "m"


def convert_to_mm(value: float, unit: str) -> float:
    if unit == "mm":
        return value
    if unit == "cm":
        return value * 10
    if unit == "m":
        return value * 1000
    raise ValueError(f"不支持的截面单位: {unit}")


def convert_to_m(value: float, unit: str) -> float:
    if unit == "m":
        return value
    if unit == "cm":
        return value / 100
    if unit == "mm":
        return value / 1000
    raise ValueError(f"不支持的长度单位: {unit}")


def format_number(value: float) -> str:
    rounded = round(value)
    if math.isclose(value, rounded, abs_tol=1e-9):
        return str(int(rounded))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def standardize_pile_spec(spec_value: object) -> str | None:
    parts = parse_spec_parts(spec_value)
    if not parts or len(parts) != 3:
        return None

    side_units = [
        infer_side_unit(parts[0]["value"], parts[0]["unit"], parts[1]["unit"]),
        infer_side_unit(parts[1]["value"], parts[1]["unit"], parts[0]["unit"]),
    ]
    side_values_mm = [
        round(convert_to_mm(parts[0]["value"], side_units[0])),
        round(convert_to_mm(parts[1]["value"], side_units[1])),
    ]

    length_unit = infer_length_unit(
        parts[2]["value"],
        parts[2]["unit"],
        [parts[0]["value"], parts[1]["value"]],
        side_units,
    )
    length_value_m = convert_to_m(parts[2]["value"], length_unit)

    if min(side_values_mm) <= 0 or length_value_m <= 0:
        return None

    return (
        f"{format_number(side_values_mm[0])}x"
        f"{format_number(side_values_mm[1])}x"
        f"{format_number(length_value_m)}m"
    )


def should_prefix_pile_cost(item_name: object) -> bool:
    if item_name is None:
        return False

    item_text = str(item_name).strip()
    if not item_text or TARGET_KEYWORD not in item_text:
        return False
    if item_text.startswith(PILE_COST_PREFIX):
        return False

    return any(keyword in item_text for keyword in PILE_COST_KEYWORDS)


def get_header_map(worksheet):
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet):
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

    item_col = header_map["物品名称"]
    spec_col = header_map["规格"]
    category_col = header_map.get("材料归类")

    matched_rows = 0
    updated_rows = 0
    prefixed_rows = 0
    category_updated = 0
    unresolved_rows = []

    for row_idx in range(2, worksheet.max_row + 1):
        item_cell = worksheet.cell(row_idx, item_col)
        item_name = item_cell.value
        if item_name is None or TARGET_KEYWORD not in str(item_name):
            continue

        matched_rows += 1

        # 材料归类 → 桩（空/其他/错填均覆盖）
        if category_col is not None:
            cat_cell = worksheet.cell(row_idx, category_col)
            old_cat = "" if cat_cell.value is None else str(cat_cell.value).strip()
            if old_cat != TARGET_CATEGORY:
                cat_cell.value = TARGET_CATEGORY
                category_updated += 1

        if should_prefix_pile_cost(item_name):
            item_cell.value = f"{PILE_COST_PREFIX}{str(item_name).strip()}"
            item_name = item_cell.value
            prefixed_rows += 1

        spec_cell = worksheet.cell(row_idx, spec_col)
        standardized = standardize_pile_spec(spec_cell.value)

        if standardized is None:
            unresolved_rows.append(
                (row_idx, str(item_name).strip(), "" if spec_cell.value is None else str(spec_cell.value).strip())
            )
            continue

        old_spec = "" if spec_cell.value is None else str(spec_cell.value).strip()
        if old_spec != standardized:
            spec_cell.value = standardized
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows + category_updated,
        "prefixed_rows": prefixed_rows,
        "unresolved_rows": unresolved_rows,
        "skipped": False,
    }


def build_output_path(input_file: Path, suffix: str) -> Path:
    ANALYSIS_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    return ANALYSIS_OUTPUT_ROOT / f"{input_file.stem}{suffix}{input_file.suffix}"


def extract_year_month_from_name(path: Path) -> tuple[int, int] | None:
    m = re.search(r"(?<!\d)((?:19|20)\d{2})年(0?[1-9]|1[0-2])月", path.stem)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def find_latest_summary_file() -> Path | None:
    if not SUMMARY_INPUT_ROOT.is_dir():
        return None
    candidates = [p for p in SUMMARY_INPUT_ROOT.glob(INPUT_GLOB) if p.is_file()]
    if not candidates:
        return None

    def sort_key(p: Path):
        ym = extract_year_month_from_name(p)
        if ym is not None:
            return (1, ym[0], ym[1], p.stat().st_mtime)
        return (0, 0, 0, p.stat().st_mtime)

    return max(candidates, key=sort_key)


def parse_args():
    parser = argparse.ArgumentParser(description="标准化采购台账中桩相关物品的规格格式")
    parser.add_argument(
        "input_file",
        nargs="?",
        default="",
        help="待处理的 Excel 文件路径；不传时自动读取 B_汇总数据 下最新 *_材料采购_merged.xlsx",
    )
    parser.add_argument(
        "--suffix",
        default=OUTPUT_SUFFIX,
        help="输出文件名后缀，默认追加到原文件名后",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    if args.input_file:
        input_file = Path(args.input_file).expanduser().resolve()
    else:
        latest = find_latest_summary_file()
        if latest is None:
            raise FileNotFoundError(f"未在 {SUMMARY_INPUT_ROOT} 找到 {INPUT_GLOB}")
        input_file = latest.resolve()

    output_file = build_output_path(input_file, args.suffix)

    if not input_file.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    print(f"正在处理文件: {input_file}")
    workbook = load_workbook(input_file)

    all_results = []
    total_matched = 0
    total_updated = 0
    total_prefixed = 0
    total_unresolved = 0

    for worksheet in workbook.worksheets:
        result = process_worksheet(worksheet)
        all_results.append(result)
        if result["skipped"]:
            print(f"工作表 {worksheet.title}: 未找到“物品名称”或“规格”列，跳过")
            continue

        total_matched += result["matched_rows"]
        total_updated += result["updated_rows"]
        total_prefixed += result["prefixed_rows"]
        total_unresolved += len(result["unresolved_rows"])

        print(
            f"工作表 {worksheet.title}: 命中 {result['matched_rows']} 行，"
            f"名称加前缀 {result['prefixed_rows']} 行，"
            f"成功标准化 {result['updated_rows']} 行，"
            f"未识别 {len(result['unresolved_rows'])} 行"
        )

    workbook.save(output_file)

    print("\n处理完成")
    print(f"输出文件: {output_file}")
    print(f"命中总行数: {total_matched}")
    print(f"名称加前缀: {total_prefixed}")
    print(f"成功标准化: {total_updated}")
    print(f"未识别规格: {total_unresolved}")

    unresolved_preview = []
    for result in all_results:
        unresolved_preview.extend(
            (result["sheet_name"], row_idx, item_name, spec_value)
            for row_idx, item_name, spec_value in result["unresolved_rows"]
        )

    if unresolved_preview:
        print("\n以下是前 20 条未识别规格的样本:")
        for sheet_name, row_idx, item_name, spec_value in unresolved_preview[:20]:
            print(
                f"  - 工作表={sheet_name}, 行号={row_idx}, 物品名称={item_name}, 原规格={spec_value}"
            )


if __name__ == "__main__":
    main()