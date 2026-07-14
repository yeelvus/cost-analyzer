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
OUTPUT_SUFFIX = "_钢筋规格标准化"

TARGET_KEYWORDS = ("钢筋网片", "螺纹钢", "钢筋")
MESH_KEYWORDS = ("钢筋网片", "网片")
REBAR_KEYWORDS = ("螺纹钢",)

PHI_PATTERN = re.compile(r"[φΦØøDd]\s*(\d+(?:\.\d+)?)")
HRB_DIA_PATTERN = re.compile(r"(hrb|hpb|rrb)\s*[- ]?\s*(\d{3,4})([a-zA-Z]*)\s*[- ]?\s*(\d{1,2}(?:\.\d+)?)", re.IGNORECASE)
GRADE_PATTERN = re.compile(r"(hrb|hpb|rrb)\s*[- ]?\s*(\d{3,4})([a-zA-Z]*)", re.IGNORECASE)
MM_DIA_PATTERN = re.compile(r"(?<!\d)(\d{1,2}(?:\.\d+)?)\s*(?:mm|毫米)\b", re.IGNORECASE)
RAW_DIA_PATTERN = re.compile(r"(?<!\d)(\d{1,2}(?:\.\d+)?)(?!\d)")
LENGTH_PATTERN = re.compile(r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:m|米)\b", re.IGNORECASE)
SPACING_PATTERN = re.compile(r"(?<!\d)(\d{2,3})\s*[xX×*＊/]\s*(\d{2,3})(?!\d)")
SHEET_SIZE_PATTERN = re.compile(r"(?<!\d)(\d+(?:\.\d+)?)\s*[xX×*＊]\s*(\d+(?:\.\d+)?)\s*(?:m|米)", re.IGNORECASE)


def format_number(value: float) -> str:
    rounded = round(value)
    if math.isclose(value, rounded, abs_tol=1e-9):
        return str(int(rounded))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""

    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("，", ",")
    text = text.replace("：", ":")
    text = text.replace("＋", "+")
    text = text.replace("Ｘ", "x").replace("×", "x").replace("＊", "x").replace("*", "x")
    text = text.replace("Ｍ", "m").replace("ｍ", "m")
    text = text.replace("毫米", "mm").replace("米", "m")
    text = re.sub(r"\s+", "", text)
    return text


def find_diameter(text: str, *, prefer_mesh: bool = False) -> float | None:
    if not text:
        return None

    match = PHI_PATTERN.search(text)
    if match:
        return float(match.group(1))

    match = HRB_DIA_PATTERN.search(text)
    if match:
        return float(match.group(4))

    match = MM_DIA_PATTERN.search(text)
    if match:
        return float(match.group(1))

    # 钢筋直径通常在 4~50mm，网片钢丝常见 3~12mm；此处按关键字场景约束范围
    low, high = (2, 20) if prefer_mesh else (4, 50)
    for m in RAW_DIA_PATTERN.finditer(text):
        val = float(m.group(1))
        if low <= val <= high:
            return val
    return None


def find_grade(text: str) -> str | None:
    match = GRADE_PATTERN.search(text)
    if not match:
        return None
    prefix = match.group(1).upper()
    level = match.group(2)
    suffix = match.group(3).upper() if match.group(3) else ""
    return f"{prefix}{level}{suffix}"


def find_length_m(text: str) -> float | None:
    # 钢筋常见长度 6m/9m/12m，避免把网片间距误识别成长度
    for m in LENGTH_PATTERN.finditer(text):
        val = float(m.group(1))
        if 1 <= val <= 18:
            return val
    return None


def standardize_rebar_or_steel(spec_value: object) -> str | None:
    text = normalize_text(spec_value)
    if not text:
        return None

    diameter = find_diameter(text, prefer_mesh=False)
    if diameter is None:
        return None

    grade = find_grade(text)
    length = find_length_m(text)

    base = f"Φ{format_number(diameter)}"
    if length is not None:
        base = f"{base}x{format_number(length)}m"
    if grade:
        return f"{grade}-{base}"
    return base


def find_spacing(text: str) -> tuple[int, int] | None:
    for m in SPACING_PATTERN.finditer(text):
        a = int(m.group(1))
        b = int(m.group(2))
        # 网片间距一般是 50~400mm
        if 50 <= a <= 400 and 50 <= b <= 400:
            return a, b
    return None


def find_sheet_size(text: str) -> tuple[float, float] | None:
    for m in SHEET_SIZE_PATTERN.finditer(text):
        a = float(m.group(1))
        b = float(m.group(2))
        # 网片尺寸一般在 0.5~12m
        if 0.5 <= a <= 12 and 0.5 <= b <= 12:
            return a, b
    return None


def standardize_mesh(spec_value: object) -> str | None:
    """网片规格委托给专用模块，输出「间距 型号 大小」，如 200x200 Φ8 3x6m。"""
    try:
        from pathlib import Path
        import importlib.util

        mesh_path = Path(__file__).resolve().parent / "钢筋网片规格标准化.py"
        if mesh_path.exists():
            spec = importlib.util.spec_from_file_location("_mesh_spec_std", str(mesh_path))
            if spec is not None and spec.loader is not None:
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                return mod.standardize_mesh(spec_value)
    except Exception:
        pass

    # 回退：旧逻辑（尽量给出直径）
    text = normalize_text(spec_value)
    if not text:
        return None
    diameter = find_diameter(text, prefer_mesh=True)
    if diameter is None:
        return None
    return f"Φ{format_number(diameter)}"


def detect_category(item_name: object) -> str | None:
    if item_name is None:
        return None

    item_text = str(item_name).strip()
    if not item_text:
        return None

    if any(keyword in item_text for keyword in MESH_KEYWORDS):
        return "mesh"
    if any(keyword in item_text for keyword in REBAR_KEYWORDS):
        return "rebar"
    if "钢筋" in item_text:
        return "steel"
    return None


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
            "unresolved_rows": [],
            "skipped": True,
        }

    item_col = header_map["物品名称"]
    spec_col = header_map["规格"]

    matched_rows = 0
    updated_rows = 0
    unresolved_rows = []

    for row_idx in range(2, worksheet.max_row + 1):
        item_name = worksheet.cell(row_idx, item_col).value
        category = detect_category(item_name)
        if category is None:
            continue

        matched_rows += 1
        spec_cell = worksheet.cell(row_idx, spec_col)

        if category == "mesh":
            standardized = standardize_mesh(spec_cell.value)
        else:
            standardized = standardize_rebar_or_steel(spec_cell.value)

        if standardized is None:
            unresolved_rows.append(
                (
                    row_idx,
                    str(item_name).strip(),
                    "" if spec_cell.value is None else str(spec_cell.value).strip(),
                    category,
                )
            )
            continue

        old_text = "" if spec_cell.value is None else str(spec_cell.value).strip()
        if old_text != standardized:
            spec_cell.value = standardized
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows,
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
    parser = argparse.ArgumentParser(description="标准化采购台账中钢筋/螺纹钢/钢筋网片的规格格式")
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
    total_unresolved = 0

    for worksheet in workbook.worksheets:
        result = process_worksheet(worksheet)
        all_results.append(result)

        if result["skipped"]:
            print(f"工作表 {worksheet.title}: 未找到“物品名称”或“规格”列，跳过")
            continue

        total_matched += result["matched_rows"]
        total_updated += result["updated_rows"]
        total_unresolved += len(result["unresolved_rows"])

        print(
            f"工作表 {worksheet.title}: 命中 {result['matched_rows']} 行，"
            f"成功标准化 {result['updated_rows']} 行，"
            f"未识别 {len(result['unresolved_rows'])} 行"
        )

    workbook.save(output_file)

    print("\n处理完成")
    print(f"输出文件: {output_file}")
    print(f"命中总行数: {total_matched}")
    print(f"成功标准化: {total_updated}")
    print(f"未识别规格: {total_unresolved}")

    unresolved_preview = []
    for result in all_results:
        unresolved_preview.extend(
            (result["sheet_name"], row_idx, item_name, spec_value, category)
            for row_idx, item_name, spec_value, category in result["unresolved_rows"]
        )

    if unresolved_preview:
        print("\n以下是前 20 条未识别规格样本:")
        for sheet_name, row_idx, item_name, spec_value, category in unresolved_preview[:20]:
            print(
                f"  - 工作表={sheet_name}, 行号={row_idx}, 类别={category}, 物品名称={item_name}, 原规格={spec_value}"
            )


if __name__ == "__main__":
    main()
