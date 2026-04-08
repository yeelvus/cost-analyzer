from __future__ import annotations

import argparse
import importlib.util
import re
from pathlib import Path
from typing import Any

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
DATA_ROOT = PROJECT_ROOT / "9_采购部采购台账"
EXCEL_FORMATTING_UTILS_PATH = Path('/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/excel_formatting_utils.py')
INPUT_ROOT = DATA_ROOT / "B_汇总数据" / "B_标准化汇总"
OUTPUT_ROOT = DATA_ROOT / "C_数据分析"
INPUT_GLOB = "*_材料采购_merged.xlsx"
DEFAULT_SHEET = "项目汇总"
DEFAULT_START_PERIOD = "2025-11"
DEFAULT_OUTPUT_NAME = "项目采购金额增加量汇总.xlsx"


def _load_module_from_path(module_name: str, module_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(module_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_excel_formatting_utils = _load_module_from_path('excel_formatting_utils', EXCEL_FORMATTING_UTILS_PATH)
apply_common_excel_layout = _excel_formatting_utils.apply_common_excel_layout


def normalize_col_name(name: str) -> str:
    s = str(name).strip().lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s_\-]+", "", s)
    return s


def find_column(columns: list[str], aliases: list[str]) -> str | None:
    norm_map = {normalize_col_name(c): c for c in columns}
    for alias in aliases:
        key = normalize_col_name(alias)
        if key in norm_map:
            return norm_map[key]

    for col in columns:
        col_key = normalize_col_name(col)
        for alias in aliases:
            alias_key = normalize_col_name(alias)
            if alias_key and alias_key in col_key:
                return col
    return None


def extract_period_from_name(path: Path) -> str | None:
    m = re.search(r"(?<!\d)((?:19|20)\d{2})年(0?[1-9]|1[0-2])月", path.stem)
    if not m:
        return None
    y = int(m.group(1))
    mm = int(m.group(2))
    return f"{y}-{mm:02d}"


def period_key(period: str) -> tuple[int, int]:
    y, m = period.split("-")
    return int(y), int(m)


def month_text(year_val: Any, month_val: Any) -> str | None:
    try:
        y = int(float(str(year_val).strip()))
        m = int(float(str(month_val).strip()))
        if y < 2000 or y > 2100 or m < 1 or m > 12:
            return None
        return f"{y}-{m:02d}"
    except Exception:
        return None


def normalize_text(val: Any) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in {"", "nan", "none", "null"}:
        return ""
    return s


def build_month_label_map(periods: list[str]) -> dict[str, str]:
    label_map: dict[str, str] = {}
    for p in periods:
        y, m = p.split("-")
        label_map[p] = f"{int(y)}年{int(m)}月"
    return label_map


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按项目统计月度采购金额及增加量（从指定起始月份到最新）")
    parser.add_argument("--input-root", default=str(INPUT_ROOT), help="输入目录，默认 B_汇总数据/B_标准化汇总")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="工作表名称，默认 合并数据")
    parser.add_argument("--start-period", default=DEFAULT_START_PERIOD, help="起始月份，格式 YYYY-MM，默认 2025-11")
    parser.add_argument("--output", default=str(OUTPUT_ROOT / DEFAULT_OUTPUT_NAME), help="输出 Excel 路径")
    return parser.parse_args()


def load_one_file(path: Path, sheet_name: str) -> tuple[pd.DataFrame | None, str]:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        return None, f"读取失败: {exc}"

    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)

    project_code_col = find_column(cols, ["项目编码", "编码", "项目编号", "工程编码", "合同编码"])
    project_name_col = find_column(cols, ["标准项目名称", "项目标准名称", "项目名称", "项目", "工程名称", "工程项目", "项目简称", "项目名"])
    amount_col = find_column(cols, ["总采购金额", "采购金额", "金额", "含税金额", "价税合计", "总金额", "合计金额", "采购总额", "本币金额", "总价"])

    period_from_name = extract_period_from_name(path)

    if amount_col is None:
        return None, f"缺少必要列: 金额列={amount_col}"

    if project_code_col is None and project_name_col is None:
        return None, "缺少必要列: 项目编码/标准项目名称至少要有一个"

    if period_from_name is None:
        return None, f"无法从文件名提取期间: {path.name}"

    if project_code_col is not None:
        code_series = df[project_code_col]
    else:
        code_series = pd.Series([""] * len(df), index=df.index)

    if project_name_col is not None:
        name_series = df[project_name_col]
    else:
        name_series = pd.Series([""] * len(df), index=df.index)

    work = pd.DataFrame(
        {
            "项目编码": code_series,
            "标准项目名称": name_series,
            "采购金额": df[amount_col],
        },
        index=df.index,
    )

    work["项目编码"] = work["项目编码"].map(normalize_text)
    work["标准项目名称"] = work["标准项目名称"].map(normalize_text)
    work = work[(work["项目编码"] != "") | (work["标准项目名称"] != "")].copy()

    work["采购金额"] = pd.to_numeric(work["采购金额"], errors="coerce")
    work = work[work["采购金额"].notna()].copy()

    work["期间"] = period_from_name
    work["来源文件"] = path.name

    return work[["期间", "项目编码", "标准项目名称", "采购金额", "来源文件"]], ""


def main() -> None:
    args = parse_args()

    input_root = Path(args.input_root).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    start_period = str(args.start_period).strip()

    if not re.fullmatch(r"(?:19|20)\d{2}-(0[1-9]|1[0-2])", start_period):
        raise ValueError(f"start-period 格式错误: {start_period}，应为 YYYY-MM")

    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    files = sorted([p for p in input_root.glob(INPUT_GLOB) if p.is_file()])
    if not files:
        raise FileNotFoundError(f"未找到输入文件: {input_root / INPUT_GLOB}")

    print(f"输入目录: {input_root}")
    print(f"命中文件数: {len(files)}")

    frames: list[pd.DataFrame] = []
    skipped: list[tuple[str, str]] = []

    for f in files:
        part, err = load_one_file(f, args.sheet)
        if part is None:
            skipped.append((f.name, err))
            continue
        frames.append(part)

    if not frames:
        raise RuntimeError("所有文件均无法解析，请检查列名与 sheet 名称")

    all_rows = pd.concat(frames, ignore_index=True)
    all_rows = all_rows[all_rows["期间"].notna()].copy()

    valid_periods = sorted(all_rows["期间"].dropna().unique().tolist(), key=period_key)
    valid_periods = [p for p in valid_periods if period_key(p) >= period_key(start_period)]

    if not valid_periods:
        raise RuntimeError(f"没有 {start_period} 及之后的可用期间")

    all_rows = all_rows[all_rows["期间"].isin(valid_periods)].copy()

    monthly_project_amount = (
        all_rows.groupby(["期间", "项目编码", "标准项目名称"], as_index=False)
        .agg(采购金额=("采购金额", "sum"))
        .sort_values(["项目编码", "标准项目名称", "期间"])
    )

    pivot = monthly_project_amount.pivot(index=["项目编码", "标准项目名称"], columns="期间", values="采购金额").fillna(0.0)
    pivot = pivot.reindex(columns=valid_periods, fill_value=0.0)

    delta = pivot.diff(axis=1)
    first_col = valid_periods[0]
    delta[first_col] = 0.0

    label_map = build_month_label_map(valid_periods)
    amount_wide = pivot.rename(columns=label_map).reset_index()
    delta_wide = delta.rename(columns=label_map).reset_index()

    month_cols = [label_map[p] for p in valid_periods]
    amount_wide = amount_wide[["项目编码", "标准项目名称", *month_cols]]
    delta_wide = delta_wide[["项目编码", "标准项目名称", *month_cols]]

    amount_wide = amount_wide.rename(columns={"标准项目名称": "项目名称"})
    delta_wide = delta_wide.rename(columns={"标准项目名称": "项目名称"})

    sort_cols = [label_map[p] for p in reversed(valid_periods)]
    delta_with_key = delta_wide.copy()
    delta_with_key["__编码为空"] = delta_with_key["项目编码"].astype(str).str.strip().eq("")
    delta_with_key = delta_with_key.sort_values(["__编码为空", *sort_cols], ascending=[True] + [False] * len(sort_cols))

    ordered_keys = delta_with_key[["项目编码", "项目名称"]].drop_duplicates()
    delta_wide = delta_with_key.drop(columns=["__编码为空"])
    amount_wide = ordered_keys.merge(amount_wide, on=["项目编码", "项目名称"], how="left")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        amount_wide.to_excel(writer, sheet_name="采购总金额", index=False)
        delta_wide.to_excel(writer, sheet_name="增加量", index=False)

        workbook = writer.book
        for sheet_name in ["采购总金额", "增加量"]:
            worksheet = workbook[sheet_name]
            apply_common_excel_layout(worksheet, numeric_start_col=3)

    print("\n处理完成")
    print(f"起始期间: {start_period}")
    print(f"结束期间: {valid_periods[-1]}")
    print(f"项目数量: {pivot.shape[0]}")
    print(f"输出文件: {output_path}")
    print(f"输出列: 项目编码, 项目名称, {', '.join(month_cols)}")

    if skipped:
        print("\n以下文件被跳过：")
        for name, reason in skipped:
            print(f"- {name}: {reason}")


if __name__ == "__main__":
    main()
