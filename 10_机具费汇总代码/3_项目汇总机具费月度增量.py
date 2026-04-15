# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import importlib.util
from pathlib import Path

import pandas as pd


INPUT_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum")
OUTPUT_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/C_机具租赁分析报告")
TARGET_SHEET = "项目汇总机具费"
OUTPUT_PREFIX = "机具费_项目月度总额与增加量"

FILE_MONTH_RE = re.compile(r"(20\d{2})\s*年\s*(1[0-2]|0?[1-9])\s*月")

FORMATTER_DIR = Path(__file__).resolve().parents[1] / "0_项目基本信息情况"
EXCEL_FORMATTING_UTILS_PATH = Path("/Users/wyattsun/Desktop/code/0_项目基本信息/excel_formatting_utils.py")


def _load_module_from_path(module_name: str, module_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(module_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_excel_formatting_utils = _load_module_from_path("excel_formatting_utils", EXCEL_FORMATTING_UTILS_PATH)
apply_common_excel_layout = _excel_formatting_utils.apply_common_excel_layout


def extract_month_from_filename(file_path: Path) -> str | None:
    """从文件名提取 YYYY-MM。"""
    match = FILE_MONTH_RE.search(file_path.name)
    if not match:
        return None
    year = int(match.group(1))
    month = int(match.group(2))
    return f"{year:04d}-{month:02d}"


def list_latest_file_per_month(input_dir: Path) -> dict[str, Path]:
    """扫描目录内所有xlsx，按月份保留最新修改时间文件。"""
    if not input_dir.exists():
        return {}

    month_to_file: dict[str, Path] = {}
    month_to_mtime: dict[str, float] = {}

    for file_path in input_dir.glob("*.xlsx"):
        if not file_path.is_file() or file_path.name.startswith("~$"):
            continue
        month = extract_month_from_filename(file_path)
        if not month:
            continue
        mtime = file_path.stat().st_mtime
        if month not in month_to_file or mtime > month_to_mtime[month]:
            month_to_file[month] = file_path
            month_to_mtime[month] = mtime

    return dict(sorted(month_to_file.items(), key=lambda item: item[0]))


def pick_amount_column(df: pd.DataFrame) -> str:
    candidates = ["费用合计(泰铢)", "费用合计", "本次费用合计"]
    for col in candidates:
        if col in df.columns:
            return col
    raise ValueError("未找到金额列，期望列之一: 费用合计(泰铢) / 费用合计 / 本次费用合计")


def build_monthly_total_table(month_to_file: dict[str, Path]) -> pd.DataFrame:
    records: list[dict[str, object]] = []

    for month, file_path in month_to_file.items():
        df = pd.read_excel(file_path, sheet_name=TARGET_SHEET)

        required_cols = ["项目编码", "项目名称"]
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"文件 {file_path.name} 的sheet {TARGET_SHEET} 缺少列: {col}")

        amount_col = pick_amount_column(df)

        temp = df[["项目编码", "项目名称", amount_col]].copy()
        temp["项目编码"] = temp["项目编码"].fillna("").astype(str).str.strip()
        temp["项目名称"] = temp["项目名称"].fillna("").astype(str).str.strip()
        temp[amount_col] = pd.to_numeric(temp[amount_col], errors="coerce").fillna(0.0)
        temp = temp[(temp["项目编码"] != "") | (temp["项目名称"] != "")]

        grouped = temp.groupby(["项目编码", "项目名称"], as_index=False).agg(总金额=(amount_col, "sum"))
        grouped["年月"] = month

        for _, row in grouped.iterrows():
            records.append(
                {
                    "项目编码": row["项目编码"],
                    "项目名称": row["项目名称"],
                    "年月": row["年月"],
                    "总金额": float(row["总金额"]),
                }
            )

    if not records:
        return pd.DataFrame(columns=["项目编码", "项目名称"])

    long_df = pd.DataFrame(records)

    total_pivot = long_df.pivot_table(
        index=["项目编码", "项目名称"],
        columns="年月",
        values="总金额",
        aggfunc="sum",
        fill_value=0.0,
    )

    total_pivot = total_pivot.reindex(columns=sorted(total_pivot.columns))
    total_pivot = total_pivot.reset_index()
    total_pivot = total_pivot.sort_values(["项目编码", "项目名称"], kind="stable").reset_index(drop=True)

    return total_pivot


def build_monthly_increase_table(total_df: pd.DataFrame) -> pd.DataFrame:
    if total_df.empty:
        return total_df.copy()

    month_cols = [col for col in total_df.columns if col not in ["项目编码", "项目名称"]]
    increase_df = total_df.copy()

    if not month_cols:
        return increase_df

    # 首月不计算增加量，用 "-" 标记。
    first_month = month_cols[0]
    increase_df[first_month] = "-"

    for idx in range(1, len(month_cols)):
        current_col = month_cols[idx]
        prev_col = month_cols[idx - 1]
        increase_df[current_col] = total_df[current_col] - total_df[prev_col]

    return increase_df


def main() -> None:
    month_to_file = list_latest_file_per_month(INPUT_DIR)
    if not month_to_file:
        print(f"未找到可处理文件：{INPUT_DIR}")
        return

    print("将使用以下月份文件：")
    for month, file_path in month_to_file.items():
        print(f"  {month}: {file_path.name}")

    total_df = build_monthly_total_table(month_to_file)
    increase_df = build_monthly_increase_table(total_df)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{OUTPUT_PREFIX}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # 排序第一：项目月度总金额
        total_df.to_excel(writer, sheet_name="项目月度总金额", index=False)
        increase_df.to_excel(writer, sheet_name="项目月度增加量", index=False)

        ws_total = writer.sheets["项目月度总金额"]
        ws_increase = writer.sheets["项目月度增加量"]
        apply_common_excel_layout(ws_total, numeric_start_col=3, freeze_panes="C2")
        apply_common_excel_layout(ws_increase, numeric_start_col=3, freeze_panes="C2")

    print("\n处理完成")
    print(f"输出文件: {output_path}")
    print(f"项目数: {len(total_df)}")
    print(f"月份数: {max(len(total_df.columns) - 2, 0)}")


if __name__ == "__main__":
    main()
