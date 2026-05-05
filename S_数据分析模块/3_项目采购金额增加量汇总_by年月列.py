"""
按「年月」列统计项目月度采购金额及增加量

数据源: B_采购部采购台账_sum/*.xlsx  →  合并数据 sheet
  - 每行包含 项目编码、项目名称、年月、含税金额 等列
  - 年月 格式为 YYYY-MM（来自签订日期）

输出: 两个 sheet
  - 采购总金额: 项目 × 月份  宽表
  - 增加量    : 本月与上月之差（第一个有效月份增加量为 0）

用法（直接运行，或命令行参数覆盖）:
  python 3_项目采购金额增加量汇总_by年月列.py
  python 3_项目采购金额增加量汇总_by年月列.py --start-period 2025-09 --latest-only
"""

from __future__ import annotations

import argparse
import importlib.util
import re
from pathlib import Path
from typing import Any

import pandas as pd


# ============= 路径配置 =============
SCRIPT_DIR = Path(__file__).resolve().parent

INPUT_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/B_采购部采购台账_sum")
DEFAULT_OUTPUT_PATH = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report/项目采购金额增加量汇总_by年月.xlsx")

EXCEL_FORMATTING_UTILS_PATH = Path("/Users/wyattsun/Desktop/code/0_项目基本信息/excel_formatting_utils.py")

DEFAULT_SHEET = "合并数据"
# ====================================


# ── 辅助模块加载 ──────────────────────────────────────────────

def _load_module_from_path(module_name: str, module_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(module_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


if EXCEL_FORMATTING_UTILS_PATH.is_file():
    _fmt = _load_module_from_path("excel_formatting_utils", EXCEL_FORMATTING_UTILS_PATH)
    apply_common_excel_layout = _fmt.apply_common_excel_layout
else:
    def apply_common_excel_layout(*args, **kwargs):
        return None
    print("警告: 未找到 excel_formatting_utils.py，已跳过输出样式设置。")


# ── 期间工具 ─────────────────────────────────────────────────

def period_key(period: str) -> tuple[int, int]:
    y, m = period.split("-")
    return int(y), int(m)


def build_month_label_map(periods: list[str]) -> dict[str, str]:
    return {p: f"{int(p.split('-')[0])}年{int(p.split('-')[1])}月" for p in periods}


def normalize_text(val: Any) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in {"", "nan", "none", "null"} else s


# ── 文件选取 ─────────────────────────────────────────────────

def find_sum_files(root: Path, latest_only: bool) -> list[Path]:
    """列出 B_sum 目录下的 std_merged.xlsx，按文件名排序。"""
    files = sorted(
        [p for p in root.glob("*std_merged.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.name,
    )
    if not files:
        raise FileNotFoundError(f"未找到 *std_merged.xlsx 文件，目录: {root}")
    if latest_only:
        print(f"  仅读取最新文件: {files[-1].name}")
        return [files[-1]]
    return files


# ── 单文件读取 ────────────────────────────────────────────────

def load_one_file(path: Path, sheet_name: str) -> tuple[pd.DataFrame | None, str]:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        return None, f"读取失败: {exc}"

    df.columns = [str(c).strip() for c in df.columns]

    # 定位必要列（兼容轻微列名差异）
    def find_col(aliases: list[str]) -> str | None:
        for a in aliases:
            if a in df.columns:
                return a
        return None

    project_code_col = find_col(["项目编码", "编码", "项目编号"])
    project_name_col = find_col(["项目名称", "项目", "工程名称"])
    ym_col           = find_col(["年月"])
    amount_col       = find_col(["含税金额", "总金额", "价税合计", "采购金额"])

    missing = [name for name, col in [
        ("年月", ym_col), ("含税金额", amount_col)
    ] if col is None]
    if missing:
        return None, f"缺少必要列: {', '.join(missing)}"

    if project_code_col is None and project_name_col is None:
        return None, "缺少必要列: 项目编码/项目名称至少要有一个"

    code_series = df[project_code_col].map(normalize_text) if project_code_col else pd.Series([""] * len(df))
    name_series = df[project_name_col].map(normalize_text) if project_name_col else pd.Series([""] * len(df))
    ym_series   = df[ym_col].map(normalize_text)

    work = pd.DataFrame({
        "项目编码":   code_series,
        "项目名称":   name_series,
        "年月":       ym_series,
        "含税金额":   pd.to_numeric(df[amount_col], errors="coerce"),
    })

    # 过滤掉无效行
    work = work[(work["项目编码"] != "") | (work["项目名称"] != "")]
    work = work[work["年月"].str.fullmatch(r"(?:19|20)\d{2}-(?:0[1-9]|1[0-2])", na=False)]
    work = work[work["含税金额"].notna()]

    if work.empty:
        return None, "无有效数据行（年月格式或金额为空）"

    work["来源文件"] = path.name
    return work[["年月", "项目编码", "项目名称", "含税金额", "来源文件"]], ""


# ── 主逻辑 ────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按「年月」列统计项目月度采购金额（读取最新汇总文件）")
    parser.add_argument("--input-root",  default=str(INPUT_ROOT),         help="B_sum 目录")
    parser.add_argument("--sheet",       default=DEFAULT_SHEET,            help="工作表名，默认 合并数据")
    parser.add_argument("--output",      default=str(DEFAULT_OUTPUT_PATH), help="输出 Excel 路径")
    parser.add_argument("--all-files",   action="store_true",              help="读取全部文件而非只读最新文件")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    input_root  = Path(args.input_root).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    all_files   = args.all_files  # 默认 False，只读最新文件

    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    files = find_sum_files(input_root, latest_only=not all_files)
    print(f"\n输入目录: {input_root}")
    print(f"命中文件数: {len(files)}")

    frames: list[pd.DataFrame] = []
    skipped: list[tuple[str, str]] = []

    for f in files:
        part, err = load_one_file(f, args.sheet)
        if part is None:
            skipped.append((f.name, err))
            print(f"  跳过 {f.name}: {err}")
        else:
            frames.append(part)
            print(f"  读取 {f.name}: {len(part)} 行")

    if not frames:
        raise RuntimeError("所有文件均无法解析，请检查 sheet 名称与列名")

    all_rows = pd.concat(frames, ignore_index=True)

    # 按项目 + 年月 汇总含税金额
    grouped = (
        all_rows
        .groupby(["项目编码", "项目名称", "年月"], as_index=False)
        .agg(含税金额=("含税金额", "sum"))
    )

    valid_periods = sorted(grouped["年月"].unique().tolist(), key=period_key)

    pivot = grouped.pivot_table(
        index=["项目编码", "项目名称"],
        columns="年月",
        values="含税金额",
        aggfunc="sum",
        fill_value=0.0,
    ).reindex(columns=valid_periods, fill_value=0.0)

    label_map   = build_month_label_map(valid_periods)
    month_cols  = [label_map[p] for p in valid_periods]
    amount_wide = pivot.rename(columns=label_map).reset_index()

    # 按最新月金额降序排列（编码非空优先）
    sort_cols = [label_map[p] for p in reversed(valid_periods)]
    amount_wide["__编码为空"] = amount_wide["项目编码"].astype(str).str.strip().eq("")
    amount_wide = amount_wide.sort_values(["__编码为空", *sort_cols], ascending=[True] + [False] * len(sort_cols))
    amount_wide = amount_wide.drop(columns=["__编码为空"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        amount_wide.to_excel(writer, sheet_name="采购总金额", index=False)
        apply_common_excel_layout(writer.book["采购总金额"], numeric_start_col=3)

    print("\n处理完成")
    print(f"期间范围: {valid_periods[0]} ~ {valid_periods[-1]}")
    print(f"项目数量: {pivot.shape[0]}")
    print(f"输出列:   项目编码, 项目名称, {', '.join(month_cols)}")
    print(f"输出文件: {output_path}")

    if skipped:
        print("\n以下文件被跳过：")
        for name, reason in skipped:
            print(f"  - {name}: {reason}")


if __name__ == "__main__":
    main()
