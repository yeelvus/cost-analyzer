# -*- coding: utf-8 -*-
"""
分包合同台账 · 1+2+4 全量合并

流程：
  1) 原始台账 raw → 标准化月度明细 std（步骤1）
  2) 月度对比 / 透视 / 协作体统计（步骤2）
  3) 最新月 合同额+已付款 项目汇总（步骤4）
  → 全部 Sheet 写入同一个 Excel

输出默认：
    {DATA}/5.1_分包合同/3_分包合同_sum/分包合同台账_汇总.xlsx

用法：
  python 00_全量合并汇总.py
  python 00_全量合并汇总.py --skip-std   # 跳过步骤1，直接用已有 std
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

_THIS = Path(__file__).resolve().parent
if str(_THIS) not in sys.path:
    sys.path.insert(0, str(_THIS))

from config import MERGED_OUTPUT_NAME, STD_DIR, SUM_DIR  # noqa: E402
from step1_std_ledger import process_raw_to_std  # noqa: E402
from step2_comparison import (  # noqa: E402
    build_comparison_sheets,
    format_comparison_sheets,
    scan_contract_files,
)
from step4_paid_summary import (  # noqa: E402
    build_paid_summary_frames,
    find_latest_excel,
    format_paid_sheets,
)


def export_merged(
    skip_std: bool = False,
    output_path: str | Path | None = None,
) -> Path:
    print("=" * 80)
    print("分包合同台账 · 1+2+4 全量合并汇总")
    print("=" * 80)

    # ---- 步骤1：标准化 ----
    if not skip_std:
        print("\n>>> [1/3] 原始台账 → 标准化明细 …")
        exported = process_raw_to_std()
        print(f"  本步导出 {len(exported)} 个标准化文件")
    else:
        print("\n>>> [1/3] 跳过标准化（使用已有 std 目录）")

    if not STD_DIR.exists() or not any(STD_DIR.glob("分包签订合同汇总_*.xlsx")):
        raise FileNotFoundError(
            f"标准化目录无有效文件：{STD_DIR}\n请先运行步骤1或去掉 --skip-std"
        )

    # ---- 步骤2：对比与透视 ----
    print("\n>>> [2/3] 逐月对比与项目透视 …")
    files_data = scan_contract_files(STD_DIR)
    if not files_data:
        raise SystemExit("std 目录下没有可解析的月度合同文件")
    comparison_sheets = build_comparison_sheets(files_data)

    # ---- 步骤4：付款汇总（基于最新月） ----
    print("\n>>> [3/3] 最新月合同额与已付款汇总 …")
    latest = find_latest_excel(STD_DIR)
    print(f"  使用最新标准化文件: {latest.name}")
    paid_sheets = build_paid_summary_frames(latest)

    # ---- 合并写入 ----
    SUM_DIR.mkdir(parents=True, exist_ok=True)
    output_path = Path(output_path or (SUM_DIR / MERGED_OUTPUT_NAME))
    if output_path.exists():
        try:
            os.remove(output_path)
        except PermissionError as e:
            raise PermissionError(f"文件被占用，请先关闭：{output_path}") from e

    print(f"\n写入合并 Excel: {output_path}")
    all_sheets: dict[str, pd.DataFrame] = {}
    all_sheets.update(comparison_sheets)
    all_sheets.update(paid_sheets)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            safe = str(name)[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            print(f"  + {safe}  ({len(df)} 行)")

    wb = load_workbook(output_path)
    format_comparison_sheets(wb)
    format_paid_sheets(wb)
    wb.save(output_path)
    wb.close()

    print("\n" + "=" * 80)
    print("✅ 合并完成")
    print("=" * 80)
    print(f"文件: {output_path}")
    print("Sheet 列表：")
    for i, name in enumerate(all_sheets.keys(), 1):
        print(f"  {i}. {name}")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="分包合同台账 1+2+4 全量合并")
    parser.add_argument(
        "--skip-std",
        action="store_true",
        help="跳过步骤1标准化，直接基于 2_分包合同台账_std",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出 Excel 路径（默认 sum 目录下 分包合同台账_汇总.xlsx）",
    )
    args = parser.parse_args()
    export_merged(skip_std=args.skip_std, output_path=args.output)


if __name__ == "__main__":
    main()
