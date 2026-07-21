# -*- coding: utf-8 -*-
"""
合同汇总 · 全量合并输出

将原先两份 Excel 的全部 Sheet 合并到一个文件：
  - 明细侧（原「总包合同+分包合同汇总.xlsx」）
      总包合同 / 总包合同_不分类汇总 / 分包合同 / 分包合同_不分类汇总
  - 金额侧（原「合同管理_总包分包汇总.xlsx」）
      1_总包分包汇总(全部) / 2_人民币项目(CNY) / 3_泰铢项目(THB) / 4_混合币种项目

默认输出：{合同根目录}/合同管理_总包分包全量汇总.xlsx
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

# 保证可从本目录导入
_THIS_DIR = Path(__file__).resolve().parent
if str(_THIS_DIR) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR))

from amount_summary import (  # noqa: E402
    build_amount_dataframes,
    format_amount_sheets,
    scan_all_projects,
    write_amount_sheets,
)
from detail_summary import (  # noqa: E402
    build_detail_dataframes,
    format_detail_sheets,
    scan_contract_details,
    write_detail_sheets,
)
from summary_config import (  # noqa: E402
    MERGED_OUTPUT_NAME,
    get_contract_root,
)


def export_merged(
    output_path: str | None = None,
    root_dir: str | None = None,
) -> str:
    root_dir = root_dir or get_contract_root()
    if not os.path.isdir(root_dir):
        raise FileNotFoundError(f"合同根目录无效：{root_dir}")

    output_path = output_path or os.path.join(root_dir, MERGED_OUTPUT_NAME)
    print("=" * 80)
    print("合同汇总 · 全量合并输出")
    print("=" * 80)
    print(f"合同根目录: {root_dir}")
    print(f"输出文件  : {output_path}")

    # ---- 1) 明细扫描 ----
    print("\n>>> [1/3] 扫描合同明细 JSON …")
    all_data, settlement_results, _ = scan_contract_details(root_dir=root_dir)
    detail_sheets = build_detail_dataframes(all_data, settlement_results)

    # ---- 2) 金额扫描 ----
    print("\n>>> [2/3] 扫描项目总包/分包金额 …")
    projects_data = scan_all_projects(base_dir=root_dir)
    amount_sheets = build_amount_dataframes(projects_data) if projects_data else {}

    if not detail_sheets and not amount_sheets:
        raise SystemExit("未扫描到任何合同数据，已中止。")

    # ---- 3) 写入同一 Excel ----
    print("\n>>> [3/3] 写入合并 Excel …")
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except PermissionError as e:
            raise PermissionError(f"文件被占用，请先关闭：{output_path}") from e

    detail_names: list[str] = []
    amount_names: list[str] = []

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if detail_sheets:
            detail_names = write_detail_sheets(writer, detail_sheets)
            print(f"  明细 Sheet: {detail_names}")
        if amount_sheets:
            amount_names = write_amount_sheets(writer, amount_sheets)
            print(f"  金额 Sheet: {amount_names}")

    # ---- 4) 分别套用格式 ----
    wb = load_workbook(output_path)
    if detail_names:
        format_detail_sheets(wb, detail_names)
    if amount_names:
        format_amount_sheets(wb, amount_names)
    wb.save(output_path)
    wb.close()

    print("\n" + "=" * 80)
    print("✅ 合并完成")
    print("=" * 80)
    print(f"文件: {output_path}")
    print(f"共 {len(detail_names) + len(amount_names)} 个 Sheet：")
    for i, name in enumerate(detail_names + amount_names, 1):
        print(f"  {i}. {name}")
    if projects_data:
        df = pd.DataFrame(projects_data)
        print(
            f"\n金额侧概览: 项目 {len(df)} 个 | "
            f"总包 {df['总包合同额(万泰铢)'].sum():,.2f} 万泰铢 | "
            f"分包 {df['分包合同额(万泰铢)'].sum():,.2f} 万泰铢"
        )
    return output_path


def main() -> None:
    export_merged()


if __name__ == "__main__":
    main()
