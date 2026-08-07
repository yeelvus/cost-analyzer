# -*- coding: utf-8 -*-
"""
合同汇总 · 全量合并输出

前置：
  先运行「5.1_分包合同台账处理/00_全量合并汇总.py」
  （raw→std + 对比透视 + 付款汇总，刷新分包台账，供金额侧对比列使用）

再将原先两份 Excel 的全部 Sheet 合并到一个文件：
  - 明细侧（原「总包合同+分包合同汇总.xlsx」）
      总包合同 / 总包合同_不分类汇总 / 分包合同 / 分包合同_不分类汇总
  - 金额侧（原「合同管理_总包分包汇总.xlsx」）
      1_总包分包汇总(全部) / 2_人民币项目(CNY) / 3_泰铢项目(THB) / 4_混合币种项目

默认输出：{合同根目录}/合同管理_总包分包全量汇总.xlsx

用法：
  python 合同汇总_全量合并输出.py
  python 合同汇总_全量合并输出.py --skip-subpackage   # 跳过 5.1 台账全量合并
  python 合同汇总_全量合并输出.py --skip-std          # 台账跳过步骤1，仅用已有 std
"""

from __future__ import annotations

import argparse
import importlib.util
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

# 5.1 分包合同台账 · 全量合并入口
_SUBPACKAGE_MERGE_SCRIPT = (
    Path(__file__).resolve().parents[2]
    / "5.1_分包合同台账处理"
    / "00_全量合并汇总.py"
)


def run_subpackage_ledger_merge(*, skip_std: bool = False) -> Path:
    """先跑 5.1 分包台账 1+2+4 全量合并，刷新 std 与汇总。"""
    script = _SUBPACKAGE_MERGE_SCRIPT
    if not script.is_file():
        raise FileNotFoundError(f"未找到分包台账全量合并脚本: {script}")

    print("=" * 80)
    print("前置步骤 · 分包合同台账全量合并")
    print("=" * 80)
    print(f"脚本: {script}")

    spec = importlib.util.spec_from_file_location(
        "subpackage_ledger_merge_00",
        script,
    )
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {script}")

    mod = importlib.util.module_from_spec(spec)
    # 模块名以数字开头，用 file location 加载
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)

    if not hasattr(mod, "export_merged"):
        raise AttributeError(f"{script.name} 缺少 export_merged()")

    out = mod.export_merged(skip_std=skip_std)
    print(f"\n✅ 分包台账全量合并完成: {out}")
    return Path(out)


def export_merged(
    output_path: str | None = None,
    root_dir: str | None = None,
    *,
    run_subpackage: bool = True,
    skip_std: bool = False,
) -> str:
    # ---- 0) 分包台账全量合并（刷新 std，供金额对比列使用）----
    if run_subpackage:
        print("\n>>> [0/3] 运行 5.1 分包合同台账全量合并 …")
        run_subpackage_ledger_merge(skip_std=skip_std)
    else:
        print("\n>>> [0/3] 跳过分包台账全量合并（--skip-subpackage）")

    root_dir = root_dir or get_contract_root()
    if not os.path.isdir(root_dir):
        raise FileNotFoundError(f"合同根目录无效：{root_dir}")

    output_path = output_path or os.path.join(root_dir, MERGED_OUTPUT_NAME)
    print("\n" + "=" * 80)
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
    print("✅ 合同汇总合并完成")
    print("=" * 80)
    print(f"文件: {output_path}")
    print(f"共 {len(detail_names) + len(amount_names)} 个 Sheet：")
    for i, name in enumerate(detail_names + amount_names, 1):
        print(f"  {i}. {name}")
    if projects_data:
        from amount_summary import COL_GENERAL, COL_SUBPACKAGE

        df = pd.DataFrame(projects_data)
        print(
            f"\n金额侧概览: 项目 {len(df)} 个 | "
            f"总包 {df[COL_GENERAL].sum():,.4f} 泰铢 | "
            f"分包 {df[COL_SUBPACKAGE].sum():,.4f} 泰铢"
        )
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="合同汇总全量合并（默认先跑 5.1 分包台账全量合并）",
    )
    parser.add_argument(
        "--skip-subpackage",
        action="store_true",
        help="跳过 5.1 分包合同台账全量合并，直接做合同汇总",
    )
    parser.add_argument(
        "--skip-std",
        action="store_true",
        help="台账步骤跳过 raw→std，直接用已有 2_分包合同台账_std",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="合同汇总输出 Excel 路径",
    )
    parser.add_argument(
        "--root",
        default=None,
        help="合同根目录（默认 summary_config 中配置）",
    )
    args = parser.parse_args()
    export_merged(
        output_path=args.output,
        root_dir=args.root,
        run_subpackage=not args.skip_subpackage,
        skip_std=args.skip_std,
    )


if __name__ == "__main__":
    main()
