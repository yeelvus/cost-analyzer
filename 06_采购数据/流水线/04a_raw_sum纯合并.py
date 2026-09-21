#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
C_采购部采购台账_std → D_采购部采购台账_raw_sum/采购部采购台账_raw_sum.xlsx

纯合并：不改内容、不跑规格/单位标准化。
与「04b_std_sum汇总」（D_raw_sum → E_sum/std_sum 对比结构）分离，需单独手动运行。

用法：
  python3 流水线/04a_raw_sum纯合并.py
  python3 流水线/04a_raw_sum纯合并.py --input DIR --output-dir DIR
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

_PKG_ROOT = Path(__file__).resolve().parents[1]
_CODE_ROOT = Path(__file__).resolve().parents[2]
for _p in (str(_PKG_ROOT), str(_CODE_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from _采购路径 import RAW_SUM_DIR, RAW_SUM_FILE_NAME, STD_DIR


def _load_step4():
    """复用 04b 中的 merge / 格式化工具。"""
    step4 = Path(__file__).resolve().parent / "04b_std_sum汇总.py"
    spec = importlib.util.spec_from_file_location("proc_step4_merge", str(step4))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载: {step4}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main(argv: list[str] | None = None) -> int:
    mod = _load_step4()

    parser = argparse.ArgumentParser(
        description="C_std 纯合并 → D_raw_sum/采购部采购台账_raw_sum.xlsx（无内容标准化）"
    )
    parser.add_argument("--input", default="", help=f"输入目录，默认 STD_DIR={STD_DIR}")
    parser.add_argument(
        "--output-dir",
        default="",
        help=f"输出目录，默认 RAW_SUM_DIR={RAW_SUM_DIR}",
    )
    parser.add_argument(
        "--output-name",
        default=RAW_SUM_FILE_NAME,
        help=f"输出文件名，默认 {RAW_SUM_FILE_NAME}",
    )
    args = parser.parse_args(argv)

    in_dir = Path(args.input).expanduser() if args.input else STD_DIR
    out_dir = Path(args.output_dir).expanduser() if args.output_dir else RAW_SUM_DIR
    if not in_dir.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    print("raw_sum 纯合并（无内容标准化）")
    print(f"- 输入 C_std: {in_dir}")
    print(f"- 输出:       {out_dir / args.output_name}")

    formal_wb, _ = mod.merge_standardized_excels(
        in_dir,
        use_content_standardization=False,
        include_compare=False,
    )
    if formal_wb is None:
        print("未产生有效数据")
        return 1

    out_path = out_dir / args.output_name
    mod.apply_unified_workbook_format(formal_wb)
    formal_wb.save(out_path)
    print(f"已输出: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
