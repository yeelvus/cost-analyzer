# -*- coding: utf-8 -*-
"""
python -m machine [command]

  merge-raw | summarize | split-sum | split-ledger | report | all | paths
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _ensure_path():
    here = Path(__file__).resolve().parent
    code = here.parent
    root = code.parent
    for p in (str(code), str(root), str(here)):
        if p not in sys.path:
            sys.path.insert(0, p)


def cmd_merge_raw(args):
    from machine.merge_raw import run_merge_raw
    run_merge_raw()


def cmd_summarize(args):
    from machine.summarize import run_summarize_batch, main as summarize_one
    if args.file:
        summarize_one(args.file, output_basename_override=args.basename)
    else:
        run_summarize_batch()


def cmd_split_sum(args):
    from machine.split import split_summary_by_project
    raise SystemExit(split_summary_by_project(args.source, args.target_dir, args.cols))


def cmd_split_ledger(args):
    from machine.split import split_ledger_by_project
    split_ledger_by_project(args.input, args.output_dir)


def cmd_report(args):
    from machine.report import run_report
    run_report(args.file, args.output_dir)


def cmd_all(args):
    print("=" * 60 + "\n全流程 1/3 merge-raw\n" + "=" * 60)
    cmd_merge_raw(args)
    print("\n" + "=" * 60 + "\n全流程 2/3 summarize\n" + "=" * 60)
    from machine.summarize import run_summarize_batch
    run_summarize_batch()
    print("\n" + "=" * 60 + "\n全流程 3/3 report\n" + "=" * 60)
    from machine.report import run_report
    try:
        run_report()
    except FileNotFoundError as e:
        print(f"⚠️ 报告跳过: {e}")
    print("\n" + "=" * 60 + "\n全流程完成\n" + "=" * 60)


def cmd_paths(args):
    from machine import config as c
    rows = [
        ("CODE_ROOT", c.CODE_ROOT),
        ("DATA_ROOT", c.DATA_ROOT),
        ("MACHINE_ROOT", c.MACHINE_ROOT),
        ("RAW_DIR", c.RAW_DIR),
        ("STD_DIR", c.STD_DIR),
        ("SUM_DIR", c.SUM_DIR),
        ("REPORT_DIR", c.REPORT_DIR),
        ("MACHINE_MAPPING_JSON", c.MACHINE_MAPPING_JSON),
        ("PROJECT_MATCH_DB", c.PROJECT_MATCH_DB),
        ("RUN_MODE", c.RUN_MODE),
    ]
    print("机具费路径:\n")
    for k, v in rows:
        mark = ""
        if isinstance(v, Path):
            mark = " ✓" if v.exists() else " ✗"
        print(f"  {k:22s} {v}{mark}")


def build_parser():
    p = argparse.ArgumentParser(prog="machine", description="机具费处理（重构版）")
    sub = p.add_subparsers(dest="command")

    sub.add_parser("merge-raw", help="raw→机具费_raw合并→按项目拆分").set_defaults(func=cmd_merge_raw)

    s = sub.add_parser("summarize", help="std → sum 汇总")
    s.add_argument("--file", default=None, help="指定 std 文件或留空批量")
    s.add_argument("--basename", default=None)
    s.set_defaults(func=cmd_summarize)

    s = sub.add_parser("split-sum", help="汇总按项目拆 workbook")
    s.add_argument("--source", default=None)
    s.add_argument("--target_dir", default=None)
    s.add_argument("--cols", nargs="*", default=None)
    s.set_defaults(func=cmd_split_sum)

    s = sub.add_parser("split-ledger", help="机具费_raw合并.xlsx 按项目拆（一般 merge-raw 已自动执行）")
    s.add_argument("--input", default=None)
    s.add_argument("--output-dir", default=None)
    s.set_defaults(func=cmd_split_ledger)

    s = sub.add_parser("report", help="HTML/Excel 分析报告")
    s.add_argument("--file", default=None)
    s.add_argument("--output-dir", default=None)
    s.set_defaults(func=cmd_report)

    sub.add_parser("all", help="merge-raw + summarize + report").set_defaults(func=cmd_all)
    sub.add_parser("paths", help="打印路径").set_defaults(func=cmd_paths)
    return p


def main(argv=None):
    _ensure_path()
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.command:
        args = parser.parse_args(["all"])
    args.func(args)


if __name__ == "__main__":
    main()
