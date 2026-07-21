# -*- coding: utf-8 -*-
"""
python -m machine [command]

  merge-raw      A_raw → B_raw_sum + C_raw_split
  summarize      C_raw_split → D_std_sum + E_std_split
  split-ledger   仅：B_raw_sum 合并 → C_raw_split
  split-sum      仅：D_std_sum → E_std_split
  report         D_std_sum → F_report
  all            全流程
  paths          打印路径
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
    from machine.split import split_std_summary_by_project
    raise SystemExit(split_std_summary_by_project(args.source, args.target_dir, args.cols))


def cmd_split_ledger(args):
    from machine.split import split_raw_merge_by_project
    split_raw_merge_by_project(args.input, args.output_dir)


def cmd_report(args):
    from machine.report import run_report
    run_report(args.file, args.output_dir)


def cmd_all(args):
    print("=" * 60 + "\n全流程 1/4 merge-raw（B_raw_sum + C_raw_split）\n" + "=" * 60)
    cmd_merge_raw(args)
    print("\n" + "=" * 60 + "\n全流程 2/4 summarize（D_std_sum）\n" + "=" * 60)
    from machine.summarize import run_summarize_batch
    run_summarize_batch()
    print("\n" + "=" * 60 + "\n全流程 3/4 确保 E_std_split\n" + "=" * 60)
    try:
        from machine.split import split_std_summary_by_project
        split_std_summary_by_project()
    except Exception as e:
        print(f"⚠️ 标准拆分: {e}")
    print("\n" + "=" * 60 + "\n全流程 4/4 report\n" + "=" * 60)
    from machine.report import run_report
    try:
        run_report()
    except FileNotFoundError as e:
        print(f"⚠️ 报告跳过: {e}")
    print("\n" + "=" * 60 + "\n全流程完成\n" + "=" * 60)


def cmd_paths(args):
    from machine import config as c
    from machine.config import ensure_data_dirs, migrate_legacy_dirs
    ensure_data_dirs()
    logs = migrate_legacy_dirs()
    if logs:
        print("已迁移旧目录文件:")
        for line in logs[:30]:
            print(" ", line)
    rows = [
        ("MACHINE_ROOT", c.MACHINE_ROOT),
        ("RAW_DIR (原始台账)", c.RAW_DIR),
        ("B RAW_SUM_DIR", c.RAW_SUM_DIR),
        ("C RAW_SPLIT_DIR", c.RAW_SPLIT_DIR),
        ("D STD_SUM_DIR", c.STD_SUM_DIR),
        ("E STD_SPLIT_DIR", c.STD_SPLIT_DIR),
        ("F REPORT_DIR", c.REPORT_DIR),
        ("RAW_MERGE_NAME", c.RAW_MERGE_NAME),
        ("FULL_SUMMARY_FILENAME", c.FULL_SUMMARY_FILENAME),
        ("RUN_MODE", c.RUN_MODE),
    ]
    print("机具费数据目录:\n")
    for k, v in rows:
        mark = ""
        if isinstance(v, Path):
            mark = " ✓" if v.exists() else " （将创建）"
        print(f"  {k:36s} {v}{mark}")


def build_parser():
    p = argparse.ArgumentParser(prog="machine", description="机具费处理（目录结构化）")
    sub = p.add_subparsers(dest="command")

    sub.add_parser("merge-raw", help="A_raw → B_raw_sum + C_raw_split").set_defaults(func=cmd_merge_raw)

    s = sub.add_parser("summarize", help="C_raw_split → D_std_sum (+E_std_split)")
    s.add_argument("--file", default=None)
    s.add_argument("--basename", default=None)
    s.set_defaults(func=cmd_summarize)

    s = sub.add_parser("split-sum", help="D_std_sum → E_std_split")
    s.add_argument("--source", default=None)
    s.add_argument("--target_dir", default=None)
    s.add_argument("--cols", nargs="*", default=None)
    s.set_defaults(func=cmd_split_sum)

    s = sub.add_parser("split-ledger", help="B 合并文件 → C_raw_split")
    s.add_argument("--input", default=None)
    s.add_argument("--output-dir", default=None)
    s.set_defaults(func=cmd_split_ledger)

    s = sub.add_parser("report", help="D_std_sum → F_report")
    s.add_argument("--file", default=None)
    s.add_argument("--output-dir", default=None)
    s.set_defaults(func=cmd_report)

    sub.add_parser("all", help="全流程").set_defaults(func=cmd_all)
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
