#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
采购台账一键流水线（推荐入口）

默认只跑日常步骤：
  4 汇总 + 内容标准化 → sum

步骤 1（重命名）、2（月度去重）、3（格式标准化→std）按需**手动单独运行**，
不进默认全流程。

目录：
  流水线/     1 重命名 | 2 去重 | 3 std（均手动）| 4 汇总（默认）| 5 分类
  规格标准化/ 物品规格规则（由步骤4 调用）
  单位转换/   单位换算（由步骤4 调用）
  分析看板/   钢筋价格 / 桩成本
  工具/       旁路脚本

用法：
  python3 0_run_采购全流程.py                 # 默认仅 4
  python3 0_run_采购全流程.py --with-5
  python3 0_run_采购全流程.py --with-analysis
  python3 0_run_采购全流程.py --only 3        # 手动跑格式标准化
  python3 0_run_采购全流程.py --only 2        # 手动跑去重
  python3 0_run_采购全流程.py --with-prep     # 显式 1→2→3→4
  python3 0_run_采购全流程.py --show-paths
  python3 0_run_采购全流程.py --tree
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
_CODE = ROOT.parent
if str(_CODE) not in sys.path:
    sys.path.insert(0, str(_CODE))

from _采购路径 import (  # noqa: E402
    ANALYSIS_BATCH,
    ANALYSIS_CODE_DIR,
    ANALYSIS_DIR,
    CATEGORY_SUMMARY_NAME,
    CONTENT_DIR,
    EXTRACT_DIR,
    MERGED_PLAIN_NAME,
    MERGED_STD_NAME,
    PACKAGE_ROOT,
    PIPELINE_DIR,
    RAW_DIR,
    RAW_SUM_DIR,
    SPEC_DIR,
    STD_DIR,
    SUM_DIR,
    TOOLS_DIR,
    UNIT_DIR,
)

# 全部步骤定义；默认日常只跑 4
ALL_STEPS: list[tuple[str, str, Path]] = [
    ("1", "重命名月度文件（手动）", PIPELINE_DIR / "1_rename_excel_month_filenames.py"),
    ("2", "月度去重提取（手动）", PIPELINE_DIR / "2_采购台账月度去重.py"),
    ("3", "Excel格式标准化→std（手动）", PIPELINE_DIR / "3_采购部excel格式标准化.py"),
    ("4", "汇总+内容标准化→sum", PIPELINE_DIR / "4_采购部数据汇总.py"),
    ("5", "项目分类汇总", PIPELINE_DIR / "5_项目采购分类汇总.py"),
]
# 默认日常流水线（不含 1/2/3）
DEFAULT_STEP_IDS = ("4",)
# 仅手动的步骤（准备数据，不进默认）
MANUAL_STEP_IDS = ("1", "2", "3")


def _run(script: Path, extra: list[str] | None = None) -> int:
    if not script.is_file():
        print(f"❌ 脚本不存在: {script}")
        return 2
    cmd = [sys.executable, str(script)] + (extra or [])
    print(f"\n>>> {' '.join(cmd)}")
    t0 = time.time()
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        [str(_CODE), str(ROOT)]
        + ([env["PYTHONPATH"]] if env.get("PYTHONPATH") else [])
    )
    r = subprocess.run(cmd, cwd=str(ROOT), env=env)
    print(f"    耗时 {time.time() - t0:.1f}s  exit={r.returncode}")
    return int(r.returncode)


def _print_tree() -> None:
    print("代码结构")
    print(f"  {PACKAGE_ROOT.name}/")
    for label, path in [
        ("流水线/", PIPELINE_DIR),
        ("规格标准化/", SPEC_DIR),
        ("单位转换/", UNIT_DIR),
        ("内容标准化/", CONTENT_DIR),
        ("分析看板/", ANALYSIS_CODE_DIR),
        ("工具/", TOOLS_DIR),
    ]:
        print(f"  ├── {label}")
        if path.is_dir():
            for f in sorted(path.glob("*.py")):
                print(f"  │     {f.name}")
            for d in sorted(
                p
                for p in path.iterdir()
                if p.is_dir() and not p.name.startswith((".", "__"))
            ):
                print(f"  │     {d.name}/")
    print("\n数据目录")
    for name, path in [
        ("RAW", RAW_DIR),
        ("EXTRACT", EXTRACT_DIR),
        ("STD", STD_DIR),
        ("RAW_SUM", RAW_SUM_DIR),
        ("SUM", SUM_DIR),
        ("ANALYSIS_OUT", ANALYSIS_DIR),
    ]:
        flag = "✓" if path.exists() else "·"
        print(f"  {flag} {name:12} {path}")
    print(f"\n  plain_merged → {RAW_SUM_DIR / MERGED_PLAIN_NAME}")
    print(f"  std_merged   → {SUM_DIR / MERGED_STD_NAME}")
    print(f"  分类汇总     → {SUM_DIR / CATEGORY_SUMMARY_NAME}")
    print("\n运行约定")
    print("  默认全流程: 仅 4 汇总（不跑 1 重命名、2 去重、3 格式标准化）")
    print("  手动步骤:   1 / 2 / 3 → python3 0_run_采购全流程.py --only 1|2|3")
    print("              或: python3 流水线/1_….py / 2_….py / 3_….py")
    print("  一次准备全跑: --with-prep  → 1→2→3→4")


def _resolve_order(args) -> list[tuple[str, str, Path]]:
    by_id = {s[0]: s for s in ALL_STEPS}

    if args.only:
        sid = str(args.only)
        if sid not in by_id:
            raise SystemExit(f"未知步骤: {sid}（可选 1-5）")
        return [by_id[sid]]

    prep = bool(getattr(args, "with_prep", False) or getattr(args, "with_1_2", False))

    # 显式准备链：1→2→3→4（+可选5）
    if prep:
        ids = ["1", "2", "3", "4"]
    else:
        ids = list(DEFAULT_STEP_IDS)

    if args.with_5 and "5" not in ids:
        ids.append("5")

    # --from：仅当用户显式改了默认值时裁剪；--with-prep 默认从 1 起不裁
    start = str(getattr(args, "from_step", None) or "4")
    user_set_from = start != "4"  # 默认 argparse 为 4
    if user_set_from:
        if start in MANUAL_STEP_IDS and not prep:
            raise SystemExit(
                f"步骤 {start} 为手动步骤，默认不跑。"
                f"请用: --only {start}  或  --with-prep"
            )
        if start in ids:
            ids = ids[ids.index(start) :]
        elif start in by_id:
            all_ids = [s[0] for s in ALL_STEPS if s[0] != "5" or args.with_5]
            if start in all_ids:
                ids = all_ids[all_ids.index(start) :]
                if not prep:
                    ids = [i for i in ids if i in DEFAULT_STEP_IDS or i == "5"]

    return [by_id[i] for i in ids if i in by_id]


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="采购台账流水线（默认仅 4；1/2/3 需手动）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s                      # 日常：仅 4 汇总
  %(prog)s --only 3             # 只跑格式标准化（手动）
  %(prog)s --only 2             # 只跑月度去重（手动）
  %(prog)s --only 1             # 只跑重命名（手动）
  %(prog)s --with-prep          # 完整准备：1→2→3→4
  %(prog)s --no-prompt          # 若跑到步骤3：不弹项目匹配窗
        """,
    )
    p.add_argument(
        "--from",
        dest="from_step",
        default="4",
        help="从步骤号开始（默认 4；1/2/3 需 --only 或 --with-prep）",
    )
    p.add_argument("--only", default=None, help="只跑某一步 1-5")
    p.add_argument(
        "--with-prep",
        action="store_true",
        help="把手动准备步骤 1+2+3 一并纳入，再跑 4",
    )
    p.add_argument(
        "--with-1-2",
        action="store_true",
        help="同 --with-prep（兼容旧参数，现含步骤3）",
    )
    p.add_argument("--with-5", action="store_true", help="包含步骤5分类汇总")
    p.add_argument("--with-analysis", action="store_true", help="跑钢筋/桩分析看板")
    p.add_argument("--show-paths", action="store_true")
    p.add_argument("--tree", action="store_true", help="打印代码/数据目录树")
    p.add_argument(
        "--no-prompt",
        action="store_true",
        help="步骤3不弹项目匹配窗（新名保留原名，适合无人值守）",
    )
    args = p.parse_args(argv)

    if args.tree or args.show_paths:
        _print_tree()
        if args.show_paths:
            print(f"\n  std_merged → {SUM_DIR / MERGED_STD_NAME}")
        return 0

    os.environ["PYTHONPATH"] = os.pathsep.join(
        [str(_CODE), str(ROOT)]
        + ([os.environ["PYTHONPATH"]] if os.environ.get("PYTHONPATH") else [])
    )

    try:
        order = _resolve_order(args)
    except SystemExit as e:
        print(f"❌ {e}")
        return 2

    if not order:
        print("❌ 无步骤可执行")
        return 2

    print("=" * 60)
    print("采购台账流水线")
    print(f"  步骤: {' → '.join(s[0] + s[1] for s in order)}")
    if not args.only and not args.with_prep and not args.with_1_2:
        print("  （1 重命名 / 2 去重 / 3 格式标准化 为手动步骤，未纳入本次运行）")
    print(f"  EXTRACT: {EXTRACT_DIR}")
    print(f"  STD:     {STD_DIR}")
    print(f"  SUM:     {SUM_DIR}")
    print("=" * 60)

    t_all = time.time()
    for sid, title, script in order:
        print(f"\n{'=' * 60}\n  步骤 {sid}: {title}\n{'=' * 60}")
        extra: list[str] = []
        if sid == "3" and args.no_prompt:
            extra.append("--no-prompt")
        code = _run(script, extra or None)
        if code != 0:
            print(f"❌ 步骤 {sid} 失败，中止")
            return code

    if args.with_analysis:
        print(f"\n{'=' * 60}\n  分析看板\n{'=' * 60}")
        code = _run(ANALYSIS_BATCH)
        if code != 0:
            return code

    print(f"\n✅ 完成 总耗时 {time.time() - t_all:.1f}s")
    print(f"   汇总目录: {SUM_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
