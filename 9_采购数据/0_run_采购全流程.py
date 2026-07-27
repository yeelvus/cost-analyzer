#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
采购台账一键流水线（推荐入口）

默认日常（依赖已有 C_std）：
  04a raw_sum 纯合并 → 04b std_sum 汇总 → 06 按项目拆分

手动准备（不进默认）：
  01 重命名、02 月度去重、03 excel 格式标准化 → C_std

用法：
  python3 0_run_采购全流程.py                 # 默认 04a→04b→06
  python3 0_run_采购全流程.py --no-split      # 只要 04a→04b
  python3 0_run_采购全流程.py --only 06       # 只拆分
  python3 0_run_采购全流程.py --with-prep     # 01→02→03→04a→04b→06
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
    CONTENT_DIR,
    EXTRACT_DIR,
    PACKAGE_ROOT,
    PIPELINE_DIR,
    RAW_DIR,
    RAW_SUM_DIR,
    RAW_SUM_FILE_NAME,
    SPEC_DIR,
    STD_DIR,
    STD_SPLIT_DIR,
    STD_SUM_FILE_NAME,
    SUM_DIR,
    TOOLS_DIR,
    UNIT_DIR,
)

# id, 说明, 脚本, 是否默认日常
ALL_STEPS: list[tuple[str, str, Path, bool]] = [
    ("01", "重命名月度文件（手动）", PIPELINE_DIR / "01_重命名月度文件.py", False),
    ("02", "月度去重提取（手动）", PIPELINE_DIR / "02_月度去重提取.py", False),
    ("03", "Excel格式标准化→C_std（手动）", PIPELINE_DIR / "03_excel格式标准化.py", False),
    ("04a", "raw_sum 纯合并→D_raw_sum", PIPELINE_DIR / "04a_raw_sum纯合并.py", True),
    ("04b", "std_sum 汇总→E_sum", PIPELINE_DIR / "04b_std_sum汇总.py", True),
    ("05", "分类汇总刷新（可选）", PIPELINE_DIR / "05_分类汇总刷新.py", False),
    ("06", "按项目拆分std_sum→F_split", PIPELINE_DIR / "06_按项目拆分.py", True),
]

# 兼容旧步骤号：04 / 4 → 04b
_LEGACY_ID = {
    "4": "04b",
    "04": "04b",
    "4a": "04a",
    "4b": "04b",
}


def _norm_id(s: str) -> str:
    s = str(s).strip()
    return _LEGACY_ID.get(s, s)


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
                if f.name == "__init__.py":
                    continue
                print(f"  │     {f.name}")
            for d in sorted(
                p
                for p in path.iterdir()
                if p.is_dir() and not p.name.startswith((".", "__"))
            ):
                print(f"  │     {d.name}/")
    print("\n数据目录")
    for name, path in [
        ("A_raw", RAW_DIR),
        ("B_extract", EXTRACT_DIR),
        ("C_std", STD_DIR),
        ("D_raw_sum", RAW_SUM_DIR),
        ("E_sum", SUM_DIR),
        ("F_split", STD_SPLIT_DIR),
        ("ANALYSIS", ANALYSIS_DIR),
    ]:
        flag = "✓" if path.exists() else "·"
        print(f"  {flag} {name:12} {path}")
    print(f"\n  raw_sum  → {RAW_SUM_DIR / RAW_SUM_FILE_NAME}")
    print(f"  std_sum  → {SUM_DIR / STD_SUM_FILE_NAME}")
    print(f"  split    → {STD_SPLIT_DIR}/{{编码}}_{{名称}}_采购数据.xlsx")
    print("\n目录顺序: A_raw → B_extract → C_std → D_raw_sum → E_sum → F_split")
    print("默认日常: 04a → 04b → 06")
    print("手动准备: 01 / 02 / 03")


def _resolve_order(args) -> list[tuple[str, str, Path]]:
    by_id = {s[0]: (s[0], s[1], s[2]) for s in ALL_STEPS}

    if args.only:
        sid = _norm_id(args.only)
        if sid not in by_id:
            raise SystemExit(f"未知步骤: {args.only}（可选 01/02/03/04a/04b/04/05/06）")
        return [by_id[sid]]

    if args.with_prep:
        ids = ["01", "02", "03", "04a", "04b", "06"]
    else:
        # 默认日常
        ids = [s[0] for s in ALL_STEPS if s[3]]
        if args.no_split:
            ids = [i for i in ids if i != "06"]
        if args.no_raw_sum:
            ids = [i for i in ids if i != "04a"]

    if args.with_5 and "05" not in ids:
        # 插在 04b 之后、06 之前
        if "04b" in ids:
            i = ids.index("04b") + 1
            ids.insert(i, "05")
        elif "04a" in ids:
            i = ids.index("04a") + 1
            ids.insert(i, "05")
        else:
            ids.append("05")

    return [by_id[i] for i in ids if i in by_id]


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="采购台账流水线（默认 04a→04b→06；01/02/03 手动）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
步骤编号：
  01  重命名月度文件          （手动）
  02  月度去重提取            （手动）
  03  Excel 格式标准化→C_std （手动）
  04a C_std→D_raw_sum 纯合并 （默认）
  04b D_raw_sum→E_sum        （默认；旧号 04 仍可用）
  05  刷新 std_sum 分类 sheet （可选）
  06  按项目拆分→F_std_split （默认）

示例：
  %(prog)s                    # 04a → 04b → 06
  %(prog)s --no-split         # 04a → 04b
  %(prog)s --only 06          # 只拆分
  %(prog)s --with-prep        # 01→02→03→04a→04b→06
  %(prog)s --no-prompt        # 若跑到 03：不弹项目匹配窗
        """,
    )
    p.add_argument("--only", default=None, help="只跑某一步：01/02/03/04a/04b/04/05/06")
    p.add_argument(
        "--with-prep",
        action="store_true",
        help="含手动准备 01+02+03，再跑 04a→04b→06",
    )
    p.add_argument(
        "--with-1-2",
        action="store_true",
        help="同 --with-prep（兼容旧参数）",
    )
    p.add_argument("--with-5", action="store_true", help="在 04b 后额外刷新分类 sheet")
    p.add_argument(
        "--no-split",
        action="store_true",
        help="默认流程中不跑 06 按项目拆分",
    )
    p.add_argument(
        "--no-raw-sum",
        action="store_true",
        help="默认流程中不跑 04a（要求 D_raw_sum 已存在）",
    )
    p.add_argument("--with-analysis", action="store_true", help="跑钢筋/桩分析看板")
    p.add_argument("--show-paths", action="store_true")
    p.add_argument("--tree", action="store_true", help="打印代码/数据目录树")
    p.add_argument(
        "--no-prompt",
        action="store_true",
        help="步骤03不弹项目匹配窗",
    )
    # 兼容旧 --from
    p.add_argument("--from", dest="from_step", default=None, help=argparse.SUPPRESS)
    args = p.parse_args(argv)

    if args.with_1_2:
        args.with_prep = True

    if args.tree or args.show_paths:
        _print_tree()
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
    print(f"  步骤: {' → '.join(s[0] + ' ' + s[1].split('（')[0] for s in order)}")
    print(f"  C_std:   {STD_DIR}")
    print(f"  D_raw:   {RAW_SUM_DIR}")
    print(f"  E_sum:   {SUM_DIR}")
    print(f"  F_split: {STD_SPLIT_DIR}")
    print("=" * 60)

    t_all = time.time()
    for sid, title, script in order:
        print(f"\n{'=' * 60}\n  步骤 {sid}: {title}\n{'=' * 60}")
        extra: list[str] = []
        if sid == "03" and args.no_prompt:
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
