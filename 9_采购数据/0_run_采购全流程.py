#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
采购台账一键流水线（推荐入口）

目录：
  流水线/     1 重命名 → 2 去重 → 3 格式std → 4 汇总 → 5 分类
  规格标准化/ 物品规格规则（由步骤4 内容标准化调用）
  单位转换/   单位换算（由步骤4 调用）
  分析看板/   钢筋价格 / 桩成本
  工具/       旁路脚本

用法：
  python3 0_run_采购全流程.py
  python3 0_run_采购全流程.py --with-5 --with-analysis
  python3 0_run_采购全流程.py --from 3
  python3 0_run_采购全流程.py --only 4
  python3 0_run_采购全流程.py --show-paths
  python3 0_run_采购全流程.py --tree
"""
from __future__ import annotations

import argparse
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
    MERGED_STD_NAME,
    PACKAGE_ROOT,
    PIPELINE_DIR,
    RAW_DIR,
    SPEC_DIR,
    STD_DIR,
    SUM_DIR,
    TOOLS_DIR,
    UNIT_DIR,
)

STEPS: list[tuple[str, str, Path]] = [
    ("1", "重命名月度文件", PIPELINE_DIR / "1_rename_excel_month_filenames.py"),
    ("2", "月度去重提取", PIPELINE_DIR / "2_采购台账月度去重.py"),
    ("3", "Excel格式标准化→std", PIPELINE_DIR / "3_采购部excel格式标准化.py"),
    ("4", "汇总+内容标准化→sum", PIPELINE_DIR / "4_采购部数据汇总.py"),
    ("5", "项目分类汇总", PIPELINE_DIR / "5_项目采购分类汇总.py"),
]


def _run(script: Path, extra: list[str] | None = None) -> int:
    if not script.is_file():
        print(f"❌ 脚本不存在: {script}")
        return 2
    cmd = [sys.executable, str(script)] + (extra or [])
    print(f"\n>>> {' '.join(cmd)}")
    t0 = time.time()
    # 子进程 cwd=包根，保证相对导入与路径一致
    import os
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join([str(_CODE), str(ROOT)] + ([env["PYTHONPATH"]] if env.get("PYTHONPATH") else []))
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
            for d in sorted(p for p in path.iterdir() if p.is_dir() and not p.name.startswith((".", "__"))):
                print(f"  │     {d.name}/")
    print("\n数据目录")
    for name, path in [
        ("RAW", RAW_DIR),
        ("EXTRACT", EXTRACT_DIR),
        ("STD", STD_DIR),
        ("SUM", SUM_DIR),
        ("ANALYSIS_OUT", ANALYSIS_DIR),
    ]:
        flag = "✓" if path.exists() else "·"
        print(f"  {flag} {name:12} {path}")


def main(argv: list[str] | None = None) -> int:
    import os

    p = argparse.ArgumentParser(description="采购台账一键流水线")
    p.add_argument("--from", dest="from_step", default="1", help="从步骤号开始 1-5")
    p.add_argument("--only", default=None, help="只跑某一步 1-5")
    p.add_argument("--with-5", action="store_true", help="包含步骤5分类汇总")
    p.add_argument("--with-analysis", action="store_true", help="跑钢筋/桩分析看板")
    p.add_argument("--show-paths", action="store_true")
    p.add_argument("--tree", action="store_true", help="打印代码/数据目录树")
    p.add_argument("--skip-1", action="store_true", help="跳过重命名")
    args = p.parse_args(argv)

    if args.tree or args.show_paths:
        _print_tree()
        if args.show_paths:
            print(f"\n  std_merged → {SUM_DIR / MERGED_STD_NAME}")
        return 0

    # 确保子进程能 import settings / _采购路径
    os.environ["PYTHONPATH"] = os.pathsep.join(
        [str(_CODE), str(ROOT)] + ([os.environ["PYTHONPATH"]] if os.environ.get("PYTHONPATH") else [])
    )

    order = [s for s in STEPS if s[0] != "5" or args.with_5]
    if args.only:
        order = [s for s in order if s[0] == str(args.only)]
    else:
        start = str(args.from_step)
        started = False
        filtered = []
        for s in order:
            if s[0] == start:
                started = True
            if started:
                if args.skip_1 and s[0] == "1":
                    continue
                filtered.append(s)
        order = filtered

    print("=" * 60)
    print("采购台账全流程")
    print(f"  步骤: {' → '.join(s[0] + s[1] for s in order)}")
    print(f"  RAW: {RAW_DIR}")
    print(f"  SUM: {SUM_DIR}")
    print("=" * 60)

    t_all = time.time()
    for sid, title, script in order:
        print(f"\n{'=' * 60}\n  步骤 {sid}: {title}\n{'=' * 60}")
        code = _run(script)
        if code != 0:
            print(f"❌ 步骤 {sid} 失败，中止")
            return code

    if args.with_analysis:
        print(f"\n{'=' * 60}\n  分析看板\n{'=' * 60}")
        code = _run(ANALYSIS_BATCH)
        if code != 0:
            return code

    print(f"\n✅ 全流程完成 总耗时 {time.time() - t_all:.1f}s")
    print(f"   汇总目录: {SUM_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
