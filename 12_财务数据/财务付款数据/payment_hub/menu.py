# -*- coding: utf-8 -*-
"""交互式菜单（无参数运行时进入）。"""

from __future__ import annotations

import sys
from typing import Callable, List, Optional, Sequence, Tuple

from .config import Paths

# (键, 名称, 说明, 传给 dispatch 的 argv；None=特殊)
MenuItem = Tuple[str, str, str, Optional[List[str]]]

MENU_MAIN: Sequence[MenuItem] = (
    ("1", "pool", "导出学习池 D_财务付款_学习池", ["pool"]),
    ("2", "apply", "应用确认 → 记忆 + 抬升未确认 + 刷 B_std", ["apply"]),
    ("3", "run", "月结全流程：std → 汇总 → 现金流 HTML", ["run"]),
    ("4", "reclassify", "仅重刷 B_std 分类", ["reclassify"]),
    ("5", "show-paths", "查看默认数据路径", ["show-paths"]),
    ("6", "maint", "维护 / 进阶子菜单", None),
    ("h", "help", "打印命令说明", ["help"]),
    ("0", "quit", "退出", None),
)

MENU_MAINT: Sequence[MenuItem] = (
    ("1", "std", "仅 raw → 标准明细", ["std"]),
    ("2", "summary", "仅汇总 Excel", ["summary"]),
    ("3", "cashflow", "仅现金流 HTML", ["cashflow"]),
    ("4", "compact-memory", "压缩 memory.jsonl", ["compact-memory"]),
    ("5", "export-class", "导出付款分类总表（旧）", ["export-class"]),
    ("6", "apply-class", "应用分类总表学习（旧）", ["apply-class"]),
    ("7", "export-taxonomy", "导出科目体系 Excel", ["export-taxonomy"]),
    ("8", "import-taxonomy", "从 Excel 写回 taxonomy", ["import-taxonomy"]),
    ("9", "health", "分类健康度", ["health"]),
    ("0", "back", "返回主菜单", None),
)

_DIRECT = frozenset(
    {
        "pool",
        "apply",
        "run",
        "reclassify",
        "show-paths",
        "help",
        "std",
        "summary",
        "cashflow",
        "compact-memory",
        "export-class",
        "apply-class",
        "health",
    }
)


def print_quick_help() -> None:
    print(
        """
财务付款一体化 · 日常只需 3 个命令
────────────────────────────────────────
  python 财务付款一体化.py           # 交互式菜单（推荐）
  python 财务付款一体化.py pool      # ① 导出学习池
  python 财务付款一体化.py apply     # ② 应用确认
  python 财务付款一体化.py run       # ③ 月结全流程

改分类: 1(pool) → Excel 确认 → 2(apply)
出报表: 3(run)
其它:   help / maint / show-paths
""".strip()
    )


def print_maint_help() -> None:
    print(
        """
维护 / 进阶（菜单选 6，或不常用命令）
────────────────────────────────────────
  reclassify / std / summary / cashflow
  compact-memory / health
  export-class / apply-class（旧分类总表路径）
  export-taxonomy / import-taxonomy
  classify --summary X   单条调试
  suggest-rules / migrate / cleanup-sheets

日常请用: pool / apply / run
""".strip()
    )


def _print_banner(paths: Optional[Paths] = None) -> None:
    print()
    print("=" * 56)
    print("  财务付款一体化 · 交互菜单")
    print("=" * 56)
    if paths is not None:
        try:
            print(f"  学习池: {paths.learn_pool_root}")
            print(f"  标准明细: {paths.std_dir}")
        except Exception:
            pass
    print()
    print("  改分类: 1(pool) → Excel 确认 → 2(apply)")
    print("  出报表: 3(run)")
    print()


def _print_items(items: Sequence[MenuItem]) -> None:
    for key, name, desc, _ in items:
        if key == "0":
            print(f"  [{key}] {desc}")
        else:
            print(f"  [{key}] {name:<16} {desc}")
    print()


def _read_choice(prompt: str = "请选择编号后回车: ") -> str:
    try:
        return input(prompt).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print("\n已取消。")
        return "0"


def _pause() -> None:
    if not sys.stdin.isatty():
        return
    try:
        input("按回车返回菜单…")
    except (EOFError, KeyboardInterrupt):
        print()


def _run(dispatch: Callable[[Optional[list]], int], argv: List[str]) -> int:
    print()
    print("-" * 56)
    print(f"  ▶ 执行: {' '.join(argv)}")
    print("-" * 56)
    print()
    try:
        code = dispatch(argv)
    except SystemExit as e:
        code = int(e.code) if e.code is not None else 0
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        code = 1
    print()
    print(f"{'✅' if code == 0 else '⚠️'} 结束（退出码 {code}）")
    return code


def interactive_menu(dispatch: Callable[[Optional[list]], int]) -> int:
    """dispatch 通常为 payment_hub.cli.main。"""
    paths = Paths()
    last = 0
    while True:
        _print_banner(paths)
        _print_items(MENU_MAIN)
        choice = _read_choice()

        if choice in {"0", "q", "quit", "exit", "退出"}:
            print("再见。")
            return last

        if choice in {"6", "m", "maint"}:
            last = interactive_maint_menu(dispatch)
            continue

        argv: Optional[List[str]] = None
        for key, name, _d, a in MENU_MAIN:
            if choice in {key, name}:
                argv = a
                break
        if argv is None and choice in _DIRECT:
            argv = [choice]
        if argv is None:
            print(f"  无效选择: {choice!r}，请输入 1–6 / h / 0")
            _pause()
            continue

        last = _run(dispatch, argv)
        _pause()


def interactive_maint_menu(dispatch: Callable[[Optional[list]], int]) -> int:
    last = 0
    while True:
        print()
        print("=" * 56)
        print("  维护 / 进阶")
        print("=" * 56)
        print()
        _print_items(MENU_MAINT)
        choice = _read_choice("维护菜单，请选择: ")
        if choice in {"0", "q", "b", "back", "返回"}:
            return last
        argv = None
        for key, name, _d, a in MENU_MAINT:
            if choice in {key, name}:
                argv = a
                break
        if argv is None:
            print(f"  无效选择: {choice!r}")
            _pause()
            continue
        last = _run(dispatch, argv)
        _pause()
