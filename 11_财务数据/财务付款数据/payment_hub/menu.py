# -*- coding: utf-8 -*-
"""交互式菜单（无参数运行时进入）。"""

from __future__ import annotations

import sys
from typing import Callable, List, Optional, Sequence, Tuple

from .config import Paths

# (键, 名称, 说明, 传给 dispatch 的 argv；None=特殊)
MenuItem = Tuple[str, str, str, Optional[List[str]]]

# 主菜单：只跑月结。科目对应关系在「摘要科目分类.py」。
MENU_MAIN: Sequence[MenuItem] = (
    ("1", "run", "月结全流程：raw_std → recognize → classify → summary → std → cashflow", ["run"]),
    ("2", "reclassify", "只填入科目：按当前对应重写 A2，并重建 B_std（不改记忆）", ["reclassify"]),
    ("3", "maint", "分步出账（raw-std / recognize / classify / summary / std…）", None),
    ("4", "show-paths", "查看默认数据路径", ["show-paths"]),
    ("h", "help", "打印命令说明", ["help"]),
    ("0", "quit", "退出", None),
)

# 维护 / 分步出账。改科目记忆不在这里。
MENU_MAINT: Sequence[MenuItem] = (
    ("1", "raw-std", "raw 标准化合并（按文件夹→一个Excel）", ["raw-std"]),
    ("2", "recognize", "摘要识别：项目提示写回 A1，不另存业务摘要", ["recognize"]),
    ("3", "classify", "把业务摘要填成一/二/三级科目（写入 A2，不改记忆）", ["classify"]),
    ("4", "summary", "汇总（只读 A2_classified）", ["summary"]),
    ("5", "std", "拆分 B_std（只读 A2_classified）", ["std"]),
    ("6", "cashflow", "仅现金流 HTML", ["cashflow"]),
    ("7", "amount-mismatch", "应付≠预扣+实付 不平明细", ["amount-mismatch"]),
    ("8", "dropdowns", "B_std 加分类下拉", ["dropdowns"]),
    ("0", "back", "返回主菜单", None),
)

_DIRECT = frozenset(
    {
        "run",
        "reclassify",
        "show-paths",
        "help",
        "raw-std",
        "raw_std",
        "recognize",
        "classify",
        "std",
        "summary",
        "cashflow",
        "health",
        "amount-mismatch",
        "dropdowns",
        "maint",
    }
)


def print_quick_help() -> None:
    print(
        """
财务付款月结 · 日常命令
────────────────────────────────────────
  python 财务付款一体化.py                # 月结窗口
  python 财务付款一体化.py run            # 月结全流程
  python 财务付款一体化.py reclassify     # 只把科目填进 A2 / B_std（不改记忆）

摘要对应一二三级科目（另一个程序，不改月结文件）:
  python 摘要科目分类.py                  # 填科目 / 写入记忆
  python 摘要科目分类.py fill -i 明细.xlsx
  python 摘要科目分类.py learn -i 已填科目.xlsx

菜单:
  [1] 月结全流程  raw_std → recognize → classify → summary → std → cashflow
  [2] 只填入科目  重写 A2，并重建 B_std
  [3] 分步出账    raw-std / recognize / classify / summary / std …

数据:
  A_raw → A1_raw_std → A2_classified → Sum / B_std
  分类这一步只读科目记忆
""".strip()
    )


def print_maint_help() -> None:
    print(
        """
维护 / 分步出账
────────────────────────────────────────
  raw-std / recognize / classify / summary / std / cashflow
  amount-mismatch / dropdowns

classify 只把科目填进 A2，不改科目记忆。
改对应关系请运行: python 摘要科目分类.py
""".strip()
    )


def _print_banner(paths: Optional[Paths] = None) -> None:
    print()
    print("=" * 56)
    print("  财务付款月结")
    print("=" * 56)
    if paths is not None:
        try:
            print(f"  A1_raw_std: {paths.raw_std_dir}")
            print(f"  A2_classified: {paths.classified_dir}")
            print(f"  B_std: {paths.std_dir}/{{组名}}/…")
            print(f"  Sum: {paths.sum_dir}/{{组名}}/…")
        except Exception:
            pass
    print()
    print("  科目对应关系: python 摘要科目分类.py")
    print("  [1] 月结全流程")
    print("  [2] 只填入科目（不改记忆）")
    print("  [3] 分步出账")
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


def _maybe_prompt_project(argv: List[str]) -> List[str]:
    if argv and argv[0] in {"run", "std", "raw-std"} and sys.stdin.isatty():
        ans = _read_choice("是否弹窗补录未匹配项目名？[y/N]: ")
        if ans in {"y", "yes", "是", "1"}:
            return argv + ["--prompt-project"]
    return argv


def interactive_maint_menu(dispatch: Callable[[Optional[list]], int]) -> int:
    last = 0
    while True:
        print()
        print("=" * 56)
        print("  维护 / 分步出账")
        print("=" * 56)
        print()
        _print_items(MENU_MAINT)
        choice = _read_choice("维护菜单，请选择: ")
        if choice in {"0", "q", "b", "back", "返回"}:
            return last
        if choice in {"h", "help"}:
            print_maint_help()
            _pause()
            continue

        argv = None
        for key, name, _d, a in MENU_MAINT:
            if choice in {key, name}:
                argv = list(a) if a is not None else None
                break
        if argv is None and choice in _DIRECT:
            argv = [choice]
        if argv is None:
            print(f"  无效选择: {choice!r}")
            _pause()
            continue
        argv = _maybe_prompt_project(argv)
        last = _run(dispatch, argv)
        _pause()


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

        if choice in {"3", "m", "maint", "维护"}:
            last = interactive_maint_menu(dispatch)
            continue

        argv: Optional[List[str]] = None
        for key, name, _d, a in MENU_MAIN:
            if choice in {key, name}:
                argv = list(a) if a is not None else None
                break
        if argv is not None:
            argv = _maybe_prompt_project(argv)
        if argv is None and choice in _DIRECT:
            argv = [choice]
            argv = _maybe_prompt_project(argv)
        if argv is None:
            if choice in {"learn", "学习", "pool", "apply", "learn-ledger", "确认学习", "账簿"}:
                print("  科目对应关系在另一个程序: python 摘要科目分类.py")
            else:
                print(f"  无效选择: {choice!r}，请输入 1–4 / h / 0")
            _pause()
            continue

        last = _run(dispatch, argv)
        _pause()
