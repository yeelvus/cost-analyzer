# -*- coding: utf-8 -*-
"""交互式菜单（无参数运行时进入）。"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

from .config import Paths

# (键, 名称, 说明, 传给 dispatch 的 argv；None=特殊)
MenuItem = Tuple[str, str, str, Optional[List[str]]]

# 主菜单
MENU_MAIN: Sequence[MenuItem] = (
    ("1", "learn", "确认学习组（导出学习池 / 应用确认 / 健康度…）", None),
    ("2", "run", "月结全流程：raw_std → classify → summary → std → cashflow", ["run"]),
    ("3", "maint", "维护 / 分步出账（raw-std / classify / summary / std…）", None),
    ("4", "show-paths", "查看默认数据路径", ["show-paths"]),
    ("h", "help", "打印命令说明", ["help"]),
    ("0", "quit", "退出", None),
)

# 确认学习组
MENU_LEARN: Sequence[MenuItem] = (
    ("1", "pool", "导出学习池 → Std/财务付款/D_财务付款_学习池", ["pool"]),
    ("2", "apply", "应用确认 → 写入记忆 + 抬升未确认 + 刷分类", ["apply"]),
    ("3", "learn-ledger", "账簿已填科目 → 写入记忆（不改账簿）", ["learn-ledger"]),
    ("4", "reclassify", "仅重刷 B_std / 后续可再 classify", ["reclassify"]),
    ("5", "health", "分类健康度（命中层 / 待复核率）", ["health"]),
    ("6", "export-class", "导出付款分类总表（旧路径）", ["export-class"]),
    ("7", "apply-class", "应用分类总表学习（旧路径）", ["apply-class"]),
    ("8", "review", "导出待复核池", ["review"]),
    ("9", "compact-memory", "压缩 memory / corrections", ["compact-memory"]),
    ("0", "back", "返回主菜单", None),
)

# 维护 / 分步出账
MENU_MAINT: Sequence[MenuItem] = (
    ("1", "raw-std", "raw 标准化合并（按文件夹→一个Excel）", ["raw-std"]),
    ("2", "classify", "摘要→一/二/三级科目（A1→A2）", ["classify"]),
    ("3", "summary", "汇总（优先 A2_classified）", ["summary"]),
    ("4", "std", "拆分 B_std（优先 A2_classified）", ["std"]),
    ("5", "cashflow", "仅现金流 HTML", ["cashflow"]),
    ("6", "amount-mismatch", "应付≠预扣+实付 不平明细", ["amount-mismatch"]),
    ("7", "export-taxonomy", "导出科目体系 Excel", ["export-taxonomy"]),
    ("8", "import-taxonomy", "从 Excel 写回 taxonomy", ["import-taxonomy"]),
    ("9", "dropdowns", "B_std 加分类下拉", ["dropdowns"]),
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
        "raw-std",
        "raw_std",
        "classify",
        "std",
        "summary",
        "cashflow",
        "compact-memory",
        "export-class",
        "apply-class",
        "health",
        "review",
        "amount-mismatch",
        "export-taxonomy",
        "import-taxonomy",
        "dropdowns",
        "learn",
        "learn-ledger",
        "ledger-learn",
        "maint",
    }
)


def print_quick_help() -> None:
    print(
        """
财务付款一体化 · 日常命令
────────────────────────────────────────
  python 财务付款一体化.py                # 交互式菜单（推荐）
  python 财务付款一体化.py pool           # ① 导出学习池（确认学习）
  python 财务付款一体化.py apply          # ② 应用确认并学习
  python 财务付款一体化.py learn-ledger   # 账簿已填科目 → 写入记忆
  python 财务付款一体化.py run            # ③ 月结全流程

菜单结构:
  [1] 确认学习组  pool / apply / reclassify / health / 分类总表…
  [2] 月结全流程  raw_std → classify → summary → std → cashflow
  [3] 维护分步    raw-std / classify / summary / std / cashflow…

快路径数据:
  A_raw → A1_raw_std → A2_classified → Sum / B_std

确认学习:
  学习池: 1(pool) 导出 → Excel 改科目或点「确认」→ 2(apply) 写记忆
  账簿:   改一/二/三级 → 3(learn-ledger) 写入记忆
""".strip()
    )


def print_maint_help() -> None:
    print(
        """
维护 / 分步出账
────────────────────────────────────────
  raw-std / classify / summary / std / cashflow
  amount-mismatch / export-taxonomy / import-taxonomy
  dropdowns

确认学习请回主菜单选 [1]
""".strip()
    )


def print_learn_help() -> None:
    print(
        """
确认学习组
────────────────────────────────────────
  pool           从明细导出学习池（02_Std/12_财务付款/D_财务付款_学习池）
  apply          读确认结果 → memory + 泛化 + 重分类
  learn-ledger   账簿已填一/二/三级 → 写入记忆（不改账簿）
  reclassify     仅按当前记忆/规则重刷 B_std
  health         看命中层与待复核率
  export-class   导出分类总表（旧）
  apply-class    应用分类总表（旧）
  review         导出待复核池
  compact-memory 压缩学习记忆文件

推荐顺序:
  学习池: pool → 人工确认 Excel → apply
  账簿:   改科目后选 3(learn-ledger)
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
            print(f"  A1_raw_std: {paths.raw_std_dir}")
            print(f"  A2_classified: {paths.classified_dir}")
            print(f"  B_std: {paths.std_dir}/{{组名}}/…")
            print(f"  Sum: {paths.sum_dir}/{{组名}}/…")
        except Exception:
            pass
    print()
    print("  [1] 确认学习  →  pool / apply …")
    print("  [2] 月结出账  →  run 全流程")
    print("  [3] 维护分步  →  raw-std / classify / summary / std …")
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


def _prompt_learn_ledger(argv: List[str]) -> List[str]:
    """交互确认账簿路径；可选仅预览。"""
    if not argv or argv[0] not in {"learn-ledger", "ledger-learn"}:
        return argv
    if "--ledger" in argv or not sys.stdin.isatty():
        return argv
    default = Paths().ledger_xlsx
    exists = "存在" if default.is_file() else "不存在"
    print()
    print(f"  默认账簿: {default}")
    print(f"  文件状态: {exists}")
    ans = _read_choice("回车用默认，或粘贴账簿路径: ")
    ledger = Path(ans.strip().strip('"').strip("'")).expanduser() if ans else default
    extra = ["--ledger", str(ledger)]
    dry = _read_choice("仅预览、不写入记忆？[y/N]: ")
    if dry in {"y", "yes", "是", "1"}:
        extra.append("--dry-run")
    return argv + extra


def interactive_learn_menu(dispatch: Callable[[Optional[list]], int]) -> int:
    """确认学习组子菜单。"""
    last = 0
    while True:
        print()
        print("=" * 56)
        print("  确认学习组")
        print("=" * 56)
        print()
        print("  推荐: 1(pool) → Excel 改科目/点确认 → 2(apply)")
        print("        账簿改科目后 → 3(learn-ledger) 写入记忆")
        print()
        _print_items(MENU_LEARN)
        choice = _read_choice("确认学习，请选择: ")
        if choice in {"0", "q", "b", "back", "返回"}:
            return last
        if choice in {"h", "help"}:
            print_learn_help()
            _pause()
            continue

        argv = None
        if choice in {"账簿", "账簿学习", "写入记忆", "ledger"}:
            argv = ["learn-ledger"]
        else:
            for key, name, _d, a in MENU_LEARN:
                if choice in {key, name}:
                    argv = list(a) if a is not None else None
                    break
        if argv is None and choice in _DIRECT:
            argv = [choice]
        if argv is None:
            print(f"  无效选择: {choice!r}")
            _pause()
            continue

        argv = _prompt_learn_ledger(argv)
        last = _run(dispatch, argv)
        _pause()


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

        # 确认学习组
        if choice in {"1", "learn", "学习", "确认学习", "确认学习组"}:
            last = interactive_learn_menu(dispatch)
            continue

        # 维护分步
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
            print(f"  无效选择: {choice!r}，请输入 1–4 / h / 0")
            _pause()
            continue

        last = _run(dispatch, argv)
        _pause()
