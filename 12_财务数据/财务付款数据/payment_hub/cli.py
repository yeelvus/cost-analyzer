# -*- coding: utf-8 -*-
"""
财务付款一体化 CLI

日常：pool / apply / run（无参数 → 交互菜单）
架构：payment_hub 编排 + learn_classify 分类引擎
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Optional

from .config import Paths
from .menu import interactive_menu, print_maint_help, print_quick_help
from .pipeline import (
    run_pipeline,
    step_apply_class_master,
    step_cashflow,
    step_export_class_master,
    step_export_review,
    step_export_taxonomy,
    step_health,
    step_import_taxonomy,
    step_learn,
    step_reclassify,
    step_std,
    step_summary,
)
from .pkg import ensure_code_root

# 旧名 → 规范名
_CMD_ALIASES = {
    "export-learn-pool": "pool",
    "apply-learn-pool": "apply",
    "all": "run",
    "month": "run",
    "help-guide": "help",
    "compact": "compact-memory",
}


def _paths_from_args(args: Any) -> Paths:
    p = Paths()
    if getattr(args, "raw_dir", None):
        p.raw_dir = Path(args.raw_dir)
    if getattr(args, "std_dir", None):
        p.std_dir = Path(args.std_dir)
    if getattr(args, "sum_xlsx", None):
        p.sum_xlsx = Path(args.sum_xlsx)
    if getattr(args, "sum_xlsx_z", None):
        p.sum_xlsx_z = Path(args.sum_xlsx_z)
    if getattr(args, "cashflow_html", None):
        p.cashflow_html = Path(args.cashflow_html)
    if getattr(args, "data_dir", None):
        p.classify_data = Path(args.data_dir)
        p.review_xlsx = p.classify_data / "待复核池.xlsx"
        p.rules_suggest_xlsx = p.classify_data / "规则建议.xlsx"
        p.class_master_xlsx = p.classify_data / "付款分类总表.xlsx"
    return p


def _add_path_args(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--raw-dir", help="A_财务付款_raw")
    sp.add_argument("--std-dir", help="B_财务付款_std")
    sp.add_argument("--sum-xlsx", help="汇总 Excel")
    sp.add_argument("--sum-xlsx-z", help="汇总副本 Z_report")
    sp.add_argument("--cashflow-html", help="现金流 HTML")
    sp.add_argument("--data-dir", help="learn_classify/data")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="财务付款一体化",
        description="财务付款：学习池 + 月结（日常 pool / apply / run）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="无参数且在终端中 → 交互菜单",
    )
    _add_path_args(p)
    p.add_argument("--menu", action="store_true", help="强制交互菜单")
    p.add_argument("--no-menu", action="store_true", help="无子命令只打印帮助")
    p.add_argument("--skip-std", action="store_true")
    p.add_argument("--skip-summary", action="store_true")
    p.add_argument("--skip-cashflow", action="store_true")
    p.add_argument("--reclassify", action="store_true", help="[run] 汇总前重刷分类")
    p.add_argument("--no-reclassify", action="store_true")
    p.add_argument("--no-reexport", action="store_true")
    p.add_argument("--only-confirmed", action="store_true")
    p.add_argument("--steps", nargs="+", default=None)
    p.add_argument("-i", "--input", default=None)
    p.add_argument("-o", "--output", default=None)

    sub = p.add_subparsers(dest="cmd", required=False)

    # 日常
    for name, aliases, help_ in [
        ("pool", ["export-learn-pool"], "导出学习池"),
        ("apply", ["apply-learn-pool"], "应用确认并泛化"),
        ("run", ["all", "month"], "月结全流程"),
    ]:
        c = sub.add_parser(name, aliases=aliases, help=help_)
        _add_path_args(c)
        if name == "apply":
            c.add_argument("--no-reclassify", action="store_true")
        if name == "run":
            c.add_argument("--skip-std", action="store_true")
            c.add_argument("--skip-summary", action="store_true")
            c.add_argument("--skip-cashflow", action="store_true")
            c.add_argument("--reclassify", action="store_true")
            c.add_argument("--steps", nargs="+", default=None)

    sub.add_parser("help", aliases=["help-guide"], help="用法说明")
    sub.add_parser("maint", help="维护子菜单 / 列表")
    c = sub.add_parser("show-paths", help="默认路径")
    _add_path_args(c)

    for name, help_ in [
        ("std", "仅标准化"),
        ("summary", "仅汇总"),
        ("cashflow", "仅现金流"),
        ("reclassify", "重刷 B_std 分类"),
        ("health", "分类健康度"),
        ("dropdowns", "B_std 加下拉"),
        ("cleanup-sheets", "删废弃 sheet"),
        ("export-class", "导出分类总表（旧）"),
        ("export-taxonomy", "导出科目体系"),
        ("import-taxonomy", "导入科目体系"),
    ]:
        c = sub.add_parser(name, help=help_)
        _add_path_args(c)
        if name in {"export-class", "export-taxonomy"}:
            c.add_argument("-o", "--output", default=None)
        if name == "import-taxonomy":
            c.add_argument("-i", "--input", default=None)

    c = sub.add_parser("apply-class", help="应用分类总表（旧）")
    _add_path_args(c)
    c.add_argument("-i", "--input", default=None)
    c.add_argument("--no-reclassify", action="store_true")
    c.add_argument("--no-reexport", action="store_true")
    c.add_argument("--only-confirmed", action="store_true")

    c = sub.add_parser("review", help="[旧] 导出待复核池")
    _add_path_args(c)
    c.add_argument("-o", "--output", default=None)
    c.add_argument("--conf", type=float, default=0.55)
    c.add_argument("--max-rows", type=int, default=None)

    c = sub.add_parser("learn", help="[旧] 从纠正表学习")
    _add_path_args(c)
    c.add_argument("-i", "--input", default=None)

    c = sub.add_parser("suggest-rules", help="规则建议")
    _add_path_args(c)
    c.add_argument("-o", "--output", default=None)
    c.add_argument("--min-support", type=int, default=3)

    c = sub.add_parser("approve-rules", help="合并 approved 规则")
    _add_path_args(c)
    c.add_argument("-i", "--input", required=True)

    c = sub.add_parser("migrate", help="taxonomy 迁移")
    _add_path_args(c)
    c.add_argument("--map", required=True)
    c.add_argument("--no-backup", action="store_true")
    c.add_argument("--with-reclassify", action="store_true")

    c = sub.add_parser("classify", help="单条摘要调试")
    _add_path_args(c)
    c.add_argument("--subject", default="")
    c.add_argument("--summary", required=True)
    c.add_argument("--vendor", default="")

    c = sub.add_parser("compact-memory", aliases=["compact"], help="压缩 memory/corrections")
    _add_path_args(c)
    c.add_argument("--dry-run", action="store_true")
    c.add_argument("--no-backup", action="store_true")
    c.add_argument("--memory-only", action="store_true")
    c.add_argument("--corrections-only", action="store_true")
    c.add_argument("--corrections-keep-last", type=int, default=None)

    return p


def _normalize_cmd(cmd: Optional[str]) -> Optional[str]:
    if not cmd:
        return None
    return _CMD_ALIASES.get(cmd, cmd)


def _cmd_help_guide() -> int:
    print_quick_help()
    guide = Path(__file__).resolve().parent.parent / "使用说明_财务付款一体化.md"
    if guide.exists():
        print()
        print(f"完整说明: {guide}")
    return 0


def _dispatch(cmd: str, paths: Paths, args: Any) -> int:
    ensure_code_root()

    if cmd == "show-paths":
        for k, v in paths.__dict__.items():
            print(f"{k:22} {v}")
        return 0

    if cmd == "run":
        run_pipeline(
            paths,
            steps=getattr(args, "steps", None),
            skip_std=bool(getattr(args, "skip_std", False)),
            skip_summary=bool(getattr(args, "skip_summary", False)),
            skip_cashflow=bool(getattr(args, "skip_cashflow", False)),
            reclassify_before_summary=bool(getattr(args, "reclassify", False)),
        )
        return 0

    if cmd == "std":
        step_std(paths)
        return 0
    if cmd == "summary":
        step_summary(paths)
        return 0
    if cmd == "cashflow":
        step_cashflow(paths)
        return 0
    if cmd == "reclassify":
        step_reclassify(paths)
        return 0
    if cmd == "health":
        step_health(paths)
        return 0

    if cmd == "pool":
        from .learn_pool import export_learn_pool

        stats = export_learn_pool(paths)
        print(f"✅ pool: {stats}")
        print(f"   打开: {paths.learn_pool_all_xlsx}")
        print("   改完后: apply")
        return 0

    if cmd == "apply":
        from .learn_pool import apply_learn_pool

        stats = apply_learn_pool(
            paths,
            reclassify=not bool(getattr(args, "no_reclassify", False)),
        )
        print(f"✅ apply: {stats}")
        return 0

    if cmd == "review":
        if getattr(args, "output", None):
            paths.review_xlsx = Path(args.output)
        step_export_review(
            paths,
            conf=getattr(args, "conf", 0.55),
            max_rows=getattr(args, "max_rows", None),
        )
        return 0

    if cmd == "learn":
        step_learn(paths, Path(args.input) if getattr(args, "input", None) else None)
        return 0

    if cmd == "export-class":
        step_export_class_master(
            paths, Path(args.output) if getattr(args, "output", None) else None
        )
        return 0

    if cmd == "apply-class":
        step_apply_class_master(
            paths,
            Path(args.input) if getattr(args, "input", None) else None,
            reclassify=not bool(getattr(args, "no_reclassify", False)),
            reexport=not bool(getattr(args, "no_reexport", False)),
            only_confirmed=bool(getattr(args, "only_confirmed", False)),
        )
        return 0

    if cmd == "export-taxonomy":
        step_export_taxonomy(
            paths, Path(args.output) if getattr(args, "output", None) else None
        )
        return 0

    if cmd == "import-taxonomy":
        step_import_taxonomy(
            paths, Path(args.input) if getattr(args, "input", None) else None
        )
        return 0

    if cmd == "dropdowns":
        from .category_dropdowns import apply_category_dropdowns_to_std_dir

        print(f"✅ 下拉: {apply_category_dropdowns_to_std_dir(paths.std_dir, data_dir=paths.classify_data)}")
        return 0

    if cmd == "cleanup-sheets":
        from .category_summary import remove_deprecated_sheets_in_std_dir

        print(f"✅ 清理: {remove_deprecated_sheets_in_std_dir(paths.std_dir)}")
        return 0

    if cmd == "suggest-rules":
        from learn_classify.learn import export_rule_suggestions

        out = Path(args.output) if getattr(args, "output", None) else paths.rules_suggest_xlsx
        export_rule_suggestions(
            paths.classify_data, out, min_support=getattr(args, "min_support", 3)
        )
        print(f"✅ 规则建议: {out}")
        return 0

    if cmd == "approve-rules":
        from learn_classify.learn import approve_rule_suggestions

        n = approve_rule_suggestions(paths.classify_data, Path(args.input))
        print(f"✅ 合并规则 {n} 条")
        return 0

    if cmd == "migrate":
        from learn_classify.migrate import run_migration

        report = run_migration(
            paths.classify_data,
            Path(args.map),
            do_backup=not args.no_backup,
            reclassify_std=paths.std_dir if args.with_reclassify else None,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0

    if cmd == "classify":
        from learn_classify.engine import ClassifyEngine

        r = ClassifyEngine.from_data_dir(paths.classify_data).classify(
            args.subject, args.summary, args.vendor
        )
        print(
            json.dumps(
                {
                    "一级": r.l1_name,
                    "二级": r.l2_name,
                    "三级": r.l3_name,
                    "层": r.layer,
                    "置信度": r.confidence,
                    "原因": r.reason,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    if cmd == "compact-memory":
        from learn_classify.store import compact_data_dir

        only_mem = bool(getattr(args, "memory_only", False))
        only_corr = bool(getattr(args, "corrections_only", False))
        do_memory = only_mem or not only_corr
        do_corr = only_corr or not only_mem
        if only_mem and only_corr:
            do_memory = do_corr = True
        report = compact_data_dir(
            paths.classify_data,
            memory=do_memory,
            corrections=do_corr,
            backup=not bool(getattr(args, "no_backup", False)),
            corrections_keep_last=getattr(args, "corrections_keep_last", None),
            dry_run=bool(getattr(args, "dry_run", False)),
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0

    raise ValueError(f"未知命令: {cmd}")


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    cmd = _normalize_cmd(args.cmd)

    if not cmd:
        if getattr(args, "menu", False) or (
            not getattr(args, "no_menu", False)
            and sys.stdin.isatty()
            and sys.stdout.isatty()
        ):
            return interactive_menu(main)
        print_quick_help()
        return 0

    if cmd == "help":
        return _cmd_help_guide()

    if cmd == "maint":
        if sys.stdin.isatty() and sys.stdout.isatty():
            from .menu import interactive_maint_menu

            return interactive_maint_menu(main)
        print_maint_help()
        return 0

    paths = _paths_from_args(args)
    try:
        return _dispatch(cmd, paths, args)
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
