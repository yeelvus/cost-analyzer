# -*- coding: utf-8 -*-
"""
财务付款一体化 CLI

融合：
  3_ 原始→标准明细
  4_ 汇总分析 Excel
  5_ 项目现金流 HTML
  learn_classify 可持续分类学习
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import Paths
from .pipeline import (
    run_pipeline,
    step_cashflow,
    step_export_review,
    step_health,
    step_learn,
    step_reclassify,
    step_std,
    step_summary,
)


def _paths_from_args(args) -> Paths:
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
    return p


def _add_path_args(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--raw-dir", help="A_财务付款_raw")
    sp.add_argument("--std-dir", help="B_财务付款_std")
    sp.add_argument("--sum-xlsx", help="汇总 Excel（C_财务付款_sum）")
    sp.add_argument("--sum-xlsx-z", help="汇总 Excel 副本（Z_report）")
    sp.add_argument("--cashflow-html", help="现金流 HTML 输出路径")
    sp.add_argument("--data-dir", help="learn_classify/data 目录")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="payment_hub",
        description="财务付款一体化：明细标准化 + 汇总分析 + 现金流 HTML + 分类学习",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
常用示例:
  python 财务付款一体化.py                   # 直接跑全流程（默认）
  python 财务付款一体化.py --skip-std        # 跳过标准化，只汇总+HTML
  python 财务付款一体化.py --reclassify      # 全流程且汇总前重刷分类
  python 财务付款一体化.py std               # 仅原始→标准
  python 财务付款一体化.py summary           # 仅汇总
  python 财务付款一体化.py cashflow          # 仅 HTML
  python 财务付款一体化.py review            # 导出待复核
  python 财务付款一体化.py learn -i 待复核池.xlsx
  python 财务付款一体化.py help-guide        # 打印使用说明
        """,
    )
    # 无子命令时默认全流程；以下为全流程可用参数（写在主 parser 上）
    _add_path_args(p)
    p.add_argument("--skip-std", action="store_true", help="跳过标准化（已有 B_std）")
    p.add_argument("--skip-summary", action="store_true")
    p.add_argument("--skip-cashflow", action="store_true")
    p.add_argument(
        "--reclassify",
        action="store_true",
        help="汇总前用当前分类体系重刷 B_std",
    )
    p.add_argument(
        "--steps",
        nargs="+",
        default=None,
        help="自定义步骤序列: std reclassify summary cashflow review learn health",
    )

    sub = p.add_subparsers(dest="cmd", required=False)

    # ---- all（兼容旧写法）----
    c = sub.add_parser("all", help="全流程：std → summary → cashflow（与无子命令相同）")
    _add_path_args(c)
    c.add_argument("--skip-std", action="store_true", help="跳过标准化（已有 B_std）")
    c.add_argument("--skip-summary", action="store_true")
    c.add_argument("--skip-cashflow", action="store_true")
    c.add_argument(
        "--reclassify",
        action="store_true",
        help="汇总前用当前分类体系重刷 B_std",
    )
    c.add_argument(
        "--steps",
        nargs="+",
        default=None,
        help="自定义步骤序列: std reclassify summary cashflow review learn health",
    )

    # ---- single steps ----
    for name, help_ in [
        ("std", "仅 3_ 原始→标准明细"),
        ("summary", "仅 4_ 汇总分析 Excel"),
        ("cashflow", "仅 5_ 现金流 HTML"),
        ("reclassify", "重刷 B_std 一二级分类"),
        ("health", "分类健康度"),
    ]:
        c = sub.add_parser(name, help=help_)
        _add_path_args(c)

    c = sub.add_parser("review", help="导出待复核池 Excel")
    _add_path_args(c)
    c.add_argument("-o", "--output", default=None)
    c.add_argument("--conf", type=float, default=0.55)
    c.add_argument("--max-rows", type=int, default=None)

    c = sub.add_parser("learn", help="从纠正表学习（填正确二级后）")
    _add_path_args(c)
    c.add_argument("-i", "--input", default=None, help="纠正 Excel，默认待复核池")

    c = sub.add_parser("suggest-rules", help="从纠错生成规则建议")
    _add_path_args(c)
    c.add_argument("-o", "--output", default=None)
    c.add_argument("--min-support", type=int, default=3)

    c = sub.add_parser("approve-rules", help="合并 status=approved 的规则")
    _add_path_args(c)
    c.add_argument("-i", "--input", required=True)

    c = sub.add_parser("migrate", help="分类体系迁移/转移")
    _add_path_args(c)
    c.add_argument("--map", required=True, help="migration json")
    c.add_argument("--no-backup", action="store_true")
    c.add_argument(
        "--with-reclassify",
        action="store_true",
        help="迁移后自动重刷 B_std",
    )

    c = sub.add_parser("classify", help="单条摘要分类调试")
    _add_path_args(c)
    c.add_argument("--subject", default="")
    c.add_argument("--summary", required=True)
    c.add_argument("--vendor", default="")

    c = sub.add_parser("help-guide", help="打印完整使用说明")
    # no path args needed

    c = sub.add_parser("show-paths", help="显示当前默认路径")
    _add_path_args(c)

    c = sub.add_parser(
        "dropdowns",
        help="给 B_std 下已有 Excel 的一级/二级列批量加下拉（不重算数据）",
    )
    _add_path_args(c)

    return p


def cmd_help_guide() -> None:
    guide = Path(__file__).resolve().parent.parent / "使用说明_财务付款一体化.md"
    if guide.exists():
        print(guide.read_text(encoding="utf-8"))
    else:
        print("未找到 使用说明_财务付款一体化.md，请查看 payment_hub 与 learn_classify README。")


def _run_all(paths: Paths, args) -> int:
    run_pipeline(
        paths,
        steps=getattr(args, "steps", None),
        skip_std=bool(getattr(args, "skip_std", False)),
        skip_summary=bool(getattr(args, "skip_summary", False)),
        skip_cashflow=bool(getattr(args, "skip_cashflow", False)),
        reclassify_before_summary=bool(getattr(args, "reclassify", False)),
    )
    return 0


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # 无子命令：直接跑全流程
    if not args.cmd:
        paths = _paths_from_args(args)
        try:
            return _run_all(paths, args)
        except Exception as e:
            print(f"❌ {e}", file=sys.stderr)
            return 1

    if args.cmd == "help-guide":
        cmd_help_guide()
        return 0

    paths = _paths_from_args(args)

    try:
        if args.cmd == "show-paths":
            for k, v in paths.__dict__.items():
                print(f"{k:18} {v}")
            return 0

        if args.cmd == "all":
            return _run_all(paths, args)

        if args.cmd == "std":
            step_std(paths)
            return 0
        if args.cmd == "summary":
            step_summary(paths)
            return 0
        if args.cmd == "cashflow":
            step_cashflow(paths)
            return 0
        if args.cmd == "reclassify":
            step_reclassify(paths)
            return 0
        if args.cmd == "health":
            step_health(paths)
            return 0
        if args.cmd == "review":
            if args.output:
                paths.review_xlsx = Path(args.output)
            step_export_review(paths, conf=args.conf, max_rows=args.max_rows)
            return 0
        if args.cmd == "learn":
            step_learn(paths, Path(args.input) if args.input else None)
            return 0
        if args.cmd == "suggest-rules":
            from learn_classify.learn import export_rule_suggestions

            out = Path(args.output) if args.output else paths.rules_suggest_xlsx
            export_rule_suggestions(paths.classify_data, out, min_support=args.min_support)
            print(f"✅ 规则建议: {out}")
            return 0
        if args.cmd == "approve-rules":
            from learn_classify.learn import approve_rule_suggestions

            n = approve_rule_suggestions(paths.classify_data, Path(args.input))
            print(f"✅ 合并规则 {n} 条")
            return 0
        if args.cmd == "migrate":
            from learn_classify.migrate import run_migration

            report = run_migration(
                paths.classify_data,
                Path(args.map),
                do_backup=not args.no_backup,
                reclassify_std=paths.std_dir if args.with_reclassify else None,
            )
            import json

            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 0
        if args.cmd == "classify":
            from learn_classify.engine import ClassifyEngine
            import json

            eng = ClassifyEngine.from_data_dir(paths.classify_data)
            r = eng.classify(args.subject, args.summary, args.vendor)
            print(
                json.dumps(
                    {
                        "一级": r.l1_name,
                        "二级": r.l2_name,
                        "ID": r.l2_id,
                        "层": r.layer,
                        "置信度": r.confidence,
                        "原因": r.reason,
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 0

        if args.cmd == "dropdowns":
            from .category_dropdowns import apply_category_dropdowns_to_std_dir

            stats = apply_category_dropdowns_to_std_dir(
                paths.std_dir, data_dir=paths.classify_data
            )
            print(f"✅ 下拉批量完成: {stats}")
            return 0

        parser.error(f"未知命令: {args.cmd}")
        return 2
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
