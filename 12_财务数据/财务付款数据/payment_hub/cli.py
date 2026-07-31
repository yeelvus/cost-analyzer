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
        p.class_master_xlsx = p.classify_data / "付款分类总表.xlsx"
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
日常主路径:
  python 财务付款一体化.py                   # 默认 = apply-class（学分类总表叶子变更）
  python 财务付款一体化.py all               # 月结全流程: std → summary → cashflow
  python 财务付款一体化.py month             # 同 all（语义别名）

学习池 D_财务付款_学习池:
  python 财务付款一体化.py export-learn-pool # 导出 00_汇总 + 01_按项目 + 02_按月份
  python 财务付款一体化.py apply-learn-pool  # 确认/手改 → 记忆 → 已确认 → 剥离 → 刷 B_std
  # 推荐在 00_汇总/全部摘要_学习池.xlsx 快速改全量（含非THCC），改完后必须 apply

分步 / 维护:
  python 财务付款一体化.py std | summary | cashflow | reclassify
  python 财务付款一体化.py compact-memory
  python 财务付款一体化.py help-guide
        """,
    )
    # 无子命令时默认 apply-class（不是 all）；全流程请显式 all/month
    _add_path_args(p)
    p.add_argument(
        "-i",
        "--input",
        default=None,
        help="分类总表路径（默认自动找 data 或 C_sum 下的 付款分类总表.xlsx）",
    )
    p.add_argument(
        "--no-reclassify",
        action="store_true",
        help="默认 apply-class 时：只学习，不重刷 B_std",
    )
    p.add_argument(
        "--no-reexport",
        action="store_true",
        help="默认 apply-class 时：应用后不重新导出总表",
    )
    p.add_argument(
        "--only-confirmed",
        action="store_true",
        help="仅学习「确认状态=确认」的行（默认关闭：表内有效二级全部学）",
    )
    p.add_argument("--skip-std", action="store_true", help="[all] 跳过标准化（已有 B_std）")
    p.add_argument("--skip-summary", action="store_true", help="[all] 跳过汇总")
    p.add_argument("--skip-cashflow", action="store_true", help="[all] 跳过 HTML")
    p.add_argument(
        "--reclassify",
        action="store_true",
        help="[all] 汇总前用当前分类体系重刷 B_std",
    )
    p.add_argument(
        "--steps",
        nargs="+",
        default=None,
        help="[all] 自定义步骤: std reclassify summary cashflow review learn health",
    )

    sub = p.add_subparsers(dest="cmd", required=False)

    # ---- all / month（月结全流程；无参数默认是 apply-class，不是这里）----
    for all_name, all_help in [
        ("all", "月结全流程：std → summary → cashflow（不是默认无参数行为）"),
        ("month", "同 all：月结全流程 std → summary → cashflow"),
    ]:
        c = sub.add_parser(all_name, help=all_help)
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

    c = sub.add_parser(
        "export-class",
        help="导出付款分类总表（全部摘要+一级+二级，可在一个Excel里改）",
    )
    _add_path_args(c)
    c.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出路径（默认 learn_classify/data/付款分类总表.xlsx）",
    )

    c = sub.add_parser(
        "apply-class",
        help="读取分类总表有效一/二/三级 → 全部学习 → 重刷 B_std → 再导出总表",
    )
    _add_path_args(c)
    c.add_argument(
        "-i",
        "--input",
        default=None,
        help="分类总表路径（默认自动找 data 或 C_sum 下的 付款分类总表.xlsx）",
    )
    c.add_argument(
        "--no-reclassify",
        action="store_true",
        help="只学习，不重刷 B_std",
    )
    c.add_argument(
        "--no-reexport",
        action="store_true",
        help="应用后不重新导出总表",
    )
    c.add_argument(
        "--only-confirmed",
        action="store_true",
        help="仅学习「确认状态=确认」的行（默认关闭：表内有效二级全部学）",
    )

    c = sub.add_parser(
        "pending-dropdowns",
        help="给「分类修正记录.xlsx / 仍待人工_高频」加一/二/三级联动下拉",
    )
    _add_path_args(c)
    c.add_argument(
        "-i",
        "--input",
        default=None,
        help="分类修正记录路径（默认 C_sum 或 data 下 分类修正记录.xlsx）",
    )

    c = sub.add_parser(
        "apply-pending-fixes",
        help="把「仍待人工_高频」中确认行写回总表并学习叶子变更",
    )
    _add_path_args(c)
    c.add_argument(
        "-i",
        "--input",
        default=None,
        help="分类修正记录路径",
    )
    c.add_argument(
        "--master",
        default=None,
        help="付款分类总表路径",
    )
    c.add_argument(
        "--all-rows",
        action="store_true",
        help="不要求确认状态，凡有效二级都合并（默认仅确认）",
    )
    c.add_argument(
        "--no-reclassify",
        action="store_true",
        help="合并学习后不重刷 B_std",
    )

    c = sub.add_parser(
        "export-taxonomy",
        help="导出一级/二级科目体系 Excel（编辑下拉选项）",
    )
    _add_path_args(c)
    c.add_argument("-o", "--output", default=None, help="默认 data/科目体系.xlsx")

    c = sub.add_parser(
        "import-taxonomy",
        help="从科目体系 Excel 写回 taxonomy.json",
    )
    _add_path_args(c)
    c.add_argument("-i", "--input", default=None, help="默认 data/科目体系.xlsx")

    c = sub.add_parser(
        "cleanup-sheets",
        help="删除 B_std 中已废弃 sheet：分类使用说明/分类汇总/分项查看/透视源",
    )
    _add_path_args(c)

    c = sub.add_parser(
        "export-learn-pool",
        help="从 B_std 导出 D_财务付款_学习池（按项目/按月份·摘要去重）",
    )
    _add_path_args(c)

    c = sub.add_parser(
        "apply-learn-pool",
        help="学习池：确认/手改 → 记忆 → 已确认表 → 从项目/月份表剥离 → reclassify B_std",
    )
    _add_path_args(c)
    c.add_argument(
        "--no-reclassify",
        action="store_true",
        help="只学习与剥离，不重刷 B_std",
    )

    c = sub.add_parser(
        "compact-memory",
        help="压缩 memory.jsonl / corrections.jsonl（去重写回，显著减小体积）",
    )
    _add_path_args(c)
    c.add_argument(
        "--dry-run",
        action="store_true",
        help="只统计可压缩量，不写盘",
    )
    c.add_argument(
        "--no-backup",
        action="store_true",
        help="写盘前不备份 .bak_时间戳",
    )
    c.add_argument(
        "--memory-only",
        action="store_true",
        help="只压 memory.jsonl",
    )
    c.add_argument(
        "--corrections-only",
        action="store_true",
        help="只压 corrections.jsonl",
    )
    c.add_argument(
        "--corrections-keep-last",
        type=int,
        default=None,
        help="纠错日志去重后再只保留最近 N 条",
    )

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

    # 无子命令：默认 apply-class（总表有效分类全部学习 + 生成学习记录）
    if not args.cmd:
        paths = _paths_from_args(args)
        try:
            # 没有总表时先导出，再提示用户去改分类
            master = Path(args.input) if getattr(args, "input", None) else None
            candidates = []
            if master:
                candidates.append(master)
            candidates.extend(
                [paths.class_master_xlsx, paths.class_master_xlsx_sum]
            )
            has_master = any(p and Path(p).is_file() for p in candidates)
            if not has_master:
                print("未找到付款分类总表，先执行 export-class …")
                step_export_class_master(paths)
                print(
                    "\n下一步：打开 付款分类总表.xlsx → "
                    "调整一/二/三级科目 → 保存 → "
                    "再运行: python 财务付款一体化.py"
                )
                return 0
            step_apply_class_master(
                paths,
                master,
                reclassify=not getattr(args, "no_reclassify", False),
                reexport=not getattr(args, "no_reexport", False),
                only_confirmed=bool(getattr(args, "only_confirmed", False)),
            )
            return 0
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

        if args.cmd in ("all", "month"):
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

        if args.cmd == "export-class":
            step_export_class_master(
                paths, Path(args.output) if args.output else None
            )
            return 0

        if args.cmd == "apply-class":
            step_apply_class_master(
                paths,
                Path(args.input) if args.input else None,
                reclassify=not args.no_reclassify,
                reexport=not args.no_reexport,
                only_confirmed=bool(getattr(args, "only_confirmed", False)),
            )
            return 0

        if args.cmd == "pending-dropdowns":
            from learn_classify.class_master import apply_dropdowns_to_pending_report

            candidates = []
            if getattr(args, "input", None):
                candidates.append(Path(args.input))
            candidates.extend(
                [
                    paths.class_fix_report_xlsx_sum,
                    paths.class_fix_report_xlsx,
                ]
            )
            src = next((p for p in candidates if p and Path(p).is_file()), None)
            if not src:
                raise FileNotFoundError(
                    "未找到 分类修正记录.xlsx。请先生成审计/修正报告。"
                )
            n = apply_dropdowns_to_pending_report(src, paths.classify_data)
            # 同步副本
            import shutil

            for dest in (paths.class_fix_report_xlsx, paths.class_fix_report_xlsx_sum):
                if dest.resolve() != src.resolve():
                    try:
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(src, dest)
                    except Exception:
                        pass
            print(f"✅ 仍待人工_高频 下拉已就绪: {src} （{n} 表）")
            return 0

        if args.cmd == "apply-pending-fixes":
            from learn_classify.class_master import (
                merge_pending_fixes_into_master,
                apply_class_master,
            )
            import shutil

            pend_cands = []
            if getattr(args, "input", None):
                pend_cands.append(Path(args.input))
            pend_cands.extend(
                [paths.class_fix_report_xlsx_sum, paths.class_fix_report_xlsx]
            )
            pend = next((p for p in pend_cands if p and Path(p).is_file()), None)
            if not pend:
                raise FileNotFoundError("未找到 分类修正记录.xlsx")

            master_cands = []
            if getattr(args, "master", None):
                master_cands.append(Path(args.master))
            master_cands.extend(
                [paths.class_master_xlsx_sum, paths.class_master_xlsx]
            )
            master = next((p for p in master_cands if p and Path(p).is_file()), None)
            if not master:
                raise FileNotFoundError("未找到 付款分类总表.xlsx")

            print(f"待人工表: {pend}")
            print(f"分类总表: {master}")
            stats = merge_pending_fixes_into_master(
                pend,
                master,
                paths.classify_data,
                only_confirmed=not bool(getattr(args, "all_rows", False)),
            )
            print(f"   合并: {stats}")
            # 同步 data/C_sum 总表
            for dest in (paths.class_master_xlsx, paths.class_master_xlsx_sum):
                if dest.resolve() != master.resolve() and master.is_file():
                    try:
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(master, dest)
                    except Exception:
                        pass
            # 学习叶子变更
            learn_stats = apply_class_master(
                master,
                paths.std_dir,
                paths.classify_data,
                reclassify=not bool(getattr(args, "no_reclassify", False)),
                reexport=True,
                out_path=paths.class_master_xlsx,
                learn_log_path=paths.learn_log_xlsx,
            )
            if paths.class_master_xlsx.is_file():
                try:
                    shutil.copy2(paths.class_master_xlsx, paths.class_master_xlsx_sum)
                except Exception:
                    pass
            print(f"✅ 待人工确认已写回并学习: learned={learn_stats.get('learned')}")
            return 0

        if args.cmd == "export-taxonomy":
            step_export_taxonomy(
                paths, Path(args.output) if args.output else None
            )
            return 0

        if args.cmd == "import-taxonomy":
            step_import_taxonomy(
                paths, Path(args.input) if args.input else None
            )
            return 0

        if args.cmd == "cleanup-sheets":
            from .category_summary import remove_deprecated_sheets_in_std_dir

            stats = remove_deprecated_sheets_in_std_dir(paths.std_dir)
            print(f"✅ 已清理废弃 sheet: {stats}")
            return 0

        if args.cmd == "export-learn-pool":
            from .learn_pool import export_learn_pool

            stats = export_learn_pool(paths)
            print(f"✅ export-learn-pool: {stats}")
            return 0

        if args.cmd == "apply-learn-pool":
            from .learn_pool import apply_learn_pool

            stats = apply_learn_pool(
                paths,
                reclassify=not bool(getattr(args, "no_reclassify", False)),
            )
            print(f"✅ apply-learn-pool: {stats}")
            return 0

        if args.cmd == "compact-memory":
            import json
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
            if report.get("dry_run"):
                print("（dry-run：未写盘。去掉 --dry-run 执行压缩）")
            else:
                mem = report.get("memory") or {}
                corr = report.get("corrections") or {}
                if mem:
                    print(
                        f"✅ memory: {mem.get('before_lines')} → {mem.get('after_lines')} 行"
                        f"，节省 {mem.get('saved_bytes', 0) / 1024 / 1024:.1f} MB"
                    )
                if corr:
                    print(
                        f"✅ corrections: {corr.get('before_lines')} → {corr.get('after_lines')} 行"
                        f"，节省 {corr.get('saved_bytes', 0) / 1024 / 1024:.1f} MB"
                    )
            return 0

        parser.error(f"未知命令: {args.cmd}")
        return 2
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
