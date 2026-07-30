# -*- coding: utf-8 -*-
"""命令行入口。"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

DEFAULT_DATA = Path(__file__).resolve().parent / "data"
DEFAULT_STD = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/B_财务付款_std"
)


def _engine(data_dir: Path):
    from .engine import ClassifyEngine

    return ClassifyEngine.from_data_dir(data_dir)


def cmd_classify(args):
    engine = _engine(Path(args.data_dir))
    r = engine.classify(args.subject, args.summary, args.vendor)
    print(
        json.dumps(
            {
                "一级科目": r.l1_name,
                "二级科目": r.l2_name,
                "分类ID": r.l2_id,
                "命中层": r.layer,
                "置信度": r.confidence,
                "原因": r.reason,
                "候选": r.candidates,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


def cmd_export_review(args):
    from .review import export_review_pool

    engine = _engine(Path(args.data_dir))
    out = export_review_pool(
        Path(args.std_dir),
        engine,
        Path(args.output),
        conf_threshold=args.conf,
        only_review=not args.all,
        max_rows=args.max_rows,
    )
    print(f"✅ 待复核已导出: {out}")


def cmd_learn(args):
    from .learn import import_corrections_excel
    from .store import CorrectionStore, MemoryStore, VendorStats
    from .taxonomy import Taxonomy

    data_dir = Path(args.data_dir)
    stats = import_corrections_excel(
        Path(args.input),
        taxonomy=Taxonomy.load(data_dir / "taxonomy.json"),
        memory=MemoryStore(data_dir / "memory.jsonl"),
        corrections=CorrectionStore(data_dir / "corrections.jsonl"),
        vendor_stats=VendorStats(data_dir / "vendor_stats.json"),
    )
    print(f"✅ 学习完成: {stats}")


def cmd_suggest_rules(args):
    from .learn import export_rule_suggestions

    out = export_rule_suggestions(
        Path(args.data_dir),
        Path(args.output),
        min_support=args.min_support,
    )
    print(f"✅ 规则建议已导出: {out}")
    print("  在 Excel 中将 status 改为 approved 后执行: approve-rules")


def cmd_approve_rules(args):
    from .learn import approve_rule_suggestions

    n = approve_rule_suggestions(Path(args.data_dir), Path(args.input))
    print(f"✅ 已合并 {n} 条规则到 rules.json")


def cmd_migrate(args):
    from .migrate import run_migration

    report = run_migration(
        Path(args.data_dir),
        Path(args.map),
        do_backup=not args.no_backup,
        reclassify_std=Path(args.reclassify_std) if args.reclassify_std else None,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


def cmd_reclassify(args):
    from .review import reclassify_std_dir

    engine = _engine(Path(args.data_dir))
    stats = reclassify_std_dir(Path(args.std_dir), engine)
    print(f"✅ 重分类完成: {stats}")


def cmd_health(args):
    from collections import Counter

    import pandas as pd

    engine = _engine(Path(args.data_dir))
    std = Path(args.std_dir)
    layer_c = Counter()
    l1_c = Counter()
    l2_c = Counter()
    review_n = 0
    total = 0
    for f in std.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        try:
            df = pd.read_excel(f, sheet_name=0)
        except Exception:
            continue
        cols = {str(c).strip(): c for c in df.columns}
        for _, row in df.iterrows():
            total += 1
            vend_col = cols.get("供应商名称") or cols.get("供应商")
            r = engine.classify(
                row.get(cols["科目"]) if "科目" in cols else None,
                row.get(cols["摘要"]) if "摘要" in cols else None,
                row.get(vend_col) if vend_col else None,
            )
            layer_c[r.layer] += 1
            l1_c[r.l1_name] += 1
            l2_c[r.l2_name] += 1
            if r.needs_review:
                review_n += 1
    print("=== 分类健康度 ===")
    print(f"总笔数: {total}")
    print(f"待复核率: {review_n/max(total,1)*100:.1f}%")
    print("命中层:", dict(layer_c))
    print("一级 top:", l1_c.most_common(10))
    print("二级 top:", l2_c.most_common(15))
    other_n = l2_c.get("其他", 0) + l2_c.get("待复核", 0)
    print(f"其他+待复核占比: {other_n/max(total,1)*100:.1f}%")


def cmd_seed_vendor(args):
    from .learn import rebuild_vendor_stats_from_std
    from .store import VendorStats

    data_dir = Path(args.data_dir)
    engine = _engine(data_dir)
    n = rebuild_vendor_stats_from_std(
        Path(args.std_dir),
        VendorStats(data_dir / "vendor_stats.json"),
        engine,
        min_conf=args.conf,
    )
    print(f"✅ 供应商画像观察次数: {n}")


def cmd_compact(args):
    import json

    from .store import compact_data_dir

    only_mem = bool(getattr(args, "memory_only", False))
    only_corr = bool(getattr(args, "corrections_only", False))
    do_memory = only_mem or not only_corr
    do_corr = only_corr or not only_mem
    report = compact_data_dir(
        Path(args.data_dir),
        memory=do_memory,
        corrections=do_corr,
        backup=not bool(getattr(args, "no_backup", False)),
        corrections_keep_last=getattr(args, "corrections_keep_last", None),
        dry_run=bool(getattr(args, "dry_run", False)),
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


def build_parser():
    p = argparse.ArgumentParser(
        prog="learn_classify",
        description="财务付款一二级分类：可学习闭环 + 可迁移 taxonomy",
    )
    p.add_argument(
        "--data-dir",
        default=str(DEFAULT_DATA),
        help="数据目录（taxonomy/rules/memory...）",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("classify", help="单条分类调试")
    c.add_argument("--subject", default="")
    c.add_argument("--summary", required=True)
    c.add_argument("--vendor", default="")
    c.set_defaults(func=cmd_classify)

    c = sub.add_parser("export-review", help="导出待复核表")
    c.add_argument("--std-dir", default=str(DEFAULT_STD))
    c.add_argument("-o", "--output", default=str(DEFAULT_DATA / "待复核池.xlsx"))
    c.add_argument("--conf", type=float, default=0.55)
    c.add_argument("--all", action="store_true", help="导出全部而非仅待复核")
    c.add_argument("--max-rows", type=int, default=None)
    c.set_defaults(func=cmd_export_review)

    c = sub.add_parser("learn", help="从纠正 Excel 学习")
    c.add_argument("-i", "--input", required=True, help="填写了正确二级的表")
    c.set_defaults(func=cmd_learn)

    c = sub.add_parser("suggest-rules", help="从纠错生成规则建议")
    c.add_argument("-o", "--output", default=str(DEFAULT_DATA / "规则建议.xlsx"))
    c.add_argument("--min-support", type=int, default=3)
    c.set_defaults(func=cmd_suggest_rules)

    c = sub.add_parser("approve-rules", help="合并 approved 规则")
    c.add_argument("-i", "--input", required=True)
    c.set_defaults(func=cmd_approve_rules)

    c = sub.add_parser("migrate", help="执行 taxonomy 迁移/分类转移")
    c.add_argument("--map", required=True, help="migration json 路径")
    c.add_argument("--no-backup", action="store_true")
    c.add_argument("--reclassify-std", default=None, help="迁移后重写 std 目录")
    c.set_defaults(func=cmd_migrate)

    c = sub.add_parser("reclassify", help="用当前体系重分类 std")
    c.add_argument("--std-dir", default=str(DEFAULT_STD))
    c.set_defaults(func=cmd_reclassify)

    c = sub.add_parser("health", help="分类健康度")
    c.add_argument("--std-dir", default=str(DEFAULT_STD))
    c.set_defaults(func=cmd_health)

    c = sub.add_parser("seed-vendor", help="从 std 热身供应商画像")
    c.add_argument("--std-dir", default=str(DEFAULT_STD))
    c.add_argument("--conf", type=float, default=0.75)
    c.set_defaults(func=cmd_seed_vendor)

    c = sub.add_parser("compact", help="压缩 memory/corrections jsonl")
    c.add_argument("--dry-run", action="store_true")
    c.add_argument("--no-backup", action="store_true")
    c.add_argument("--memory-only", action="store_true")
    c.add_argument("--corrections-only", action="store_true")
    c.add_argument("--corrections-keep-last", type=int, default=None)
    c.set_defaults(func=cmd_compact)

    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        args.func(args)
        return 0
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
