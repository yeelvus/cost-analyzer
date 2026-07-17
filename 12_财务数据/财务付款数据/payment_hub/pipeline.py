# -*- coding: utf-8 -*-
"""
财务付款一体化流水线（自包含，不依赖 3_/4_/5_ 脚本文件）。
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Iterable, List, Optional

from .cashflow_build import run_cashflow
from .config import Paths
from .std_build import run_std
from .summary_build import run_summary


def _banner(title: str) -> None:
    print("\n" + "=" * 64)
    print(f"  {title}")
    print("=" * 64)


def step_std(paths: Paths) -> List[str]:
    _banner("步骤 · 原始付款 → 标准明细 (std)")
    t0 = time.time()
    outs = run_std(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_summary(paths: Paths) -> List[str]:
    _banner("步骤 · 汇总分析 Excel (summary)")
    t0 = time.time()
    outs = run_summary(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_cashflow(paths: Paths) -> str:
    _banner("步骤 · 项目现金流量 HTML (cashflow)")
    t0 = time.time()
    out = run_cashflow(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return out


def step_reclassify(paths: Paths) -> dict:
    _banner("分类重刷 · reclassify B_std")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.engine import ClassifyEngine
    from learn_classify.review import reclassify_std_dir

    engine = ClassifyEngine.from_data_dir(paths.classify_data)
    stats = reclassify_std_dir(paths.std_dir, engine)
    print(f"✅ 重分类: {stats}")
    return stats


def step_export_review(
    paths: Paths, conf: float = 0.55, max_rows: Optional[int] = None
) -> str:
    _banner("导出待复核池")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.engine import ClassifyEngine
    from learn_classify.review import export_review_pool

    engine = ClassifyEngine.from_data_dir(paths.classify_data)
    out = export_review_pool(
        paths.std_dir,
        engine,
        paths.review_xlsx,
        conf_threshold=conf,
        only_review=True,
        max_rows=max_rows,
    )
    print(f"✅ 待复核: {out}")
    return str(out)


def step_learn(paths: Paths, input_xlsx: Optional[Path] = None) -> dict:
    _banner("从纠正表学习")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.learn import import_corrections_excel
    from learn_classify.store import CorrectionStore, MemoryStore, VendorStats
    from learn_classify.taxonomy import Taxonomy

    data = paths.classify_data
    src = Path(input_xlsx) if input_xlsx else paths.review_xlsx
    stats = import_corrections_excel(
        src,
        taxonomy=Taxonomy.load(data / "taxonomy.json"),
        memory=MemoryStore(data / "memory.jsonl"),
        corrections=CorrectionStore(data / "corrections.jsonl"),
        vendor_stats=VendorStats(data / "vendor_stats.json"),
    )
    print(f"✅ 学习: {stats}")
    return stats


def step_health(paths: Paths) -> None:
    _banner("分类健康度")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.cli import main as lc_main

    lc_main(
        ["--data-dir", str(paths.classify_data), "health", "--std-dir", str(paths.std_dir)]
    )


def run_pipeline(
    paths: Paths,
    *,
    steps: Optional[Iterable[str]] = None,
    skip_std: bool = False,
    skip_summary: bool = False,
    skip_cashflow: bool = False,
    reclassify_before_summary: bool = False,
) -> dict:
    paths.ensure_output_parents()
    result = {"paths": {}, "ok": True}

    order = list(steps) if steps else ["std", "summary", "cashflow"]
    if skip_std:
        order = [s for s in order if s != "std"]
    if skip_summary:
        order = [s for s in order if s != "summary"]
    if skip_cashflow:
        order = [s for s in order if s != "cashflow"]
    if reclassify_before_summary and "reclassify" not in order:
        if "summary" in order:
            order.insert(order.index("summary"), "reclassify")
        else:
            order.append("reclassify")

    _banner("财务付款一体化流水线（payment_hub 自包含）")
    print(f"原始: {paths.raw_dir}")
    print(f"标准: {paths.std_dir}")
    print(f"汇总: {paths.sum_xlsx}")
    print(f"HTML: {paths.cashflow_html}")
    print(f"步骤: {' → '.join(order)}")

    t_all = time.time()
    try:
        for step in order:
            if step == "std":
                result["paths"]["std"] = step_std(paths)
            elif step == "summary":
                result["paths"]["summary"] = step_summary(paths)
            elif step == "cashflow":
                result["paths"]["cashflow"] = step_cashflow(paths)
            elif step == "reclassify":
                result["paths"]["reclassify"] = step_reclassify(paths)
            elif step == "review":
                result["paths"]["review"] = step_export_review(paths)
            elif step == "learn":
                result["paths"]["learn"] = step_learn(paths)
            elif step == "health":
                step_health(paths)
            else:
                raise ValueError(f"未知步骤: {step}")
    except Exception as e:
        result["ok"] = False
        result["error"] = str(e)
        print(f"\n❌ 流水线失败: {e}")
        raise

    _banner(f"全部完成  总耗时 {time.time() - t_all:.1f}s")
    return result
