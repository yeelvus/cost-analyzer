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


def step_export_class_master(paths: Paths, out: Optional[Path] = None) -> str:
    """导出付款分类总表（全量摘要+一二级）。"""
    _banner("导出分类总表 export-class")
    import shutil
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.class_master import export_class_master
    from learn_classify.engine import ClassifyEngine

    dest = Path(out) if out else paths.class_master_xlsx
    engine = ClassifyEngine.from_data_dir(paths.classify_data)
    # 保留已有总表（data / C_sum）中「确认」标签
    preserve = None
    for cand in (dest, paths.class_master_xlsx, paths.class_master_xlsx_sum):
        if cand and Path(cand).is_file():
            preserve = Path(cand)
            break
    path, n = export_class_master(
        paths.std_dir,
        dest,
        paths.classify_data,
        engine=engine,
        preserve_path=preserve,
    )
    # 同步一份到 C_财务付款_sum，方便业务侧打开
    try:
        paths.class_master_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, paths.class_master_xlsx_sum)
        print(f"   副本: {paths.class_master_xlsx_sum}")
    except Exception as e:
        print(f"  ⚠️ 复制到 C_sum 失败: {e}")
    print(f"✅ 分类总表: {path}  （{n} 行）")
    return str(path)


def step_apply_class_master(
    paths: Paths,
    master: Optional[Path] = None,
    *,
    reclassify: bool = True,
    reexport: bool = True,
    only_confirmed: bool = False,
) -> dict:
    """按分类总表有效一/二/三级：全部学习 + 写学习文件 + 重刷 std + 再导出总表。"""
    _banner("应用分类总表 apply-class")
    import shutil
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.class_master import apply_class_master

    paths.ensure_output_parents()

    # 优先：用户指定 → data 下总表 → C_sum 副本
    candidates = []
    if master:
        candidates.append(Path(master))
    candidates.extend(
        [paths.class_master_xlsx, paths.class_master_xlsx_sum]
    )
    src = None
    for p in candidates:
        if p and Path(p).is_file():
            src = Path(p)
            break
    if src is None:
        raise FileNotFoundError(
            "未找到分类总表。请先运行: python 财务付款一体化.py export-class"
        )

    print(f"读取: {src}")
    print(
        "学习规则: 只学「最后一级叶子科目」变更"
        "（有三级看三级，否则看二级；一二级随叶子自动带出）"
        + ("；且仅确认行" if only_confirmed else "")
    )
    stats = apply_class_master(
        src,
        paths.std_dir,
        paths.classify_data,
        reclassify=reclassify,
        reexport=reexport,
        out_path=paths.class_master_xlsx,
        learn_log_path=paths.learn_log_xlsx,
        only_confirmed=only_confirmed,
    )
    if reexport and paths.class_master_xlsx.is_file():
        try:
            paths.class_master_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.class_master_xlsx, paths.class_master_xlsx_sum)
        except Exception:
            pass
    # 同步学习记录到 C_sum
    if paths.learn_log_xlsx.is_file():
        try:
            paths.learn_log_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.learn_log_xlsx, paths.learn_log_xlsx_sum)
            print(f"   学习记录副本: {paths.learn_log_xlsx_sum}")
        except Exception as e:
            print(f"  ⚠️ 复制学习记录失败: {e}")
    print(f"✅ 应用完成: learned={stats.get('learned')} pending跳过={stats.get('skipped_pending')}")
    if stats.get("learn_log"):
        print(f"   📂 学习记录 Excel: {stats['learn_log']}")
    if stats.get("memory_file"):
        print(f"   📂 记忆 jsonl: {stats['memory_file']}")
    return stats


def step_export_taxonomy(paths: Paths, out: Optional[Path] = None) -> str:
    """导出一级/二级科目体系 Excel（改下拉选项用）。"""
    _banner("导出科目体系 export-taxonomy")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.taxonomy_io import export_taxonomy_excel

    dest = Path(out) if out else paths.taxonomy_xlsx
    path = export_taxonomy_excel(paths.classify_data, dest)
    print(f"✅ 科目体系: {path}")
    print("   编辑后执行: python 财务付款一体化.py import-taxonomy")
    return str(path)


def step_import_taxonomy(paths: Paths, src: Optional[Path] = None) -> dict:
    """从 Excel 导入科目体系 → taxonomy.json。"""
    _banner("导入科目体系 import-taxonomy")
    import sys

    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from learn_classify.taxonomy_io import import_taxonomy_excel

    excel = Path(src) if src else paths.taxonomy_xlsx
    stats = import_taxonomy_excel(paths.classify_data, excel, backup=True)
    print(f"✅ 已写入 taxonomy.json: {stats}")
    print("   建议接着: export-class 或 dropdowns 刷新下拉")
    return stats


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
