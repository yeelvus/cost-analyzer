# -*- coding: utf-8 -*-
"""
流水线编排：std / summary / cashflow + 分类相关步骤。
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Iterable, List, Optional

from .amount_mismatch import run_amount_mismatch
from .cashflow_build import run_cashflow
from .classify_raw_std import run_classify_raw_std
from .config import Paths
from .pkg import ensure_code_root
from .raw_std_build import run_raw_std
from .std_build import run_std
from .summary_build import run_summary


def _banner(title: str) -> None:
    print("\n" + "=" * 64)
    print(f"  {title}")
    print("=" * 64)


def step_raw_std(paths: Paths, *, prompt_project: bool = False) -> List[str]:
    _banner("步骤 · raw 标准化合并 (raw_std)")
    print(f"输入: {paths.raw_dir}/{{组名}}/*.xlsx")
    print(f"输出: {paths.raw_std_dir}/{{组名}}/{{组名}}_财务付款_raw_std.xlsx")
    print(
        f"项目匹配: 精确别名"
        f"{' + 弹窗' if prompt_project else '（弹窗关，--prompt-project 可开）'}"
    )
    t0 = time.time()
    outs = run_raw_std(paths, prompt_project=prompt_project)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_classify(paths: Paths) -> List[str]:
    _banner("步骤 · 摘要科目分类 (classify)")
    print(f"输入: {paths.raw_std_dir}/{{组名}}/*_raw_std.xlsx")
    print(f"输出: {paths.classified_dir}/{{组名}}/*_classified.xlsx")
    print("分类: 学习记忆 + 规则 → 一级/二级/三级科目")
    t0 = time.time()
    outs = run_classify_raw_std(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_std(paths: Paths, *, prompt_project: bool = False) -> List[str]:
    _banner("步骤 · 按项目拆分标准明细 (std)")
    print("优先: A2_classified → B_std/{组}/有支出|未支出/")
    print("回退: A_raw 全量重处理")
    t0 = time.time()
    outs = run_std(paths, prompt_project=prompt_project)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_summary(paths: Paths) -> List[str]:
    _banner("步骤 · 汇总分析 Excel (summary)")
    print("优先: A2_classified 直接汇总；回退扫 B_std")
    t0 = time.time()
    outs = run_summary(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return outs


def step_cashflow(paths: Paths):
    _banner("步骤 · 项目现金流量 HTML (cashflow，按组)")
    t0 = time.time()
    out = run_cashflow(paths)
    print(f"耗时 {time.time() - t0:.1f}s")
    return out


def step_amount_mismatch(paths: Paths, output: Optional[Path] = None):
    _banner("步骤 · 应付≠预扣+实付 不平明细汇总（按组）")
    t0 = time.time()
    out = run_amount_mismatch(paths, output_path=output)
    print(f"耗时 {time.time() - t0:.1f}s")
    return out


def step_reclassify(paths: Paths) -> dict:
    _banner("分类重刷 · reclassify B_std")
    ensure_code_root()
    from learn_classify.engine import ClassifyEngine
    from learn_classify.review import reclassify_std_dir

    stats = reclassify_std_dir(
        paths.std_dir, ClassifyEngine.from_data_dir(paths.classify_data)
    )
    print(f"✅ 重分类: {stats}")
    return stats


def step_export_review(
    paths: Paths, conf: float = 0.55, max_rows: Optional[int] = None
) -> str:
    _banner("导出待复核池")
    ensure_code_root()
    from learn_classify.engine import ClassifyEngine
    from learn_classify.review import export_review_pool

    out = export_review_pool(
        paths.std_dir,
        ClassifyEngine.from_data_dir(paths.classify_data),
        paths.review_xlsx,
        conf_threshold=conf,
        only_review=True,
        max_rows=max_rows,
    )
    print(f"✅ 待复核: {out}")
    return str(out)


def step_learn(paths: Paths, input_xlsx: Optional[Path] = None) -> dict:
    _banner("从纠正表学习")
    ensure_code_root()
    from learn_classify.learn import import_corrections_excel
    from learn_classify.store import CorrectionStore, MemoryStore, VendorStats
    from learn_classify.taxonomy import Taxonomy

    data = paths.classify_data
    stats = import_corrections_excel(
        Path(input_xlsx) if input_xlsx else paths.review_xlsx,
        taxonomy=Taxonomy.load(data / "taxonomy.json"),
        memory=MemoryStore(data / "memory.jsonl"),
        corrections=CorrectionStore(data / "corrections.jsonl"),
        vendor_stats=VendorStats(data / "vendor_stats.json"),
    )
    print(f"✅ 学习: {stats}")
    return stats


def step_learn_ledger(
    paths: Paths,
    ledger: Optional[Path] = None,
    *,
    sheet: Optional[str] = None,
    auto_taxonomy: bool = True,
    dry_run: bool = False,
    header_row: int = 1,
    summary_col: Optional[str] = None,
    l1_col: Optional[str] = None,
    l2_col: Optional[str] = None,
    l3_col: Optional[str] = None,
) -> dict:
    _banner("账簿已填科目 → 写入记忆")
    ensure_code_root()
    from ledger_book_classify import learn_from_ledger

    path = Path(ledger) if ledger else Path(paths.ledger_xlsx)
    print(f"账簿: {path}")
    print(f"学习库: {paths.classify_data}")
    if sheet:
        print(f"工作表: {sheet}")
    if summary_col or l1_col or l2_col:
        print(
            f"列: 摘要={summary_col or '自动'}  "
            f"一级={l1_col or '自动'}  二级={l2_col or '自动'}  "
            f"三级={l3_col or '自动/可空'}"
        )
    if dry_run:
        print("模式: 仅预览，不写 memory / taxonomy")
    stats = learn_from_ledger(
        path,
        paths.classify_data,
        sheet_name=(sheet or None),
        auto_taxonomy=auto_taxonomy,
        dry_run=dry_run,
        header_row=header_row,
        summary_col=summary_col,
        l1_col=l1_col,
        l2_col=l2_col,
        l3_col=l3_col,
    )
    print(f"✅ learn-ledger: learned={stats.get('learned')} unchanged={stats.get('skipped_unchanged')}")
    return stats


def step_health(paths: Paths) -> None:
    _banner("分类健康度")
    ensure_code_root()
    from collections import Counter

    import pandas as pd
    from learn_classify.engine import ClassifyEngine

    engine = ClassifyEngine.from_data_dir(paths.classify_data)
    layer_c: Counter = Counter()
    l1_c: Counter = Counter()
    l2_c: Counter = Counter()
    review_n = total = 0
    for f in paths.std_dir.rglob("*.xlsx"):
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
    print(f"待复核率: {review_n / max(total, 1) * 100:.1f}%")
    print("命中层:", dict(layer_c))
    print("一级 top:", l1_c.most_common(10))
    print("二级 top:", l2_c.most_common(15))


def step_export_class_master(paths: Paths, out: Optional[Path] = None) -> str:
    _banner("导出分类总表 export-class")
    import shutil

    ensure_code_root()
    from learn_classify.class_master import export_class_master
    from learn_classify.engine import ClassifyEngine

    dest = Path(out) if out else paths.class_master_xlsx
    engine = ClassifyEngine.from_data_dir(paths.classify_data)
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
    _banner("应用分类总表 apply-class")
    import shutil

    ensure_code_root()
    from learn_classify.class_master import apply_class_master

    paths.ensure_output_parents()
    candidates = []
    if master:
        candidates.append(Path(master))
    candidates.extend([paths.class_master_xlsx, paths.class_master_xlsx_sum])
    src = next((Path(p) for p in candidates if p and Path(p).is_file()), None)
    if src is None:
        raise FileNotFoundError(
            "未找到分类总表。可先: python 财务付款一体化.py export-class"
        )

    print(f"读取: {src}")
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
    if paths.learn_log_xlsx.is_file():
        try:
            paths.learn_log_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.learn_log_xlsx, paths.learn_log_xlsx_sum)
        except Exception:
            pass
    print(
        f"✅ 应用完成: learned={stats.get('learned')} "
        f"pending跳过={stats.get('skipped_pending')}"
    )
    return stats


def step_export_taxonomy(paths: Paths, out: Optional[Path] = None) -> str:
    _banner("导出科目体系 export-taxonomy")
    ensure_code_root()
    from learn_classify.taxonomy_io import export_taxonomy_excel

    dest = Path(out) if out else paths.taxonomy_xlsx
    path = export_taxonomy_excel(paths.classify_data, dest)
    print(f"✅ 科目体系: {path}")
    return str(path)


def step_import_taxonomy(paths: Paths, src: Optional[Path] = None) -> dict:
    _banner("导入科目体系 import-taxonomy")
    ensure_code_root()
    from learn_classify.taxonomy_io import import_taxonomy_excel

    excel = Path(src) if src else paths.taxonomy_xlsx
    stats = import_taxonomy_excel(paths.classify_data, excel, backup=True)
    print(f"✅ 已写入 taxonomy.json: {stats}")
    return stats


def run_pipeline(
    paths: Paths,
    *,
    steps: Optional[Iterable[str]] = None,
    skip_raw_std: bool = False,
    skip_classify: bool = False,
    skip_std: bool = False,
    skip_summary: bool = False,
    skip_cashflow: bool = False,
    reclassify_before_summary: bool = False,
    prompt_project: bool = False,
) -> dict:
    paths.ensure_output_parents()
    result: dict = {"paths": {}, "ok": True}

    # 快路径：raw_std → classify → summary → std → cashflow
    order = (
        list(steps)
        if steps
        else ["raw_std", "classify", "summary", "std", "cashflow"]
    )
    if skip_raw_std:
        order = [s for s in order if s != "raw_std"]
    if skip_classify:
        order = [s for s in order if s != "classify"]
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

    _banner("财务付款快路径流水线（按文件夹分组）")
    print(f"A_raw:        {paths.raw_dir}/{{组名}}/")
    print(f"A1_raw_std:   {paths.raw_std_dir}/{{组名}}/*_raw_std.xlsx")
    print(f"A2_classified:{paths.classified_dir}/{{组名}}/*_classified.xlsx")
    print(f"C_sum:        {paths.sum_dir}/{{组名}}/财务付款数据汇总.xlsx")
    print(f"B_std:        {paths.std_dir}/{{组名}}/有支出|未支出/")
    print(f"项目弹窗: {'开' if prompt_project else '关'}（--prompt-project）")
    print(f"步骤: {' → '.join(order)}")

    t_all = time.time()
    handlers = {
        "raw_std": lambda: result["paths"].__setitem__(
            "raw_std", step_raw_std(paths, prompt_project=prompt_project)
        ),
        "classify": lambda: result["paths"].__setitem__(
            "classify", step_classify(paths)
        ),
        "std": lambda: result["paths"].__setitem__(
            "std", step_std(paths, prompt_project=prompt_project)
        ),
        "summary": lambda: result["paths"].__setitem__("summary", step_summary(paths)),
        "cashflow": lambda: result["paths"].__setitem__(
            "cashflow", step_cashflow(paths)
        ),
        "reclassify": lambda: result["paths"].__setitem__(
            "reclassify", step_reclassify(paths)
        ),
        "review": lambda: result["paths"].__setitem__(
            "review", step_export_review(paths)
        ),
        "learn": lambda: result["paths"].__setitem__("learn", step_learn(paths)),
        "health": lambda: step_health(paths),
    }
    try:
        for step in order:
            if step not in handlers:
                raise ValueError(f"未知步骤: {step}")
            handlers[step]()
    except Exception as e:
        result["ok"] = False
        result["error"] = str(e)
        print(f"\n❌ 流水线失败: {e}")
        raise

    _banner(f"全部完成  总耗时 {time.time() - t_all:.1f}s")
    return result
