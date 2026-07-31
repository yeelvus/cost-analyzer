# -*- coding: utf-8 -*-
"""从人工纠错学习：写入记忆、更新供应商统计、生成规则建议。"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .engine import ClassifyEngine
from .normalize import core_memory_key, memory_key, normalize_text
from .store import (
    CorrectionEvent,
    CorrectionStore,
    MemoryEntry,
    MemoryStore,
    VendorStats,
    _now,
    load_rules,
    save_rules,
)
from .taxonomy import Taxonomy


def learn_one(
    *,
    taxonomy: Taxonomy,
    memory: MemoryStore,
    corrections: CorrectionStore,
    vendor_stats: Optional[VendorStats],
    subject: str,
    summary: str,
    vendor: str = "",
    new_l1: str = "",
    new_l2: str = "",
    new_l3: str = "",
    old_l2_id: str = "",
    source: str = "review",
    notes: str = "",
    also_summary_only: bool = True,
) -> Tuple[str, str, str]:
    """
    学习一条纠正。返回 (l1_name, l2_name, l3_name)。
    名称/ID 均可；有三级时记忆存 leaf=l3_id，否则 l2_id。
    """
    l1_id, l2_id, l3_id = taxonomy.resolve_triple(
        new_l1 or None, new_l2 or None, new_l3 or None
    )
    # 若只给了二级名但其实是三级正式名
    if not l3_id and new_l2:
        l3_try = taxonomy.resolve_l3_id(new_l2)
        if l3_try:
            l3_id = l3_try
            l2_id = taxonomy.level3[l3_id].parent
            l1_id = taxonomy.level2[l2_id].parent

    if not l2_id and not l3_id:
        raise ValueError(f"无法解析分类: l1={new_l1!r} l2={new_l2!r} l3={new_l3!r}")

    leaf_id = l3_id or l2_id
    l1_name, l2_name, l3_name = taxonomy.display_triple(leaf_id)
    old_l1_name, old_l2_name = ("", "")
    if old_l2_id:
        old_l1_name, old_l2_name, _ = taxonomy.display_triple(old_l2_id)

    corrections.append(
        CorrectionEvent(
            ts=_now(),
            subject=subject or "",
            summary=summary or "",
            vendor=vendor or "",
            old_l2_id=old_l2_id or "",
            new_l2_id=leaf_id,  # 存 leaf
            old_l1_name=old_l1_name,
            old_l2_name=old_l2_name,
            new_l1_name=l1_name,
            new_l2_name=l2_name + (f"/{l3_name}" if l3_name else ""),
            source=source,
            notes=notes,
        )
    )

    keys = [memory_key(subject, summary, vendor if vendor else None)]
    if also_summary_only and summary:
        keys.append(memory_key("", summary, None))
        # core 键：去月份/数字后的同型摘要共享（01月保安费用 ≈ 02月保安费用）
        ckey = core_memory_key(summary)
        if ckey:
            keys.append(ckey)

    for key in keys:
        if key in ("|", "||", ""):
            continue
        # core 键置信度略低，避免压过精确摘要
        is_core = key.startswith("|core:")
        memory.upsert(
            MemoryEntry(
                key=key,
                l2_id=leaf_id,
                subject=subject or "",
                summary=summary or "",
                vendor="" if is_core else (vendor or ""),
                source=source,
                confidence=0.9 if is_core else 1.0,
                notes=(notes + ("|core" if is_core else "")).strip("|"),
            )
        )

    if vendor_stats is not None and vendor:
        vendor_stats.observe(normalize_text(vendor), leaf_id, 1)
        vendor_stats.save()

    return l1_name, l2_name, l3_name


def import_corrections_excel(
    path: Path,
    *,
    taxonomy: Taxonomy,
    memory: MemoryStore,
    corrections: CorrectionStore,
    vendor_stats: Optional[VendorStats],
    sheet_name: int | str = 0,
) -> Dict[str, int]:
    """
    读取待复核/纠正表。需要列：
      摘要（或 科目+摘要）
      正确二级科目 或 正确二级 或 new_l2
    可选：正确一级科目、供应商名称、原二级科目ID、备注
    """
    df = pd.read_excel(path, sheet_name=sheet_name)
    colmap = {str(c).strip(): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in colmap:
                return colmap[n]
        return None

    c_subj = pick("科目", "原始科目")
    c_summ = pick("摘要", "业务摘要")
    c_vend = pick("供应商名称", "供应商", "收款单位")
    c_new1 = pick("正确一级科目", "正确一级", "new_l1", "一级科目_纠正")
    c_new2 = pick("正确二级科目", "正确二级", "new_l2", "二级科目_纠正")
    c_old2 = pick("分类ID", "l2_id", "原二级ID", "old_l2_id")
    c_note = pick("备注", "notes", "说明")

    if c_new2 is None:
        raise ValueError("纠正表缺少列：正确二级科目 / 正确二级 / new_l2")
    if c_summ is None and c_subj is None:
        raise ValueError("纠正表至少需要 摘要 或 科目")

    stats = {"total": 0, "learned": 0, "skipped": 0, "errors": 0}

    for _, row in df.iterrows():
        stats["total"] += 1
        new2 = "" if pd.isna(row.get(c_new2)) else str(row.get(c_new2)).strip()
        if not new2:
            stats["skipped"] += 1
            continue
        subject = "" if not c_subj or pd.isna(row.get(c_subj)) else str(row.get(c_subj)).strip()
        summary = "" if not c_summ or pd.isna(row.get(c_summ)) else str(row.get(c_summ)).strip()
        vendor = "" if not c_vend or pd.isna(row.get(c_vend)) else str(row.get(c_vend)).strip()
        new1 = "" if not c_new1 or pd.isna(row.get(c_new1)) else str(row.get(c_new1)).strip()
        old2 = "" if not c_old2 or pd.isna(row.get(c_old2)) else str(row.get(c_old2)).strip()
        notes = "" if not c_note or pd.isna(row.get(c_note)) else str(row.get(c_note)).strip()
        try:
            learn_one(
                taxonomy=taxonomy,
                memory=memory,
                corrections=corrections,
                vendor_stats=vendor_stats,
                subject=subject,
                summary=summary,
                vendor=vendor,
                new_l1=new1,
                new_l2=new2,
                old_l2_id=old2,
                source="review_excel",
                notes=notes,
            )
            stats["learned"] += 1
        except Exception:
            stats["errors"] += 1
    return stats


def suggest_rules_from_corrections(
    corrections: CorrectionStore,
    taxonomy: Taxonomy,
    existing_rules: List[Dict[str, Any]],
    min_support: int = 3,
    max_suggestions: int = 50,
) -> List[Dict[str, Any]]:
    """
    从纠错摘要中抽取高频 token，生成规则建议（不自动写入，需人确认）。
    """
    existing_targets = {
        (str(r.get("l2_id") or ""), str(r.get("含有") or r.get("contains") or ""))
        for r in existing_rules
    }
    # l2_id -> Counter of candidate keywords from summary
    buckets: Dict[str, Counter] = defaultdict(Counter)
    samples: Dict[str, List[str]] = defaultdict(list)

    stop = {
        "付款", "支付", "费用", "项目", "工程", "合同", "月份", "月", "年",
        "thai", "the", "and", "for", "费用报销", "报销",
    }

    for ev in corrections.read_all():
        if not ev.new_l2_id or not ev.summary:
            continue
        s = ev.summary
        samples[ev.new_l2_id].append(s)
        # 抽 2-6 字中文片段 + 英文词
        for m in re.findall(r"[\u4e00-\u9fff]{2,8}", s):
            if m in stop:
                continue
            buckets[ev.new_l2_id][m] += 1
        for m in re.findall(r"[A-Za-z]{3,12}", s):
            buckets[ev.new_l2_id][m.lower()] += 1

    suggestions = []
    for l2_id, counter in buckets.items():
        if l2_id not in taxonomy.level2:
            continue
        l1_name, l2_name = taxonomy.display(l2_id)
        for kw, n in counter.most_common(20):
            if n < min_support:
                continue
            contains = kw
            if (l2_id, contains) in existing_targets:
                continue
            # 避免过宽单字已过滤
            suggestions.append(
                {
                    "id": f"suggest_{l2_id}_{normalize_text(kw)[:20]}",
                    "l2_id": l2_id,
                    "priority": 120,
                    "含有": contains,
                    "不包含": "",
                    "support": n,
                    "level1": l1_name,
                    "level2": l2_name,
                    "examples": samples[l2_id][:3],
                    "status": "suggested",
                }
            )

    suggestions.sort(key=lambda x: -int(x["support"]))
    return suggestions[:max_suggestions]


def export_rule_suggestions(
    data_dir: Path,
    out_path: Path,
    min_support: int = 3,
) -> Path:
    data_dir = Path(data_dir)
    taxonomy = Taxonomy.load(data_dir / "taxonomy.json")
    corrections = CorrectionStore(data_dir / "corrections.jsonl")
    rules = load_rules(data_dir / "rules.json")
    sugg = suggest_rules_from_corrections(corrections, taxonomy, rules, min_support=min_support)
    df = pd.DataFrame(sugg)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if df.empty:
        df = pd.DataFrame(
            columns=["id", "l2_id", "priority", "含有", "不包含", "support", "level1", "level2", "status"]
        )
    df.to_excel(out_path, index=False)
    return out_path


def approve_rule_suggestions(
    data_dir: Path,
    suggestions_xlsx: Path,
    status_col: str = "status",
) -> int:
    """将 suggestions 表中 status=approved 的规则合并进 rules.json。"""
    data_dir = Path(data_dir)
    df = pd.read_excel(suggestions_xlsx)
    if status_col not in df.columns:
        raise ValueError(f"缺少列 {status_col}")
    approved = df[df[status_col].astype(str).str.lower().isin(["approved", "通过", "yes", "y", "1"])]
    rules = load_rules(data_dir / "rules.json")
    existing_ids = {str(r.get("id")) for r in rules}
    added = 0
    for _, row in approved.iterrows():
        rid = str(row.get("id") or f"rule_{added}")
        if rid in existing_ids:
            continue
        rules.append(
            {
                "id": rid,
                "l2_id": str(row.get("l2_id") or "").strip(),
                "priority": int(row.get("priority") or 120),
                "含有": str(row.get("含有") or row.get("contains") or "").strip(),
                "不包含": str(row.get("不包含") or row.get("excludes") or "").strip(),
            }
        )
        existing_ids.add(rid)
        added += 1
    if added:
        save_rules(data_dir / "rules.json", rules)
    return added


def rebuild_vendor_stats_from_std(
    std_dir: Path,
    vendor_stats: VendorStats,
    engine: ClassifyEngine,
    min_conf: float = 0.75,
) -> int:
    """用现有 std 高置信结果热身供应商画像（可选）。"""
    std_dir = Path(std_dir)
    n = 0
    for f in std_dir.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        try:
            df = pd.read_excel(f, sheet_name=0)
        except Exception:
            continue
        cols = {str(c).strip(): c for c in df.columns}
        if "供应商名称" not in cols and "供应商" not in cols:
            continue
        c_v = cols.get("供应商名称") or cols.get("供应商")
        c_s = cols.get("科目")
        c_z = cols.get("摘要")
        for _, row in df.iterrows():
            vend = "" if pd.isna(row.get(c_v)) else str(row.get(c_v)).strip()
            if not vend:
                continue
            subj = "" if not c_s or pd.isna(row.get(c_s)) else str(row.get(c_s)).strip()
            summ = "" if not c_z or pd.isna(row.get(c_z)) else str(row.get(c_z)).strip()
            r = engine.classify(subj, summ, vend)
            if r.confidence >= min_conf and r.layer not in {"pending"}:
                vendor_stats.observe(normalize_text(vend), r.l2_id, 1)
                n += 1
    vendor_stats.save()
    return n
