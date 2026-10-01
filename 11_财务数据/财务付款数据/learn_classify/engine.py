# -*- coding: utf-8 -*-
"""
多层分类引擎（可解释）：支持一级/二级/三级。

调用方传入的 summary 应是业务摘要（项目名前缀已由 summary_recognize 去掉）。
付款月结、账簿填科目、学习池刷新都走 learn_classify.classify_text。

链路：精确记忆 → core 记忆（去月份/数字同型）→ 关键词规则 → 正则模式
     → 相似记忆（特征/Jaccard 泛化）→ 科目映射 → 供应商众数 → 待复核。

记忆与规则的 leaf_id 可为 L2 或 L3；有三级时优先落到三级。
确认/手改写入记忆后，core + similar 会抬升相近未确认摘要的预分类水平。
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, FrozenSet, List, Optional, Set, Tuple

from .normalize import (
    core_memory_key,
    core_summary,
    jaccard,
    memory_key,
    normalize_text,
    soft_key,
    summary_bigrams,
    summary_features,
)
from .store import MemoryEntry, MemoryStore, VendorStats, load_rules
from .taxonomy import Taxonomy

# 摘要正则模式（关键词规则难以表达的中间插词场景）
# (compiled_regex, leaf_id, conf, reason)
_SUMMARY_PATTERNS: List[Tuple[re.Pattern, str, float, str]] = [
    # 代付 / 代xx付款 / 代溧阳付款 / 代张三支付
    (
        re.compile(r"代.{0,20}(付款|支付|垫付|垫款)|代付|代垫"),
        "L2_exclude_agent_pay",
        0.88,
        "pattern:代*付款",
    ),
]


@dataclass
class ClassifyResult:
    l1_id: str
    l2_id: str
    l3_id: str
    l1_name: str
    l2_name: str
    l3_name: str
    layer: str
    confidence: float
    rule_id: str = ""
    candidates: List[str] = field(default_factory=list)
    reason: str = ""
    # 兼容：leaf
    leaf_id: str = ""

    @property
    def needs_review(self) -> bool:
        return self.layer in {"pending", "low_confidence", "conflict"} or self.confidence < 0.55


class ClassifyEngine:
    def __init__(
        self,
        taxonomy: Taxonomy,
        rules: List[Dict[str, Any]],
        memory: Optional[MemoryStore] = None,
        vendor_stats: Optional[VendorStats] = None,
        subject_map: Optional[Dict[str, str]] = None,
        vendor_min_count: int = 3,
        vendor_min_ratio: float = 0.7,
        conf_review_threshold: float = 0.55,
    ):
        self.taxonomy = taxonomy
        self.rules = sorted(rules, key=lambda r: (-int(r.get("priority", 100)), str(r.get("id", ""))))
        self.memory = memory
        self.vendor_stats = vendor_stats
        self.subject_map = subject_map or {}
        self.vendor_min_count = vendor_min_count
        self.vendor_min_ratio = vendor_min_ratio
        self.conf_review_threshold = conf_review_threshold
        self._compiled = [self._compile_rule(r) for r in self.rules]
        # 相似记忆索引（懒构建）
        self._sim_ready = False
        self._core_to_entries: Dict[str, List[MemoryEntry]] = {}
        self._feat_to_idxs: Dict[str, List[int]] = {}
        self._sim_entries: List[Tuple[MemoryEntry, str, Set[str], FrozenSet[str]]] = []

    @classmethod
    def from_data_dir(cls, data_dir: Path) -> "ClassifyEngine":
        data_dir = Path(data_dir)
        taxonomy = Taxonomy.load(data_dir / "taxonomy.json")
        rules = load_rules(data_dir / "rules.json")
        memory = MemoryStore(data_dir / "memory.jsonl")
        vendor = VendorStats(data_dir / "vendor_stats.json")
        subject_map_path = data_dir / "subject_map.json"
        subject_map = {}
        if subject_map_path.exists():
            import json

            with open(subject_map_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            subject_map = {str(k): str(v) for k, v in raw.items()}
        return cls(taxonomy, rules, memory, vendor, subject_map)

    def _compile_rule(self, rule: Dict[str, Any]) -> Dict[str, Any]:
        contains = str(rule.get("含有") or rule.get("contains") or "")
        excludes = str(rule.get("不包含") or rule.get("excludes") or "")
        or_groups = []
        for group in [g.strip() for g in contains.split(";") if g.strip()]:
            and_kws = [normalize_text(k) for k in group.split("&") if k.strip()]
            and_kws = [k for k in and_kws if k]
            if and_kws:
                or_groups.append(and_kws)
        exclude_list = [normalize_text(k) for k in excludes.split(";") if k.strip()]
        l3_id = str(rule.get("l3_id") or "").strip()
        l2_id = str(rule.get("l2_id") or rule.get("target") or "").strip()
        return {
            "raw": rule,
            "id": str(rule.get("id") or ""),
            "l2_id": l2_id,
            "l3_id": l3_id,
            "priority": int(rule.get("priority", 100)),
            "or_groups": or_groups,
            "excludes": exclude_list,
        }

    def _match_compiled(self, summary_norm: str, compiled: Dict[str, Any]) -> bool:
        if not summary_norm or not compiled["or_groups"]:
            return False
        matched = any(all(kw in summary_norm for kw in group) for group in compiled["or_groups"])
        if not matched:
            return False
        if any(ex in summary_norm for ex in compiled["excludes"]):
            return False
        return True

    def _result_from_leaf(
        self,
        leaf_id: str,
        layer: str,
        confidence: float,
        rule_id: str = "",
        reason: str = "",
        candidates: Optional[List[str]] = None,
    ) -> ClassifyResult:
        tax = self.taxonomy
        # leaf 可能是 L3 或 L2
        if leaf_id in tax.level3:
            if not tax.is_active_leaf(leaf_id):
                leaf_id = tax.pending_l2_id
                layer = "pending"
                confidence = min(confidence, 0.4)
                reason = (reason + " | inactive→pending").strip(" |")
        elif leaf_id in tax.level2:
            if not tax.is_active_l2(leaf_id):
                leaf_id = tax.pending_l2_id
                layer = "pending"
                confidence = min(confidence, 0.4)
                reason = (reason + " | inactive→pending").strip(" |")
        else:
            leaf_id = tax.pending_l2_id
            layer = "pending"
            confidence = min(confidence, 0.4)

        l1_name, l2_name, l3_name = tax.display_triple(leaf_id)
        if leaf_id in tax.level3:
            l3_id = leaf_id
            l2_id = tax.level3[leaf_id].parent
            l1_id = tax.level2[l2_id].parent if l2_id in tax.level2 else ""
        else:
            l3_id = ""
            l2_id = leaf_id if leaf_id in tax.level2 else tax.pending_l2_id
            l1_id = tax.level2[l2_id].parent if l2_id in tax.level2 else ""

        return ClassifyResult(
            l1_id=l1_id,
            l2_id=l2_id,
            l3_id=l3_id,
            l1_name=l1_name,
            l2_name=l2_name,
            l3_name=l3_name,
            layer=layer,
            confidence=confidence,
            rule_id=rule_id,
            candidates=candidates or [],
            reason=reason,
            leaf_id=leaf_id,
        )

    # 兼容旧名
    def _result_from_l2(self, l2_id, layer, confidence, rule_id="", reason="", candidates=None):
        return self._result_from_leaf(l2_id, layer, confidence, rule_id, reason, candidates)

    def invalidate_sim_index(self) -> None:
        """记忆更新后调用，下次 classify 重建相似索引。"""
        self._sim_ready = False
        self._core_to_entries = {}
        self._feat_to_idxs = {}
        self._sim_entries = []

    @staticmethod
    def _entry_rank(e: MemoryEntry) -> Tuple[float, str]:
        """同摘要多键时择优：人工/学习池 > 高置信 > 较新。"""
        src = (e.source or "").lower()
        src_w = 0.0
        if src in {
            "learn_pool",
            "manual",
            "review",
            "review_excel",
            "user",
            "correction",
            "class_master_leaf",
            "class_master",
        }:
            src_w = 3.0
        elif src.startswith("user") or "correct" in src:
            src_w = 2.5
        elif src.startswith("ai") or "auto" in src:
            src_w = 1.0
        else:
            src_w = 1.5
        conf = float(e.confidence or 0.5)
        return (src_w + conf, e.updated_at or "")

    def _ensure_sim_index(self) -> None:
        """从 memory 构建 core / feature 倒排索引（一次/引擎生命周期）。"""
        if self._sim_ready:
            return
        self._sim_ready = True
        self._core_to_entries = {}
        self._feat_to_idxs = {}
        self._sim_entries = []
        if self.memory is None:
            return

        # 同一摘要文本可能对应多个 key / 不同 leaf → 先按摘要择优
        by_summ: Dict[str, List[MemoryEntry]] = {}
        for entry in self.memory.entries.values():
            summ = (entry.summary or "").strip()
            if not summ or not entry.l2_id:
                continue
            sn = normalize_text(summ)
            if not sn:
                continue
            by_summ.setdefault(sn, []).append(entry)

        for sn, ents in by_summ.items():
            entry = max(ents, key=self._entry_rank)
            summ = (entry.summary or "").strip()
            core = core_summary(summ)
            feats = summary_features(summ)
            bigs = summary_bigrams(core or sn)
            idx = len(self._sim_entries)
            self._sim_entries.append((entry, core, feats, bigs))

            if core:
                self._core_to_entries.setdefault(core, []).append(entry)
            for tok in feats:
                if len(tok) < 2:
                    continue
                bucket = self._feat_to_idxs.setdefault(tok, [])
                if len(bucket) < 80:
                    bucket.append(idx)

    def _majority_leaf(
        self,
        entries: List[MemoryEntry],
        *,
        min_ratio: float = 0.8,
    ) -> Optional[Tuple[str, float, str]]:
        """entries → (leaf_id, agreement_ratio, sample_summary)。"""
        if not entries:
            return None
        # 同 leaf 按置信度加权；人工/学习池权重更高
        weights: Dict[str, float] = {}
        sample: Dict[str, str] = {}
        for e in entries:
            lid = e.l2_id
            if not lid:
                continue
            w = float(e.confidence or 0.8)
            src = (e.source or "").lower()
            if src in {
                "learn_pool",
                "manual",
                "review",
                "review_excel",
                "user",
                "correction",
                "class_master_leaf",
            }:
                w *= 1.25
            weights[lid] = weights.get(lid, 0.0) + w
            # 优先保留高置信样本摘要
            prev = sample.get(lid)
            if prev is None or w >= weights[lid]:
                sample[lid] = e.summary or e.key
        if not weights:
            return None
        best_leaf, best_w = max(weights.items(), key=lambda x: x[1])
        total = sum(weights.values())
        ratio = best_w / total if total else 0.0
        if len(weights) > 1 and ratio < min_ratio:
            return None
        return best_leaf, ratio, sample.get(best_leaf, "")

    def _match_core_memory(self, summ: str) -> Optional[ClassifyResult]:
        """去数字/月份后的 core 精确命中（跨 01月/02月 等同型摘要）。"""
        # 显式写入的 |core:xxx 键（learn_one 产出）优先
        ckey = core_memory_key(summ)
        if ckey and self.memory is not None:
            entry = self.memory.get(ckey)
            if entry and entry.l2_id:
                conf = min(0.92, float(entry.confidence or 0.92))
                return self._result_from_leaf(
                    entry.l2_id,
                    layer="memory_core",
                    confidence=conf,
                    reason=f"memory_core:{entry.summary or ckey}",
                )

        self._ensure_sim_index()
        core = core_summary(summ)
        if not core:
            return None
        entries = self._core_to_entries.get(core) or []
        # 冲突保守：需 ≥80% 加权一致才 core 命中，否则交给 similar
        maj = self._majority_leaf(entries, min_ratio=0.8)
        if not maj:
            return None
        leaf_id, ratio, sample = maj
        conf = 0.82 + 0.1 * min(1.0, ratio)
        return self._result_from_leaf(
            leaf_id,
            layer="memory_core",
            confidence=min(0.92, conf),
            reason=f"memory_core≈{sample[:40]}",
        )

    def _match_similar_memory(self, summ: str) -> Optional[ClassifyResult]:
        """
        相似记忆：特征重叠 + bigram Jaccard + 子串包含。
        置信度低于精确/core 记忆与硬规则，避免覆盖明确规则。
        """
        self._ensure_sim_index()
        if not self._sim_entries:
            return None

        q_core = core_summary(summ)
        q_norm = normalize_text(summ)
        q_feats = summary_features(summ)
        q_bigs = summary_bigrams(q_core or q_norm)
        if not q_feats and not q_bigs:
            return None

        # 候选：倒排特征命中
        cand_idxs: Set[int] = set()
        ranked_feats = sorted(q_feats, key=lambda t: (-len(t), t))
        for tok in ranked_feats[:24]:
            for idx in self._feat_to_idxs.get(tok, []):
                cand_idxs.add(idx)
                if len(cand_idxs) >= 350:
                    break
            if len(cand_idxs) >= 350:
                break

        if not cand_idxs:
            return None

        scored: List[Tuple[float, str, str]] = []  # score, leaf, sample
        for idx in cand_idxs:
            entry, m_core, m_feats, m_bigs = self._sim_entries[idx]
            if not entry.l2_id:
                continue

            score = 0.0
            # 1) core 子串 / 相等
            if q_core and m_core:
                if q_core == m_core:
                    score = max(score, 0.95)
                elif q_core in m_core or m_core in q_core:
                    shorter = min(len(q_core), len(m_core))
                    longer = max(len(q_core), len(m_core))
                    if shorter >= 2 and shorter / longer >= 0.45:
                        score = max(score, 0.72 + 0.2 * (shorter / longer))

            # 2) 特征 Jaccard
            fj = jaccard(q_feats, m_feats)
            if fj > 0:
                score = max(score, 0.35 + 0.55 * fj)

            # 3) bigram Jaccard
            bj = jaccard(set(q_bigs), set(m_bigs))
            if bj > 0:
                score = max(score, 0.3 + 0.5 * bj)

            # 源加权：人工/学习池略高
            src = (entry.source or "").lower()
            if src in {
                "learn_pool",
                "manual",
                "review",
                "review_excel",
                "user",
                "correction",
            }:
                score = min(1.0, score + 0.03)

            if score < 0.52:
                continue
            scored.append((score, entry.l2_id, entry.summary or entry.key))

        if not scored:
            return None

        scored.sort(key=lambda x: -x[0])
        best_score, best_leaf, best_sample = scored[0]

        # 前几名叶子不一致且分数接近 → 放弃（歧义）
        near = [s for s in scored[:8] if s[0] >= best_score - 0.08]
        leaves = {s[1] for s in near}
        if len(leaves) > 1:
            cnt = Counter(s[1] for s in near)
            top_leaf, top_n = cnt.most_common(1)[0]
            if top_n / len(near) < 0.67:
                return None
            best_leaf = top_leaf
            best_sample = next(s[2] for s in near if s[1] == top_leaf)
            best_score = max(s[0] for s in near if s[1] == top_leaf)

        conf = 0.55 + 0.30 * min(1.0, best_score)
        conf = min(0.84, conf)
        if conf < self.conf_review_threshold:
            return None

        return self._result_from_leaf(
            best_leaf,
            layer="similar_memory",
            confidence=conf,
            reason=f"similar≈{best_sample[:48]}@{best_score:.2f}",
        )

    def classify(
        self,
        subject: Any = None,
        summary: Any = None,
        vendor: Any = None,
    ) -> ClassifyResult:
        subj = "" if subject is None else str(subject).strip()
        summ = "" if summary is None else str(summary).strip()
        vend = "" if vendor is None else str(vendor).strip()
        summ_norm = normalize_text(summ)
        subj_norm = normalize_text(subj)
        vend_norm = normalize_text(vend)

        # 精确记忆（l2_id 字段存 leaf_id）
        if self.memory is not None:
            for key in (
                memory_key(subj, summ, vend),
                memory_key(subj, summ, None),
                soft_key(summ) if summ else "",
            ):
                if not key or key in ("|", "||"):
                    continue
                entry = self.memory.get(key)
                if entry is None and key == soft_key(summ):
                    entry = self.memory.get(f"|{soft_key(summ)}")
                if entry and entry.l2_id:
                    return self._result_from_leaf(
                        entry.l2_id,
                        layer="memory",
                        confidence=float(entry.confidence or 1.0),
                        reason=f"memory:{entry.source}",
                    )

            # core 记忆：跨月份/数量变体的同型摘要
            if summ:
                core_hit = self._match_core_memory(summ)
                if core_hit is not None:
                    return core_hit

        # 关键词规则（硬规则优先于相似记忆）
        hits: List[Tuple[int, Dict[str, Any]]] = []
        for compiled in self._compiled:
            if self._match_compiled(summ_norm, compiled):
                hits.append((compiled["priority"], compiled))

        if hits:
            best_pri = hits[0][0]
            top = [c for p, c in hits if p == best_pri]
            cand_ids = []
            for _, c in hits:
                lid = c.get("l3_id") or c.get("l2_id")
                if not lid:
                    lid = self.taxonomy.resolve_leaf_id(c["raw"].get("二级科目"))
                if lid:
                    cand_ids.append(lid)
            primary = top[0]
            leaf = primary.get("l3_id") or primary.get("l2_id")
            if not leaf:
                leaf = self.taxonomy.resolve_leaf_id(primary["raw"].get("二级科目"))
            if not leaf:
                leaf = self.taxonomy.pending_l2_id
            # 若只有 l2 且该 l2 有三级、规则未指定 l3：三级留空（大类命中）
            conf = 0.9 if primary["priority"] >= 200 else (0.75 if primary["priority"] >= 100 else 0.65)
            layer = "rule_hi" if primary["priority"] >= 200 else "rule"
            if len(set(cand_ids)) > 1:
                conf = min(conf, 0.6)
                layer = "conflict" if conf < self.conf_review_threshold else layer
                reason = f"multi_hit:{','.join(sorted(set(cand_ids))[:5])}"
            else:
                reason = f"rule:{primary['id']}"
            return self._result_from_leaf(
                leaf,
                layer=layer,
                confidence=conf,
                rule_id=str(primary["id"]),
                reason=reason,
                candidates=list(dict.fromkeys(cand_ids)),
            )

        # 正则模式（如「代xx付款」中间插词）
        if summ:
            for cre, leaf_id, conf, reason in _SUMMARY_PATTERNS:
                if cre.search(summ) or cre.search(summ_norm):
                    # 叶子若已停用则跳过
                    if leaf_id in self.taxonomy.level2 or leaf_id in self.taxonomy.level3:
                        return self._result_from_leaf(
                            leaf_id,
                            layer="pattern",
                            confidence=conf,
                            reason=reason,
                        )

        # 相似记忆：用已确认/已学摘要泛化到相近未确认摘要
        if summ and self.memory is not None:
            sim_hit = self._match_similar_memory(summ)
            if sim_hit is not None:
                return sim_hit

        # 科目映射（可为 L2 或 L3 id）
        if subj:
            mapped = self.subject_map.get(subj) or self.subject_map.get(subj_norm)
            if not mapped:
                mapped = self.taxonomy.resolve_leaf_id(subj)
            if mapped:
                return self._result_from_leaf(
                    mapped,
                    layer="subject_map",
                    confidence=0.62,
                    reason=f"subject:{subj}",
                )

        # 供应商众数
        if self.vendor_stats is not None and vend_norm:
            voted = self.vendor_stats.vote(
                vend_norm,
                min_count=self.vendor_min_count,
                min_ratio=self.vendor_min_ratio,
            )
            if voted:
                leaf_id, ratio, total = voted
                return self._result_from_leaf(
                    leaf_id,
                    layer="vendor",
                    confidence=min(0.7, 0.5 + ratio * 0.3),
                    reason=f"vendor_vote:{ratio:.0%}/{total}",
                )

        return self._result_from_leaf(
            self.taxonomy.pending_l2_id,
            layer="pending",
            confidence=0.2,
            reason="no_match",
        )


    def classify_names(
        self, subject: Any = None, summary: Any = None, vendor: Any = None
    ) -> Tuple[str, str]:
        r = self.classify(subject, summary, vendor)
        return r.l1_name, r.l2_name

    def classify_names_triple(
        self, subject: Any = None, summary: Any = None, vendor: Any = None
    ) -> Tuple[str, str, str]:
        r = self.classify(subject, summary, vendor)
        return r.l1_name, r.l2_name, r.l3_name

    @staticmethod
    def _triple_key(
        subject: Any = None, summary: Any = None, vendor: Any = None
    ) -> Tuple[str, str, str]:
        subj = "" if subject is None else str(subject).strip()
        summ = "" if summary is None else str(summary).strip()
        vend = "" if vendor is None else str(vendor).strip()
        if subj.lower() == "nan":
            subj = ""
        if summ.lower() == "nan":
            summ = ""
        if vend.lower() == "nan":
            vend = ""
        return subj, summ, vend

    def classify_many(
        self,
        rows: List[Tuple[Any, Any, Any]],
    ) -> List[ClassifyResult]:
        """
        批量分类：对相同 (科目, 摘要, 供应商) 只算一次。
        rows: [(subject, summary, vendor), ...]
        """
        cache: Dict[Tuple[str, str, str], ClassifyResult] = {}
        out: List[ClassifyResult] = []
        for subject, summary, vendor in rows:
            key = self._triple_key(subject, summary, vendor)
            if key not in cache:
                cache[key] = self.classify(key[0], key[1], key[2])
            out.append(cache[key])
        return out

    def classify_unique_map(
        self,
        triples: List[Tuple[Any, Any, Any]],
    ) -> Dict[Tuple[str, str, str], ClassifyResult]:
        """对去重后的三元组建缓存表。"""
        cache: Dict[Tuple[str, str, str], ClassifyResult] = {}
        for subject, summary, vendor in triples:
            key = self._triple_key(subject, summary, vendor)
            if key not in cache:
                cache[key] = self.classify(key[0], key[1], key[2])
        return cache
