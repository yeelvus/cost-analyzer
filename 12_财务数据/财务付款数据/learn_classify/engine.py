# -*- coding: utf-8 -*-
"""
多层分类引擎（可解释）：支持一级/二级/三级。

记忆与规则的 leaf_id 可为 L2 或 L3；有三级时优先落到三级。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .normalize import memory_key, normalize_text, soft_key
from .store import MemoryStore, VendorStats, load_rules
from .taxonomy import Taxonomy


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

        # 关键词规则
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
