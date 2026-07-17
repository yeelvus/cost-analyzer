# -*- coding: utf-8 -*-
"""
多层分类引擎（可解释）：

  L0  特殊/不计入成本（规则 priority 最高段）
  L1  精确记忆（科目|摘要 [+|供应商]）
  L2  高优先级关键词规则
  L3  原始科目 → 默认二级映射
  L4  普通关键词规则
  L5  供应商历史众数
  L6  待复核（不再默认塞管理费）
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
    l1_name: str
    l2_name: str
    layer: str
    confidence: float
    rule_id: str = ""
    candidates: List[str] = field(default_factory=list)
    reason: str = ""

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
            # name -> l2_id
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
        return {
            "raw": rule,
            "id": str(rule.get("id") or rule.get("二级科目") or ""),
            "l2_id": str(rule.get("l2_id") or rule.get("target") or ""),
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

    def _result_from_l2(
        self,
        l2_id: str,
        layer: str,
        confidence: float,
        rule_id: str = "",
        reason: str = "",
        candidates: Optional[List[str]] = None,
    ) -> ClassifyResult:
        tax = self.taxonomy
        if not tax.is_active_l2(l2_id):
            l2_id = tax.pending_l2_id
            layer = "pending"
            confidence = min(confidence, 0.4)
            reason = (reason + " | inactive→pending").strip(" |")
        l1_name, l2_name = tax.display(l2_id)
        l1_id = tax.level2[l2_id].parent if l2_id in tax.level2 else ""
        return ClassifyResult(
            l1_id=l1_id,
            l2_id=l2_id,
            l1_name=l1_name,
            l2_name=l2_name,
            layer=layer,
            confidence=confidence,
            rule_id=rule_id,
            candidates=candidates or [],
            reason=reason,
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

        # L1 精确记忆
        if self.memory is not None:
            for key in (
                memory_key(subj, summ, vend),
                memory_key(subj, summ, None),
                soft_key(summ) if summ else "",
            ):
                if not key or key in ("|", "||"):
                    continue
                # soft_key 单独查：用 summary-only 索引
                entry = self.memory.get(key)
                if entry is None and key == soft_key(summ):
                    # 也尝试以 |summary 形式
                    entry = self.memory.get(f"|{soft_key(summ)}")
                if entry and entry.l2_id:
                    return self._result_from_l2(
                        entry.l2_id,
                        layer="memory",
                        confidence=float(entry.confidence or 1.0),
                        reason=f"memory:{entry.source}",
                    )

        # 关键词规则（按 priority）
        hits: List[Tuple[int, Dict[str, Any]]] = []
        for compiled in self._compiled:
            if self._match_compiled(summ_norm, compiled):
                hits.append((compiled["priority"], compiled))

        if hits:
            # 同最高优先级若多个 → conflict 降置信
            best_pri = hits[0][0]
            top = [c for p, c in hits if p == best_pri]
            # 也收集全部候选
            cand_ids = []
            for _, c in hits:
                lid = c["l2_id"] or self.taxonomy.resolve_l2_id(c["raw"].get("二级科目"))
                if lid:
                    cand_ids.append(lid)
            primary = top[0]
            l2_id = primary["l2_id"] or self.taxonomy.resolve_l2_id(primary["raw"].get("二级科目"))
            if not l2_id:
                l2_id = self.taxonomy.pending_l2_id
            conf = 0.9 if primary["priority"] >= 200 else (0.75 if primary["priority"] >= 100 else 0.65)
            layer = "rule_hi" if primary["priority"] >= 200 else "rule"
            if len(set(cand_ids)) > 1:
                conf = min(conf, 0.6)
                layer = "conflict" if conf < self.conf_review_threshold else layer
                reason = f"multi_hit:{','.join(sorted(set(cand_ids))[:5])}"
            else:
                reason = f"rule:{primary['id']}"
            return self._result_from_l2(
                l2_id,
                layer=layer,
                confidence=conf,
                rule_id=str(primary["id"]),
                reason=reason,
                candidates=list(dict.fromkeys(cand_ids)),
            )

        # L3 科目映射
        if subj:
            mapped = self.subject_map.get(subj) or self.subject_map.get(subj_norm)
            if not mapped:
                # 允许用 taxonomy 别名
                mapped = self.taxonomy.resolve_l2_id(subj)
            if mapped:
                return self._result_from_l2(
                    mapped,
                    layer="subject_map",
                    confidence=0.62,
                    reason=f"subject:{subj}",
                )

        # L5 供应商众数
        if self.vendor_stats is not None and vend_norm:
            voted = self.vendor_stats.vote(
                vend_norm,
                min_count=self.vendor_min_count,
                min_ratio=self.vendor_min_ratio,
            )
            if voted:
                l2_id, ratio, total = voted
                return self._result_from_l2(
                    l2_id,
                    layer="vendor",
                    confidence=min(0.7, 0.5 + ratio * 0.3),
                    reason=f"vendor_vote:{ratio:.0%}/{total}",
                )

        # L6 待复核
        return self._result_from_l2(
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
