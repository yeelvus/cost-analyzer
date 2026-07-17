# -*- coding: utf-8 -*-
"""持久化：记忆库、纠错日志、供应商标签、规则。"""

from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


@dataclass
class MemoryEntry:
    key: str
    l2_id: str
    subject: str = ""
    summary: str = ""
    vendor: str = ""
    source: str = "manual"  # manual | correction | import | rule_seed
    confidence: float = 1.0
    updated_at: str = field(default_factory=_now)
    hit_count: int = 0
    notes: str = ""


@dataclass
class CorrectionEvent:
    ts: str
    subject: str
    summary: str
    vendor: str
    old_l2_id: str
    new_l2_id: str
    old_l1_name: str = ""
    old_l2_name: str = ""
    new_l1_name: str = ""
    new_l2_name: str = ""
    source: str = "review"
    notes: str = ""


class JsonlStore:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def append(self, obj: Dict[str, Any]) -> None:
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(json.dumps(obj, ensure_ascii=False) + "\n")

    def read_all(self) -> List[Dict[str, Any]]:
        if not self.path.exists():
            return []
        rows = []
        with open(self.path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
        return rows

    def rewrite(self, rows: Iterable[Dict[str, Any]]) -> None:
        with open(self.path, "w", encoding="utf-8") as f:
            for row in rows:
                f.write(json.dumps(row, ensure_ascii=False) + "\n")


class MemoryStore:
    """key -> MemoryEntry；jsonl 为追加日志，加载时后写覆盖前写。"""

    def __init__(self, path: Path):
        self.path = Path(path)
        self._store = JsonlStore(self.path)
        self.entries: Dict[str, MemoryEntry] = {}
        self.reload()

    def reload(self) -> None:
        self.entries = {}
        for row in self._store.read_all():
            key = str(row.get("key", "")).strip()
            l2_id = str(row.get("l2_id", "")).strip()
            if not key or not l2_id:
                continue
            self.entries[key] = MemoryEntry(
                key=key,
                l2_id=l2_id,
                subject=str(row.get("subject", "")),
                summary=str(row.get("summary", "")),
                vendor=str(row.get("vendor", "")),
                source=str(row.get("source", "manual")),
                confidence=float(row.get("confidence", 1.0)),
                updated_at=str(row.get("updated_at", _now())),
                hit_count=int(row.get("hit_count", 0)),
                notes=str(row.get("notes", "")),
            )

    def get(self, key: str) -> Optional[MemoryEntry]:
        return self.entries.get(key)

    def upsert(self, entry: MemoryEntry, persist: bool = True) -> None:
        entry.updated_at = _now()
        self.entries[entry.key] = entry
        if persist:
            self._store.append(asdict(entry))

    def apply_id_map(self, id_map: Dict[str, str]) -> int:
        """把记忆中的 l2_id 按 map 转移，返回变更条数。"""
        changed = 0
        new_entries = {}
        for key, e in self.entries.items():
            new_id = id_map.get(e.l2_id, e.l2_id)
            if new_id != e.l2_id:
                changed += 1
                e = MemoryEntry(**{**asdict(e), "l2_id": new_id, "updated_at": _now()})
            new_entries[key] = e
        self.entries = new_entries
        self._store.rewrite(asdict(e) for e in self.entries.values())
        return changed


class CorrectionStore:
    def __init__(self, path: Path):
        self._store = JsonlStore(path)

    def append(self, event: CorrectionEvent) -> None:
        self._store.append(asdict(event))

    def read_all(self) -> List[CorrectionEvent]:
        rows = []
        for row in self._store.read_all():
            rows.append(
                CorrectionEvent(
                    ts=str(row.get("ts", "")),
                    subject=str(row.get("subject", "")),
                    summary=str(row.get("summary", "")),
                    vendor=str(row.get("vendor", "")),
                    old_l2_id=str(row.get("old_l2_id", "")),
                    new_l2_id=str(row.get("new_l2_id", "")),
                    old_l1_name=str(row.get("old_l1_name", "")),
                    old_l2_name=str(row.get("old_l2_name", "")),
                    new_l1_name=str(row.get("new_l1_name", "")),
                    new_l2_name=str(row.get("new_l2_name", "")),
                    source=str(row.get("source", "review")),
                    notes=str(row.get("notes", "")),
                )
            )
        return rows

    def apply_id_map(self, id_map: Dict[str, str]) -> int:
        rows = self._store.read_all()
        changed = 0
        out = []
        for row in rows:
            for field in ("old_l2_id", "new_l2_id"):
                old = str(row.get(field, ""))
                if old in id_map:
                    row[field] = id_map[old]
                    changed += 1
            out.append(row)
        self._store.rewrite(out)
        return changed


class VendorStats:
    """供应商 → {l2_id: count}，用于众数投票。"""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.data: Dict[str, Dict[str, int]] = {}
        self.reload()

    def reload(self) -> None:
        if self.path.exists():
            with open(self.path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            self.data = {str(k): {str(a): int(b) for a, b in v.items()} for k, v in raw.items()}
        else:
            self.data = {}

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def observe(self, vendor_key: str, l2_id: str, n: int = 1) -> None:
        if not vendor_key or not l2_id:
            return
        bucket = self.data.setdefault(vendor_key, {})
        bucket[l2_id] = bucket.get(l2_id, 0) + n

    def vote(
        self, vendor_key: str, min_count: int = 3, min_ratio: float = 0.7
    ) -> Optional[tuple]:
        """返回 (l2_id, ratio, total) 或 None。"""
        bucket = self.data.get(vendor_key) or {}
        total = sum(bucket.values())
        if total < min_count:
            return None
        best_id, best_n = max(bucket.items(), key=lambda x: x[1])
        ratio = best_n / total
        if ratio < min_ratio:
            return None
        return best_id, ratio, total

    def apply_id_map(self, id_map: Dict[str, str]) -> int:
        changed = 0
        new_data: Dict[str, Dict[str, int]] = {}
        for vendor, bucket in self.data.items():
            nb: Dict[str, int] = {}
            for l2_id, n in bucket.items():
                nid = id_map.get(l2_id, l2_id)
                if nid != l2_id:
                    changed += 1
                nb[nid] = nb.get(nid, 0) + n
            new_data[vendor] = nb
        self.data = new_data
        self.save()
        return changed


def load_rules(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    if isinstance(data, dict) and "rules" in data:
        return list(data["rules"])
    return []


def save_rules(path: Path, rules: List[Dict[str, Any]], meta: Optional[Dict] = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": (meta or {}).get("version", "1"),
        "description": (meta or {}).get("description", "关键词规则，target 为 level2 稳定 ID"),
        "rules": rules,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
