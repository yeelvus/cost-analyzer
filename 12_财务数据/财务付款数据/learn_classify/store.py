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

    def line_count(self) -> int:
        if not self.path.exists():
            return 0
        n = 0
        with open(self.path, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    n += 1
        return n

    def compact(
        self,
        *,
        backup: bool = True,
        drop_empty_summary: bool = True,
    ) -> Dict[str, Any]:
        """
        压缩 memory.jsonl：只保留每个 key 的最终条目（加载时后写覆盖）。

        - 去掉历史追加造成的重复行（主要体积来源）
        - 可选丢弃无摘要的空键
        返回统计；写入前可备份为 memory.jsonl.bak_时间戳
        """
        before_lines = self.line_count()
        before_keys = len(self.entries)
        before_bytes = self.path.stat().st_size if self.path.exists() else 0

        kept: Dict[str, MemoryEntry] = {}
        dropped_empty = 0
        for key, e in self.entries.items():
            if drop_empty_summary and not str(e.summary or "").strip() and not str(
                e.key or ""
            ).replace("|", "").strip():
                dropped_empty += 1
                continue
            if not key or not e.l2_id:
                dropped_empty += 1
                continue
            kept[key] = e

        self.entries = kept
        bak_path = ""
        if backup and self.path.exists() and before_lines > 0:
            ts = time.strftime("%Y%m%d_%H%M%S")
            bak_path = str(self.path.with_name(f"{self.path.name}.bak_{ts}"))
            import shutil

            shutil.copy2(self.path, bak_path)

        self._store.rewrite(asdict(e) for e in sorted(self.entries.values(), key=lambda x: x.key))
        after_bytes = self.path.stat().st_size if self.path.exists() else 0
        return {
            "before_lines": before_lines,
            "after_lines": len(self.entries),
            "unique_keys_before_load": before_keys,
            "unique_keys_after": len(self.entries),
            "dropped_empty": dropped_empty,
            "before_bytes": before_bytes,
            "after_bytes": after_bytes,
            "saved_bytes": max(0, before_bytes - after_bytes),
            "backup": bak_path,
        }


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

    def line_count(self) -> int:
        if not self._store.path.exists():
            return 0
        n = 0
        with open(self._store.path, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    n += 1
        return n

    def compact(
        self,
        *,
        backup: bool = True,
        keep_last: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        压缩纠错日志。

        - 默认：对 (subject, summary, vendor, new_l2_id) 去重，只留最后一条
        - keep_last: 再截断为最近 N 条（按文件顺序，末尾为新）
        """
        path = self._store.path
        before_lines = self.line_count()
        before_bytes = path.stat().st_size if path.exists() else 0
        rows = self._store.read_all()

        # 后写覆盖：按指纹保留最后出现
        latest: Dict[str, Dict[str, Any]] = {}
        order: List[str] = []
        for row in rows:
            fp = "|".join(
                [
                    str(row.get("subject", "")).strip(),
                    str(row.get("summary", "")).strip(),
                    str(row.get("vendor", "")).strip(),
                    str(row.get("new_l2_id", "")).strip(),
                ]
            )
            if fp not in latest:
                order.append(fp)
            latest[fp] = row
        out = [latest[fp] for fp in order]

        if keep_last is not None and keep_last > 0 and len(out) > keep_last:
            out = out[-keep_last:]

        bak_path = ""
        if backup and path.exists() and before_lines > 0:
            ts = time.strftime("%Y%m%d_%H%M%S")
            bak_path = str(path.with_name(f"{path.name}.bak_{ts}"))
            import shutil

            shutil.copy2(path, bak_path)

        self._store.rewrite(out)
        after_bytes = path.stat().st_size if path.exists() else 0
        return {
            "before_lines": before_lines,
            "after_lines": len(out),
            "before_bytes": before_bytes,
            "after_bytes": after_bytes,
            "saved_bytes": max(0, before_bytes - after_bytes),
            "backup": bak_path,
        }


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


def compact_data_dir(
    data_dir: Path,
    *,
    memory: bool = True,
    corrections: bool = True,
    backup: bool = True,
    corrections_keep_last: Optional[int] = None,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """压缩 data 目录下的 memory / corrections jsonl。"""
    data_dir = Path(data_dir)
    report: Dict[str, Any] = {"data_dir": str(data_dir), "dry_run": dry_run}

    if memory:
        mem = MemoryStore(data_dir / "memory.jsonl")
        if dry_run:
            lines = mem.line_count()
            report["memory"] = {
                "before_lines": lines,
                "unique_keys": len(mem.entries),
                "would_remove_lines": max(0, lines - len(mem.entries)),
                "before_bytes": mem.path.stat().st_size if mem.path.exists() else 0,
            }
        else:
            report["memory"] = mem.compact(backup=backup)

    if corrections:
        corr = CorrectionStore(data_dir / "corrections.jsonl")
        if dry_run:
            rows = corr._store.read_all()
            fps = set()
            for row in rows:
                fps.add(
                    "|".join(
                        [
                            str(row.get("subject", "")).strip(),
                            str(row.get("summary", "")).strip(),
                            str(row.get("vendor", "")).strip(),
                            str(row.get("new_l2_id", "")).strip(),
                        ]
                    )
                )
            after = len(fps)
            if corrections_keep_last and after > corrections_keep_last:
                after = corrections_keep_last
            report["corrections"] = {
                "before_lines": len(rows),
                "unique_fingerprints": len(fps),
                "would_keep": after,
                "before_bytes": corr._store.path.stat().st_size
                if corr._store.path.exists()
                else 0,
            }
        else:
            report["corrections"] = corr.compact(
                backup=backup, keep_last=corrections_keep_last
            )

    return report
