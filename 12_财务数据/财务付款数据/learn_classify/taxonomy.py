# -*- coding: utf-8 -*-
"""分类树：稳定 ID + 可改显示名，支持版本与迁移。"""

from __future__ import annotations

import json
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class Level1:
    id: str
    name: str
    order: int = 0
    active: bool = True
    notes: str = ""


@dataclass
class Level2:
    id: str
    name: str
    parent: str
    order: int = 0
    active: bool = True
    aliases: List[str] = field(default_factory=list)
    notes: str = ""


class Taxonomy:
    def __init__(self, data: Dict[str, Any]):
        self.version: str = str(data.get("version", "0.0.0"))
        self.description: str = str(data.get("description", ""))
        self.level1: Dict[str, Level1] = {}
        self.level2: Dict[str, Level2] = {}
        self._name_to_l1: Dict[str, str] = {}
        self._name_to_l2: Dict[str, str] = {}  # name or alias -> id

        for item in data.get("level1", []):
            l1 = Level1(
                id=str(item["id"]).strip(),
                name=str(item["name"]).strip(),
                order=int(item.get("order", 0)),
                active=bool(item.get("active", True)),
                notes=str(item.get("notes", "")),
            )
            self.level1[l1.id] = l1
            self._name_to_l1[l1.name] = l1.id
            self._name_to_l1[l1.name.lower()] = l1.id

        for item in data.get("level2", []):
            l2 = Level2(
                id=str(item["id"]).strip(),
                name=str(item["name"]).strip(),
                parent=str(item["parent"]).strip(),
                order=int(item.get("order", 0)),
                active=bool(item.get("active", True)),
                aliases=[str(a).strip() for a in item.get("aliases", []) if str(a).strip()],
                notes=str(item.get("notes", "")),
            )
            if l2.parent not in self.level1:
                raise ValueError(f"level2 {l2.id} parent 不存在: {l2.parent}")
            self.level2[l2.id] = l2
            self._register_l2_names(l2)

        self.default_l2_id: str = str(
            data.get("default_l2_id")
            or self._find_id_by_name_l2("待复核")
            or self._find_id_by_name_l2("其他")
            or next(iter(self.level2), "")
        )
        self.pending_l2_id: str = str(
            data.get("pending_l2_id")
            or self._find_id_by_name_l2("待复核")
            or self.default_l2_id
        )

    def _register_l2_names(self, l2: Level2) -> None:
        for label in [l2.name, *l2.aliases]:
            key = label.strip()
            if not key:
                continue
            self._name_to_l2[key] = l2.id
            self._name_to_l2[key.lower()] = l2.id

    def _find_id_by_name_l2(self, name: str) -> Optional[str]:
        for l2 in self.level2.values():
            if l2.name == name:
                return l2.id
        return None

    # ----- resolve -----
    def resolve_l2_id(self, value: Any) -> Optional[str]:
        """接受 level2 id / 显示名 / 别名。"""
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text in self.level2:
            return text
        return self._name_to_l2.get(text) or self._name_to_l2.get(text.lower())

    def resolve_l1_id(self, value: Any) -> Optional[str]:
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text in self.level1:
            return text
        return self._name_to_l1.get(text) or self._name_to_l1.get(text.lower())

    def resolve_pair(
        self, level1: Any = None, level2: Any = None
    ) -> Tuple[Optional[str], Optional[str]]:
        """解析到 (l1_id, l2_id)。若仅给二级，自动带出一级。"""
        l2_id = self.resolve_l2_id(level2) if level2 is not None else None
        l1_id = self.resolve_l1_id(level1) if level1 is not None else None
        if l2_id:
            parent = self.level2[l2_id].parent
            if l1_id and l1_id != parent:
                # 以二级为准（业务纠正时更可靠）
                l1_id = parent
            else:
                l1_id = parent
        return l1_id, l2_id

    def display(self, l2_id: Optional[str]) -> Tuple[str, str]:
        if not l2_id or l2_id not in self.level2:
            return "", ""
        l2 = self.level2[l2_id]
        l1 = self.level1.get(l2.parent)
        return (l1.name if l1 else ""), l2.name

    def display_l1(self, l1_id: Optional[str]) -> str:
        if not l1_id or l1_id not in self.level1:
            return ""
        return self.level1[l1_id].name

    def is_active_l2(self, l2_id: str) -> bool:
        l2 = self.level2.get(l2_id)
        if not l2 or not l2.active:
            return False
        l1 = self.level1.get(l2.parent)
        return bool(l1 and l1.active)

    def pending_display(self) -> Tuple[str, str]:
        return self.display(self.pending_l2_id)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": self.version,
            "description": self.description,
            "default_l2_id": self.default_l2_id,
            "pending_l2_id": self.pending_l2_id,
            "level1": [
                {
                    "id": x.id,
                    "name": x.name,
                    "order": x.order,
                    "active": x.active,
                    "notes": x.notes,
                }
                for x in sorted(self.level1.values(), key=lambda z: (z.order, z.id))
            ],
            "level2": [
                {
                    "id": x.id,
                    "name": x.name,
                    "parent": x.parent,
                    "order": x.order,
                    "active": x.active,
                    "aliases": x.aliases,
                    "notes": x.notes,
                }
                for x in sorted(self.level2.values(), key=lambda z: (z.order, z.id))
            ],
        }

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: Path) -> "Taxonomy":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(data)

    def apply_id_map(self, id_map: Dict[str, str]) -> "Taxonomy":
        """
        按 old_id -> new_id 转移二级（及可选一级）。
        new_id 必须已在当前 taxonomy 中；仅做校验用辅助。
        真正改树请编辑 taxonomy.json 后用 migrate 改记忆库。
        """
        data = deepcopy(self.to_dict())
        # 标记 inactive old ids if mapped away
        for old_id, new_id in id_map.items():
            if old_id in self.level2 and new_id in self.level2:
                for item in data["level2"]:
                    if item["id"] == old_id:
                        item["active"] = False
                        item["notes"] = (item.get("notes") or "") + f" | migrated→{new_id}"
        return Taxonomy(data)


def apply_migration_to_taxonomy_file(
    taxonomy: Taxonomy,
    migration: Dict[str, Any],
) -> Taxonomy:
    """
    应用迁移描述：
    {
      "to_version": "1.1.0",
      "rename": {"L2_xxx": "新显示名"},
      "reparent": {"L2_xxx": "L1_new"},
      "add_level2": [...],
      "deactivate": ["L2_old"],
      "id_map": {"L2_old": "L2_new"}  # 仅记录，树结构用 add/reparent
    }
    """
    data = taxonomy.to_dict()
    if migration.get("to_version"):
        data["version"] = str(migration["to_version"])

    by_l2 = {x["id"]: x for x in data["level2"]}
    by_l1 = {x["id"]: x for x in data["level1"]}

    for item in migration.get("add_level1", []) or []:
        by_l1[item["id"]] = {
            "id": item["id"],
            "name": item["name"],
            "order": item.get("order", 99),
            "active": item.get("active", True),
            "notes": item.get("notes", ""),
        }

    for item in migration.get("add_level2", []) or []:
        by_l2[item["id"]] = {
            "id": item["id"],
            "name": item["name"],
            "parent": item["parent"],
            "order": item.get("order", 99),
            "active": item.get("active", True),
            "aliases": item.get("aliases", []),
            "notes": item.get("notes", ""),
        }

    for l2_id, new_name in (migration.get("rename") or {}).items():
        if l2_id in by_l2:
            by_l2[l2_id]["name"] = new_name
        if l2_id in by_l1:
            by_l1[l2_id]["name"] = new_name

    for l2_id, new_parent in (migration.get("reparent") or {}).items():
        if l2_id in by_l2:
            by_l2[l2_id]["parent"] = new_parent

    for l2_id in migration.get("deactivate") or []:
        if l2_id in by_l2:
            by_l2[l2_id]["active"] = False
        if l2_id in by_l1:
            by_l1[l2_id]["active"] = False

    if migration.get("pending_l2_id"):
        data["pending_l2_id"] = migration["pending_l2_id"]
    if migration.get("default_l2_id"):
        data["default_l2_id"] = migration["default_l2_id"]

    data["level1"] = list(by_l1.values())
    data["level2"] = list(by_l2.values())
    return Taxonomy(data)
