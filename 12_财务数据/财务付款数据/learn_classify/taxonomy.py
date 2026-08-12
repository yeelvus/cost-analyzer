# -*- coding: utf-8 -*-
"""分类树：一级 / 二级 / 三级；稳定 ID + 显示名。"""

from __future__ import annotations

import json
import re
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


@dataclass
class Level3:
    id: str
    name: str
    parent: str  # level2 id
    order: int = 0
    active: bool = True
    aliases: List[str] = field(default_factory=list)
    notes: str = ""


class Taxonomy:
    def __init__(self, data: Dict[str, Any]):
        self.version: str = str(data.get("version", "0.0.0"))
        self.description: str = str(data.get("description", ""))
        self.granularity = data.get("granularity") or {}
        self.level1: Dict[str, Level1] = {}
        self.level2: Dict[str, Level2] = {}
        self.level3: Dict[str, Level3] = {}
        self._name_to_l1: Dict[str, str] = {}
        self._name_to_l2: Dict[str, str] = {}
        self._name_to_l3: Dict[str, str] = {}

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
            self._register_names(self._name_to_l2, l2.name, l2.aliases, l2.id)

        for item in data.get("level3", []) or []:
            l3 = Level3(
                id=str(item["id"]).strip(),
                name=str(item["name"]).strip(),
                parent=str(item["parent"]).strip(),
                order=int(item.get("order", 0)),
                active=bool(item.get("active", True)),
                aliases=[str(a).strip() for a in item.get("aliases", []) if str(a).strip()],
                notes=str(item.get("notes", "")),
            )
            if l3.parent not in self.level2:
                raise ValueError(f"level3 {l3.id} parent(L2) 不存在: {l3.parent}")
            self.level3[l3.id] = l3
            self._register_names(self._name_to_l3, l3.name, l3.aliases, l3.id)

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
        # 可选默认三级
        self.default_l3_id: str = str(data.get("default_l3_id") or "")
        self.pending_l3_id: str = str(data.get("pending_l3_id") or "")

    @staticmethod
    def _register_names(mapping: Dict[str, str], name: str, aliases: List[str], oid: str) -> None:
        for label in [name, *aliases]:
            key = label.strip()
            if not key:
                continue
            mapping[key] = oid
            mapping[key.lower()] = oid

    def _find_id_by_name_l2(self, name: str) -> Optional[str]:
        for l2 in self.level2.values():
            if l2.name == name:
                return l2.id
        return None

    def children_l2(self, l1_id: str) -> List[Level2]:
        return sorted(
            [x for x in self.level2.values() if x.parent == l1_id and x.active],
            key=lambda z: (z.order, z.name),
        )

    def children_l3(self, l2_id: str) -> List[Level3]:
        return sorted(
            [x for x in self.level3.values() if x.parent == l2_id and x.active],
            key=lambda z: (z.order, z.name),
        )

    def resolve_l3_id(self, value: Any) -> Optional[str]:
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text in self.level3:
            return text
        return self._name_to_l3.get(text) or self._name_to_l3.get(text.lower())

    def resolve_l2_id(self, value: Any) -> Optional[str]:
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

    def resolve_leaf_id(self, value: Any) -> Optional[str]:
        """解析到最细 ID（优先三级，再二级）。"""
        return self.resolve_l3_id(value) or self.resolve_l2_id(value)

    def resolve_triple(
        self,
        level1: Any = None,
        level2: Any = None,
        level3: Any = None,
    ) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """解析 (l1_id, l2_id, l3_id)。以最细级为准向上补全。"""
        l3_id = self.resolve_l3_id(level3) if level3 is not None else None
        l2_id = self.resolve_l2_id(level2) if level2 is not None else None
        l1_id = self.resolve_l1_id(level1) if level1 is not None else None

        if l3_id:
            l2_id = self.level3[l3_id].parent
            l1_id = self.level2[l2_id].parent if l2_id in self.level2 else l1_id
        elif l2_id:
            l1_id = self.level2[l2_id].parent
        return l1_id, l2_id, l3_id

    def resolve_pair(
        self, level1: Any = None, level2: Any = None
    ) -> Tuple[Optional[str], Optional[str]]:
        """兼容旧接口：返回 (l1_id, l2_id)。若 level2 实际是三级名，会解析到其父二级。"""
        # 先试三级
        l3 = self.resolve_l3_id(level2)
        if l3:
            l2 = self.level3[l3].parent
            l1 = self.level2[l2].parent
            return l1, l2
        l1, l2, _ = self.resolve_triple(level1, level2, None)
        return l1, l2

    def display(self, leaf_or_l2_id: Optional[str]) -> Tuple[str, str]:
        """兼容旧接口：返回 (一级名, 二级名)。若是三级，二级名为父二级。"""
        l1n, l2n, _ = self.display_triple(leaf_or_l2_id)
        return l1n, l2n

    def display_triple(self, leaf_id: Optional[str]) -> Tuple[str, str, str]:
        """返回 (一级名, 二级名, 三级名)。"""
        if not leaf_id:
            return "", "", ""
        if leaf_id in self.level3:
            l3 = self.level3[leaf_id]
            l2 = self.level2.get(l3.parent)
            l1 = self.level1.get(l2.parent) if l2 else None
            return (
                l1.name if l1 else "",
                l2.name if l2 else "",
                l3.name,
            )
        if leaf_id in self.level2:
            l2 = self.level2[leaf_id]
            l1 = self.level1.get(l2.parent)
            return (l1.name if l1 else "", l2.name, "")
        return "", "", ""

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

    def _unique_node_id(self, prefix: str, name: str) -> str:
        """生成稳定可读 ID：L2_auto_名称；冲突则加序号。"""
        raw = re.sub(r"\s+", "", str(name or "").strip())
        raw = re.sub(r'[\\/:*?"<>|]', "", raw)
        raw = raw[:28] if raw else "new"
        base = f"{prefix}_{raw}"
        # ID 中尽量 ASCII；保留中文亦可
        existing = set(self.level1) | set(self.level2) | set(self.level3)
        if base not in existing:
            return base
        n = 2
        while f"{base}_{n}" in existing:
            n += 1
        return f"{base}_{n}"

    def ensure_l1(
        self,
        name: str,
        *,
        notes: str = "auto from learn_pool",
    ) -> Tuple[str, bool]:
        """确保一级存在。返回 (l1_id, created)。"""
        text = str(name or "").strip()
        if not text:
            raise ValueError("一级科目名称为空")
        hit = self.resolve_l1_id(text)
        if hit:
            return hit, False
        lid = self._unique_node_id("L1_auto", text)
        order = max((x.order for x in self.level1.values()), default=0) + 1
        node = Level1(id=lid, name=text, order=order, active=True, notes=notes)
        self.level1[lid] = node
        self._name_to_l1[text] = lid
        self._name_to_l1[text.lower()] = lid
        return lid, True

    def ensure_l2(
        self,
        name: str,
        parent_l1_id: str,
        *,
        notes: str = "auto from learn_pool",
    ) -> Tuple[str, bool]:
        """确保二级存在于指定一级下。返回 (l2_id, created)。"""
        text = str(name or "").strip()
        if not text:
            raise ValueError("二级科目名称为空")
        if parent_l1_id not in self.level1:
            raise ValueError(f"一级不存在: {parent_l1_id}")
        hit = self.resolve_l2_id(text)
        if hit:
            # 已存在但挂在别的一级：不挪动，直接复用（避免破坏历史）
            return hit, False
        lid = self._unique_node_id("L2_auto", text)
        siblings = [x for x in self.level2.values() if x.parent == parent_l1_id]
        order = max((x.order for x in siblings), default=0) + 1
        node = Level2(
            id=lid,
            name=text,
            parent=parent_l1_id,
            order=order,
            active=True,
            aliases=[],
            notes=notes,
        )
        self.level2[lid] = node
        self._register_names(self._name_to_l2, text, [], lid)
        return lid, True

    def ensure_l3(
        self,
        name: str,
        parent_l2_id: str,
        *,
        notes: str = "auto from learn_pool",
    ) -> Tuple[str, bool]:
        """确保三级存在。返回 (l3_id, created)。"""
        text = str(name or "").strip()
        if not text:
            raise ValueError("三级科目名称为空")
        if parent_l2_id not in self.level2:
            raise ValueError(f"二级不存在: {parent_l2_id}")
        hit = self.resolve_l3_id(text)
        if hit:
            return hit, False
        lid = self._unique_node_id("L3_auto", text)
        siblings = [x for x in self.level3.values() if x.parent == parent_l2_id]
        order = max((x.order for x in siblings), default=0) + 1
        node = Level3(
            id=lid,
            name=text,
            parent=parent_l2_id,
            order=order,
            active=True,
            aliases=[],
            notes=notes,
        )
        self.level3[lid] = node
        self._register_names(self._name_to_l3, text, [], lid)
        return lid, True

    def ensure_triple(
        self,
        level1: Any = None,
        level2: Any = None,
        level3: Any = None,
        *,
        default_l1_id: str = "L1_mgmt",
        notes: str = "auto from learn_pool",
    ) -> Tuple[str, str, Optional[str], List[str]]:
        """
        解析三元组；缺失则自动登记。
        返回 (l1_id, l2_id, l3_id|None, created_log)。
        created_log 如 ["L1:新类", "L2:xxx"]。
        """
        created: List[str] = []
        l1_name = str(level1 or "").strip()
        l2_name = str(level2 or "").strip()
        l3_name = str(level3 or "").strip()

        # 已能完整解析则直接返回
        l1_id, l2_id, l3_id = self.resolve_triple(
            l1_name or None, l2_name or None, l3_name or None
        )
        if l3_id or l2_id:
            if l3_id and not l2_id:
                l2_id = self.level3[l3_id].parent
            if l2_id and not l1_id:
                l1_id = self.level2[l2_id].parent
            return l1_id or "", l2_id or "", l3_id, created

        # 仅给了三级名：当作二级登记（常见误填）
        if not l2_name and l3_name:
            l2_name = l3_name
            l3_name = ""

        if not l2_name:
            raise ValueError(
                f"无法自动登记：缺少二级科目名 l1={l1_name!r} l2={l2_name!r} l3={l3_name!r}"
            )

        # 一级
        if l1_name:
            l1_id, c1 = self.ensure_l1(l1_name, notes=notes)
            if c1:
                created.append(f"L1:{l1_name}")
        else:
            l1_id = default_l1_id if default_l1_id in self.level1 else None
            if not l1_id:
                # 回退任意一级
                l1_id = next(iter(self.level1), "")
            if not l1_id:
                l1_id, c1 = self.ensure_l1("6_管理费", notes=notes)
                if c1:
                    created.append("L1:6_管理费")

        l2_id, c2 = self.ensure_l2(l2_name, l1_id, notes=notes)
        if c2:
            created.append(f"L2:{l2_name}->{self.level1[l1_id].name}")

        l3_out: Optional[str] = None
        if l3_name:
            l3_out, c3 = self.ensure_l3(l3_name, l2_id, notes=notes)
            if c3:
                created.append(f"L3:{l3_name}")

        return l1_id, l2_id, l3_out, created

    def is_active_leaf(self, leaf_id: str) -> bool:
        if leaf_id in self.level3:
            l3 = self.level3[leaf_id]
            if not l3.active:
                return False
            return self.is_active_l2(l3.parent)
        return self.is_active_l2(leaf_id)

    def pending_display(self) -> Tuple[str, str]:
        return self.display(self.pending_l2_id)

    def pending_display_triple(self) -> Tuple[str, str, str]:
        return self.display_triple(self.pending_l2_id)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": self.version,
            "description": self.description,
            "granularity": self.granularity,
            "default_l2_id": self.default_l2_id,
            "pending_l2_id": self.pending_l2_id,
            "default_l3_id": self.default_l3_id,
            "pending_l3_id": self.pending_l3_id,
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
            "level3": [
                {
                    "id": x.id,
                    "name": x.name,
                    "parent": x.parent,
                    "order": x.order,
                    "active": x.active,
                    "aliases": x.aliases,
                    "notes": x.notes,
                }
                for x in sorted(self.level3.values(), key=lambda z: (z.order, z.id))
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
        data = deepcopy(self.to_dict())
        for old_id, new_id in id_map.items():
            if old_id in self.level2 and new_id in self.level2:
                for item in data["level2"]:
                    if item["id"] == old_id:
                        item["active"] = False
                        item["notes"] = (item.get("notes") or "") + f" | migrated→{new_id}"
            if old_id in self.level3 and new_id in self.level3:
                for item in data["level3"]:
                    if item["id"] == old_id:
                        item["active"] = False
                        item["notes"] = (item.get("notes") or "") + f" | migrated→{new_id}"
        return Taxonomy(data)


def apply_migration_to_taxonomy_file(
    taxonomy: Taxonomy,
    migration: Dict[str, Any],
) -> Taxonomy:
    data = taxonomy.to_dict()
    if migration.get("to_version"):
        data["version"] = str(migration["to_version"])

    by_l2 = {x["id"]: x for x in data["level2"]}
    by_l1 = {x["id"]: x for x in data["level1"]}
    by_l3 = {x["id"]: x for x in data.get("level3") or []}

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

    for item in migration.get("add_level3", []) or []:
        by_l3[item["id"]] = {
            "id": item["id"],
            "name": item["name"],
            "parent": item["parent"],
            "order": item.get("order", 99),
            "active": item.get("active", True),
            "aliases": item.get("aliases", []),
            "notes": item.get("notes", ""),
        }

    for lid, new_name in (migration.get("rename") or {}).items():
        if lid in by_l2:
            by_l2[lid]["name"] = new_name
        if lid in by_l1:
            by_l1[lid]["name"] = new_name
        if lid in by_l3:
            by_l3[lid]["name"] = new_name

    for lid, new_parent in (migration.get("reparent") or {}).items():
        if lid in by_l2:
            by_l2[lid]["parent"] = new_parent
        if lid in by_l3:
            by_l3[lid]["parent"] = new_parent

    # 可选：按 id 覆盖 level2 字段（order / notes / aliases 等）
    for lid, fields in (migration.get("patch_level2") or {}).items():
        if lid not in by_l2 or not isinstance(fields, dict):
            continue
        for key, val in fields.items():
            if key == "id":
                continue
            by_l2[lid][key] = val

    if migration.get("taxonomy_description") is not None:
        data["description"] = str(migration["taxonomy_description"])

    for lid in migration.get("deactivate") or []:
        if lid in by_l2:
            by_l2[lid]["active"] = False
        if lid in by_l1:
            by_l1[lid]["active"] = False
        if lid in by_l3:
            by_l3[lid]["active"] = False

    if migration.get("pending_l2_id"):
        data["pending_l2_id"] = migration["pending_l2_id"]
    if migration.get("default_l2_id"):
        data["default_l2_id"] = migration["default_l2_id"]

    data["level1"] = list(by_l1.values())
    data["level2"] = list(by_l2.values())
    data["level3"] = list(by_l3.values())
    return Taxonomy(data)
