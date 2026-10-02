# -*- coding: utf-8 -*-
"""01_Raw / 02_Std / 03_Sum 的公共目录表。

业务文件夹统一为两位序号：``02_成本统计收集``、``31_责任成本``。
查找时忽略序号，因此 ``3.1_责任成本`` 与 ``31_责任成本`` 是同一目录。
新建目录时用本表里的正式名。流水线内部的 ``A1_``、``B_``、``C_`` 不在此表，名称不动。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

_ORDER_PREFIX_RE = re.compile(r"^(?:\d+(?:\.\d+)?[_-])+")
_NON_KEY_RE = re.compile(r"[^\w\u4e00-\u9fff]+", re.UNICODE)


@dataclass(frozen=True)
class FolderSpec:
    code: str
    key: str
    title: str


STAGES: tuple[FolderSpec, ...] = (
    FolderSpec("01", "raw", "01_Raw"),
    FolderSpec("02", "std", "02_Std"),
    FolderSpec("03", "sum", "03_Sum"),
)

# 三个阶段目录下的业务文件夹。同一逻辑名在 Raw / Std / Sum 用同一个正式名。
DOMAINS: tuple[FolderSpec, ...] = (
    FolderSpec("02", "成本统计收集", "02_成本统计收集"),
    FolderSpec("03", "预算部分", "03_预算部分"),
    FolderSpec("31", "责任成本", "31_责任成本"),
    FolderSpec("04", "业主清单", "04_业主清单"),
    FolderSpec("04", "合同管理", "04_合同管理"),
    FolderSpec("05", "分包合同", "05_分包合同"),
    FolderSpec("09", "采购台账", "09_采购台账"),
    FolderSpec("91", "商混站", "91_商混站"),
    FolderSpec("91", "材料欠款", "91_材料欠款"),
    FolderSpec("92", "混凝土采购台账", "92_混凝土采购台账"),
    FolderSpec("10", "机具费", "10_机具费"),
    FolderSpec("11", "人工费", "11_人工费"),
    FolderSpec("12", "财务付款", "12_财务付款"),
    FolderSpec("13", "成本复盘", "13_成本复盘"),
    FolderSpec("13", "业主结算额", "13_业主结算额"),
)

# 业务目录内部仍用序号的子文件夹。
CHILDREN: tuple[FolderSpec, ...] = (
    FolderSpec("01", "人工费台账", "01_人工费台账"),
    FolderSpec("02", "机票", "02_机票"),
)

_BY_KEY: dict[str, FolderSpec] = {}
for _spec in (*STAGES, *DOMAINS, *CHILDREN):
    _BY_KEY.setdefault(_spec.key, _spec)


def strip_order_prefix(name: str) -> str:
    """去掉目录名开头的 2_、3.1_、05_、12- 这类序号。"""
    return _ORDER_PREFIX_RE.sub("", str(name or "").strip())


def dir_key(name: str) -> str:
    """比较用的名字：去序号、去符号、小写。"""
    text = strip_order_prefix(name)
    text = _NON_KEY_RE.sub("", text)
    return text.lower()


def _spec_for(name: str) -> FolderSpec | None:
    key = dir_key(name)
    if not key:
        return None
    for spec in (*DOMAINS, *CHILDREN, *STAGES):
        if dir_key(spec.key) == key or dir_key(spec.title) == key:
            return spec
    return None


def canonical_title(name: str) -> str | None:
    """逻辑名或旧文件夹名 → 正式名。对不上目录表时返回 None。"""
    spec = _spec_for(name)
    return spec.title if spec else None


def stage_title(stage: str) -> str:
    spec = _spec_for(stage)
    if spec is None or spec not in STAGES:
        raise KeyError(stage)
    return spec.title


def _find_child(parent: Path, logical: str) -> Path | None:
    if not parent.is_dir():
        return None
    target = dir_key(logical)
    if not target:
        return None
    for child in sorted(parent.iterdir(), key=lambda p: p.name):
        if child.is_dir() and not child.name.startswith(".") and dir_key(child.name) == target:
            return child
    return None


def domain_dir(data_root: Path | str, stage: str, key: str) -> Path:
    """阶段根下的业务目录。已有旧序号文件夹时直接返回它，否则用正式名。"""
    root = Path(data_root)
    stage_root = _find_child(root, stage) or (root / stage_title(stage))
    found = _find_child(stage_root, key)
    if found is not None:
        return found
    return stage_root / (canonical_title(key) or key)


def _rename_children(parent: Path) -> list[tuple[str, str]]:
    moved: list[tuple[str, str]] = []
    if not parent.is_dir():
        return moved
    for child in list(parent.iterdir()):
        if not child.is_dir() or child.name.startswith("."):
            continue
        title = canonical_title(child.name)
        if not title or title == child.name:
            continue
        dest = parent / title
        if dest.exists():
            raise FileExistsError(f"{child} → {dest}")
        child.rename(dest)
        moved.append((str(child), str(dest)))
    return moved


def rename_canonical(data_root: Path | str) -> list[tuple[str, str]]:
    """把三个阶段目录下的业务文件夹改成目录表里的正式名。"""
    root = Path(data_root)
    moved: list[tuple[str, str]] = []
    for stage in STAGES:
        stage_root = _find_child(root, stage.key)
        if stage_root is None:
            continue
        moved.extend(_rename_children(stage_root))
        labor = _find_child(stage_root, "人工费")
        if labor is not None:
            moved.extend(_rename_children(labor))
    return moved


__all__ = [
    "CHILDREN",
    "DOMAINS",
    "FolderSpec",
    "STAGES",
    "canonical_title",
    "dir_key",
    "domain_dir",
    "rename_canonical",
    "stage_title",
    "strip_order_prefix",
]
