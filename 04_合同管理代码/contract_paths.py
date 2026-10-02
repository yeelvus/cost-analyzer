# -*- coding: utf-8 -*-
"""合同管理根目录。

现行目录在 OneDrive。OneDrive 未挂上时，才回退到桌面成本统计里的同名目录。
"""

from __future__ import annotations

from pathlib import Path

CONTRACT_ROOT = Path(
    "/Users/wyattsun/Library/CloudStorage/OneDrive-个人/10-邯郸中材-工作/合同管理📑"
)

_LEGACY_ROOTS = (
    Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/合同管理📑"),
    Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑"),
)


def resolve_contract_root() -> Path:
    """返回当前可用的合同管理根目录。"""
    if CONTRACT_ROOT.is_dir():
        return CONTRACT_ROOT
    for legacy in _LEGACY_ROOTS:
        if legacy.is_dir():
            return legacy
    return CONTRACT_ROOT
