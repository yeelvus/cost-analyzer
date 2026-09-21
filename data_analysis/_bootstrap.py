# -*- coding: utf-8 -*-
"""确保 Desktop/code 根目录与脚本所在目录在 sys.path 中。"""

from __future__ import annotations

import sys
from pathlib import Path


def ensure_code_root(start: Path | None = None) -> Path:
    """向上查找含 settings.py 的 code 根，并插入 sys.path。"""
    p = (start or Path(__file__)).resolve()
    if p.is_file():
        p = p.parent
    for candidate in [p, *p.parents]:
        if (candidate / "settings.py").exists():
            s = str(candidate)
            if s not in sys.path:
                sys.path.insert(0, s)
            return candidate
    # 回退：S_数据分析模块 的上一级
    fallback = Path(__file__).resolve().parent.parent
    s = str(fallback)
    if s not in sys.path:
        sys.path.insert(0, s)
    return fallback


def ensure_script_dir(script_file: str | Path) -> Path:
    """把脚本所在目录加入 path（便于同目录互相 import）。"""
    d = Path(script_file).resolve().parent
    s = str(d)
    if s not in sys.path:
        sys.path.insert(0, s)
    return d


# 导入本模块时自动挂载 code 根
CODE_ROOT = ensure_code_root()
