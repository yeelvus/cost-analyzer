#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口（脚本已迁入 S_数据分析模块）。

真实位置：S_数据分析模块/analysis/成本_数据分析.py
"""

from __future__ import annotations

import runpy
import sys
from pathlib import Path

_here = Path(__file__).resolve().parent
_code = _here if (_here / "settings.py").exists() else _here.parent
_target = _code / "S_数据分析模块/analysis/成本_数据分析.py"
if not _target.is_file():
    raise SystemExit(f"未找到已迁移脚本: {_target}")

sys.argv[0] = str(_target)
runpy.run_path(str(_target), run_name="__main__")
