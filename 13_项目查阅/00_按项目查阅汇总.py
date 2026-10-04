# -*- coding: utf-8 -*-
"""生成按项目查阅的总汇总和各项目明细工作簿。

用法：
    python 13_项目查阅/00_按项目查阅汇总.py

结果在 01成本统计/19_项目查阅。不改现有流水线、02_Std、03_Sum 和 18_output。
"""

from __future__ import annotations

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_CODE_ROOT = _HERE.parent
for _path in (str(_HERE), str(_CODE_ROOT)):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from project_view import main


if __name__ == "__main__":
    main()
