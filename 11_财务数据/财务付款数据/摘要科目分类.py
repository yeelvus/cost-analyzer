#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""摘要 → 一级 / 二级 / 三级科目。

和月结分开运行：
  python 摘要科目分类.py                         # 图形窗口
  python 摘要科目分类.py fill -i 明细.xlsx        # 另存「明细_科目.xlsx」
  python 摘要科目分类.py learn -i 已填科目.xlsx    # 只把已填科目写入记忆

不改 A1 / A2 / B_std / 汇总。月结填科目请用「财务付款一体化.py」。
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from subject_tool import main

if __name__ == "__main__":
    raise SystemExit(main())
