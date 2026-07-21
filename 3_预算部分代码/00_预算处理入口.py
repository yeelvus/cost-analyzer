#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预算处理统一入口（推荐）。

等价于：
  cd 3_预算部分代码 && python -m budget [command]

示例：
  python 00_预算处理入口.py              # 默认 all：广联达 + 看板
  python 00_预算处理入口.py all
  python 00_预算处理入口.py glodon
  python 00_预算处理入口.py dashboard
  python 00_预算处理入口.py responsibility   # 责任成本（原 process_responsibility_costs_complete.py）
  python 00_预算处理入口.py units-test
  python 00_预算处理入口.py --help
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(ROOT.parent) not in sys.path:
    sys.path.insert(0, str(ROOT.parent))

from budget.__main__ import main

if __name__ == "__main__":
    main()
