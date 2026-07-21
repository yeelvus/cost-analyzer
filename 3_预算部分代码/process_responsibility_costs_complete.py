# -*- coding: utf-8 -*-
"""兼容入口：责任成本处理。

推荐：
  python 00_预算处理入口.py responsibility
  python -m budget responsibility
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(ROOT.parent) not in sys.path:
    sys.path.insert(0, str(ROOT.parent))

from budget.responsibility import main

if __name__ == "__main__":
    main()
