#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务付款月结入口。

  python 财务付款一体化.py           # 月结窗口
  python 财务付款一体化.py --menu    # 终端菜单
  python 财务付款一体化.py run       # 月结全流程

摘要对应一二三级科目在旁边的「摘要科目分类.py」，两边分开运行。
旧的 pool / apply / learn-ledger 仍可在命令行用，apply 不再改月结文件。
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from payment_hub.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
