#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务付款一体化入口

推荐：
  python 财务付款一体化.py           # 交互式菜单（直接选 1/2/3…）

也可直接子命令：
  python 财务付款一体化.py pool      # 导出学习池
  python 财务付款一体化.py apply     # 应用确认
  python 财务付款一体化.py run       # 月结出账

旧命令名（export-learn-pool / all / apply-class 等）仍兼容。
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
