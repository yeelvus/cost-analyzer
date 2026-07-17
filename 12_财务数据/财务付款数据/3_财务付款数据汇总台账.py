#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
兼容入口：逻辑已迁入 payment_hub.std_build。
请优先使用：python 财务付款一体化.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from payment_hub.config import Paths
from payment_hub.std_build import run_std


def main():
    paths = Paths()
    outs = run_std(paths)
    print(f"\n处理完成，输出 {len(outs)} 个文件")
    for p in outs:
        print(f"- {p}")


if __name__ == "__main__":
    main()
