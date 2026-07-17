#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
兼容入口：逻辑已迁入 payment_hub.cashflow_build。
请优先使用：python 财务付款一体化.py cashflow
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from payment_hub.cashflow_build import run_cashflow
from payment_hub.config import Paths


def main():
    run_cashflow(Paths())


if __name__ == "__main__":
    main()
