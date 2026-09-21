#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机具费统一入口（重构版）。

  python 00_机具费处理入口.py
  python 00_机具费处理入口.py merge-raw
  python 00_机具费处理入口.py summarize
  python 00_机具费处理入口.py report
  python 00_机具费处理入口.py paths
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(ROOT.parent) not in sys.path:
    sys.path.insert(0, str(ROOT.parent))

from machine.__main__ import main

if __name__ == "__main__":
    main()
