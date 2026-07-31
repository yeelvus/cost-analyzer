# -*- coding: utf-8 -*-
"""包路径：保证可 import learn_classify。"""

from __future__ import annotations

import sys
from pathlib import Path

# 财务付款数据/ （含 payment_hub 与 learn_classify）
CODE_ROOT = Path(__file__).resolve().parent.parent


def ensure_code_root() -> Path:
    root = str(CODE_ROOT)
    if root not in sys.path:
        sys.path.insert(0, root)
    return CODE_ROOT
