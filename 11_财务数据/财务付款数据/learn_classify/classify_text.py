# -*- coding: utf-8 -*-
"""分类唯一入口。

先取出业务摘要，再交给 ClassifyEngine。
引擎看到的摘要是业务描述，不带项目名前缀。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Tuple

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from summary_recognize import SummaryRecognition, resolve_business

from .engine import ClassifyEngine, ClassifyResult


def classify_text(
    engine: ClassifyEngine,
    subject: Any = None,
    summary: Any = None,
    vendor: Any = None,
    project_name: Any = None,
    business_summary: Any = None,
) -> Tuple[SummaryRecognition, ClassifyResult]:
    """识别（或沿用）业务摘要后分类。返回 (识别结果, 分类结果)。"""
    rec = resolve_business(summary, project_name, business_summary)
    result = engine.classify(subject, rec.business_summary, vendor)
    return rec, result
