# -*- coding: utf-8 -*-
"""
财务付款分类学习闭环。

稳定 ID 的 taxonomy + 记忆库 + 纠错学习 + 迁移重分类。
"""

from pathlib import Path

from .adapter import SubjectMapper
from .engine import ClassifyEngine, ClassifyResult
from .taxonomy import Taxonomy

__version__ = "1.0.0"
DEFAULT_DATA_DIR = Path(__file__).resolve().parent / "data"

__all__ = [
    "ClassifyEngine",
    "ClassifyResult",
    "SubjectMapper",
    "Taxonomy",
    "DEFAULT_DATA_DIR",
    "__version__",
]
