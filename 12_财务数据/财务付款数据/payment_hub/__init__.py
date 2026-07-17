# -*- coding: utf-8 -*-
"""
财务付款一体化（自包含）

模块：
  std_build      原 3_ 标准化
  summary_build  原 4_ 汇总分析
  cashflow_build 原 5_ 现金流 HTML
  pipeline / cli 编排与命令行
"""

from .config import Paths
from .pipeline import run_pipeline, step_cashflow, step_std, step_summary

__version__ = "2.0.0"
__all__ = [
    "Paths",
    "run_pipeline",
    "step_std",
    "step_summary",
    "step_cashflow",
    "__version__",
]
