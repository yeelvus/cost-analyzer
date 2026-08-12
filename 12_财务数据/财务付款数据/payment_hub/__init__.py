# -*- coding: utf-8 -*-
"""
财务付款一体化（payment_hub）

架构：
  cli / menu     入口与交互菜单
  pipeline       步骤编排（std / summary / cashflow / 分类）
  learn_pool     D_ 学习池 export / apply
  std_build      raw → B_std
  summary_build  B_std → 汇总 Excel
  cashflow_build B_std → 现金流 HTML
  config / utils / deps / pkg

分类引擎在 learn_classify/（memory / rules / taxonomy）。
"""

from .config import Paths
from .pipeline import (
    run_pipeline,
    step_cashflow,
    step_classify,
    step_raw_std,
    step_std,
    step_summary,
)

__version__ = "2.4.0"
__all__ = [
    "Paths",
    "run_pipeline",
    "step_raw_std",
    "step_classify",
    "step_std",
    "step_summary",
    "step_cashflow",
    "__version__",
]
