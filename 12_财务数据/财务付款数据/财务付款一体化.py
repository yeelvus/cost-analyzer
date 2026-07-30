#!.usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务付款一体化入口（推荐使用这个文件）

融合功能：
  · 3_ 原始付款 → 标准明细（项目匹配 + 一二级分类）
  · 4_ 标准明细 → 汇总分析 Excel
  · 5_ 标准明细 → 项目现金流 HTML
  · learn_classify 可持续学习 / 分类迁移

快速开始：
  python 财务付款一体化.py              # 默认 = apply-class（学分类总表叶子变更）
  python 财务付款一体化.py apply-class  # 同上
  python 财务付款一体化.py export-class # 导出付款分类总表
  python 财务付款一体化.py all          # 月结全流程 std → summary → cashflow
  python 财务付款一体化.py month        # 同 all
  python 财务付款一体化.py compact-memory --dry-run
  python 财务付款一体化.py help-guide
  python 财务付款一体化.py show-paths

注意：无参数不是全流程；出汇总/HTML 请用 all 或 month。
"""

from __future__ import annotations

import sys
from pathlib import Path

# 保证本目录与子包可导入
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from payment_hub.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
