# -*- coding: utf-8 -*-
"""全库共享路径配置（Desktop/code 根）。

数据与代码分离：
  CODE_ROOT   → 本仓库脚本
  DATA_ROOT   → 邯郸中材/01成本统计（业务数据）
  REPORT_ROOT → Z_report（分析报告输出）
"""
from __future__ import annotations

from pathlib import Path

CODE_ROOT = Path(__file__).resolve().parent
DATA_ROOT = Path.home() / "Desktop" / "邯郸中材" / "01成本统计"
REPORT_ROOT = DATA_ROOT / "Z_report"
# 原始收集：DATA_ROOT/0_Raw（也可写成 Raw）→ common.paths.raw_data_path
# HTML 看板：DATA_ROOT/0_html（也可写成 html）→ common.paths.html_data_path
# 合并汇总 Excel：DATA_ROOT/0_Sum（也可写成 Sum）→ common.paths.sum_data_path

# 项目名匹配库（0_项目基本信息）
MATCH_DB_PATH = CODE_ROOT / "0_项目基本信息" / "project_name_match_db.json"


def code_path(*parts: str | Path) -> Path:
    return CODE_ROOT.joinpath(*parts)


def data_path(*parts: str | Path) -> Path:
    return DATA_ROOT.joinpath(*parts)


def report_path(*parts: str | Path) -> Path:
    return REPORT_ROOT.joinpath(*parts)
