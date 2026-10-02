# -.- coding: utf-8 -*-
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
# 业务配置与学习库，放在代码仓库外，避免公开 GitHub 时带出项目、供应商和科目记忆。
# config/
#   common/project_name_match_db.json
#   project_info/unit_map.json
#   machine/机具映射表.json
#   reports/cost_analysis_report_config.json
#   payment/learn_classify/data
CONFIG_ROOT = DATA_ROOT / "config"
# 原始收集：DATA_ROOT/0_Raw（也可写成 Raw）→ common.paths.raw_data_path
# 标准化中间：DATA_ROOT/0_Std（也可写成 Std）→ common.paths.std_data_path
# HTML 看板：DATA_ROOT/04_html（也可写成 html / 0_html）→ common.paths.html_data_path
# 合并汇总 Excel：DATA_ROOT/0_Sum（也可写成 Sum）→ common.paths.sum_data_path

MATCH_DB_PATH = CONFIG_ROOT / "common" / "project_name_match_db.json"
UNIT_MAP_PATH = CONFIG_ROOT / "project_info" / "unit_map.json"
MACHINE_MAP_PATH = CONFIG_ROOT / "machine" / "机具映射表.json"
CLASSIFY_DATA_DIR = CONFIG_ROOT / "payment" / "learn_classify" / "data"
REPORT_CONFIG_PATH = CONFIG_ROOT / "reports" / "cost_analysis_report_config.json"


def code_path(*parts: str | Path) -> Path:
    return CODE_ROOT.joinpath(*parts)


def config_path(*parts: str | Path) -> Path:
    return CONFIG_ROOT.joinpath(*parts)


def data_path(*parts: str | Path) -> Path:
    return DATA_ROOT.joinpath(*parts)


def report_path(*parts: str | Path) -> Path:
    return REPORT_ROOT.joinpath(*parts)
