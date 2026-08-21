# -*- coding: utf-8 -*-
"""
预算模块统一配置。

可通过环境变量覆盖数据根目录：
  COST_DATA_ROOT=/path/to/01成本统计
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# 本包与代码仓库根
PACKAGE_DIR = Path(__file__).resolve().parent
BUDGET_CODE_DIR = PACKAGE_DIR.parent
CODE_ROOT = BUDGET_CODE_DIR.parent  # .../code

if str(CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(CODE_ROOT))

try:
    from settings import DATA_ROOT, MATCH_DB_PATH, REPORT_ROOT  # type: ignore
except Exception:
    DATA_ROOT = Path(
        os.environ.get(
            "COST_DATA_ROOT",
            str(Path.home() / "Desktop/邯郸中材/01成本统计"),
        )
    )
    MATCH_DB_PATH = CODE_ROOT / "0_项目基本信息" / "project_name_match_db.json"
    REPORT_ROOT = DATA_ROOT / "Z_report"

# ---------- 广联达人材机路径 ----------
BUDGET_ROOT = DATA_ROOT / "3_预算部分"
GLODON_RAW_DIR = BUDGET_ROOT / "1_广联达导出人材料机_raw"
GLODON_STD_DIR = BUDGET_ROOT / "2_人材机_std"
PROJECT_CODE_MAPPING_XLSX = DATA_ROOT / "0_项目基本信息情况" / "项目编码名称汇总.xlsx"

# ---------- 责任成本路径 ----------
RESP_ROOT = DATA_ROOT / "3.1_责任成本"
RESP_RAW_DIR = RESP_ROOT / "1_责任成本_raw"
RESP_STD_DIR = RESP_ROOT / "2_责任成本_std"

# ---------- 广联达处理参数 ----------
AUTO_MODE = True
SELECTED_COLUMNS: list = []  # 空=全部列
KEYWORD_IN_SHEETNAME = "标段人材机汇总"
CODE_COL = "编码"
CATEGORY_COL = "类别"
NAME_COL = "名称"
UNIT_COL = "单位"
TARGET_COL_NAME = "不含税市场价合计"
FALLBACK_TARGET_COL_NAME = "含税市场价合计"
NEW_COL = "不含税市场价合计（万）"
MATERIAL_CATS = {"材", "商浆", "商砼", "主"}
PEOPLE_CAT = "人"
PIPE_CAT = "管"
MECH_CATS = {"设", "机"}
QUANTITY_COL = "数量"
AMOUNT_COL = "金额"
SUMMARY_SHEET_NAME = "汇总_不含税市场价"
TOTAL_SUMMARY_PREFIX = "【总汇总】所有标段_"
MATERIAL_SHEET_NAME = "材料汇总"
UNIT_ANOMALY_SHEET = "单位异常明细"
DASHBOARD_HTML_NAME = "总汇总项目看板.html"
DASHBOARD_HTML_PATH = REPORT_ROOT / DASHBOARD_HTML_NAME

# include_kw, exclude_kw, unit_codes（标准码；空=不限制单位）
# 砂浆/碎石允许 m³ 与 t/kg（重量按容重折 m³，见 budget.units）
MATERIAL_SUB_FILTERS = [
    ("钢筋", ["钢筋"], [], ["KG", "TON", "EMPTY"]),
    ("混凝土", ["混凝土"], [], ["M3", "EMPTY"]),
    ("砂浆", ["砂浆"], ["踢脚线"], ["M3", "KG", "TON", "EMPTY"]),
    ("碎石", ["碎石"], [], ["M3", "KG", "TON", "EMPTY"]),
    ("砖", ["砖", "砌块"], ["釉面砖", "花岗岩地砖"], ["PCS", "EMPTY"]),
]

# 砂浆/碎石：重量→体积容重 (kg/m³)，可按项目调整
MORTAR_BULK_DENSITY_KG_PER_M3 = 1500.0
AGGREGATE_BULK_DENSITY_KG_PER_M3 = 1600.0

BASE_COLUMNS = [
    "标段人材机汇总",
    "人_明细",
    "管_明细",
    "设机_明细",
    "材类明细",
    "其他类别_明细",
    "钢筋",
    "混凝土",
    "砂浆",
    "碎石",
    "砖",
    "其他材料",
]

GROUP1 = ["人_明细", "管_明细", "设机_明细", "材类明细", "其他类别_明细"]
GROUP2 = ["钢筋", "混凝土", "砂浆", "碎石", "砖", "其他材料"]

FINAL_COLUMNS = ["序号", "项目编码", "项目名称", "币种"] + BASE_COLUMNS

MAIN_KEYWORDS = ["标段人材机汇总", "人材机汇总", "汇总表"]
PRICE_COL_CANDIDATES = [
    "不含税市场价",
    "含税市场价",
    "市场价",
    "单价",
    "材料单价",
    "不含税市场价(元)",
    "含税市场价(元)",
    TARGET_COL_NAME,
    FALLBACK_TARGET_COL_NAME,
]

MATERIAL_NAMES = ["钢筋", "混凝土", "砂浆", "碎石", "砖", "其他材料"]
