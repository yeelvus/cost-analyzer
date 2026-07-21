# -*- coding: utf-8 -*-
"""机具费统一配置。"""
from __future__ import annotations

import os
import sys
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
MACHINE_CODE_DIR = PACKAGE_DIR.parent
CODE_ROOT = MACHINE_CODE_DIR.parent

for p in (str(MACHINE_CODE_DIR), str(CODE_ROOT)):
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from settings import DATA_ROOT, MATCH_DB_PATH  # type: ignore
except Exception:
    DATA_ROOT = Path(os.environ.get("COST_DATA_ROOT", str(Path.home() / "Desktop/邯郸中材/01成本统计")))
    MATCH_DB_PATH = CODE_ROOT / "0_项目基本信息" / "project_name_match_db.json"


def first_existing(*candidates: Path) -> Path:
    for p in candidates:
        if p and Path(p).exists():
            return Path(p)
    return Path(candidates[0]) if candidates else Path()


MACHINE_ROOT = first_existing(DATA_ROOT / "10_机具费🚜", DATA_ROOT / "10_机具费")
RAW_DIR = MACHINE_ROOT / "A_机具费_raw"
STD_DIR = MACHINE_ROOT / "B_机具费_std"
SUM_DIR = MACHINE_ROOT / "B_机具费_sum"
REPORT_DIR = MACHINE_ROOT / "C_机具租赁分析报告"

MACHINE_MAPPING_JSON = first_existing(MACHINE_CODE_DIR / "机具映射表.json")
REGION_MAPPING_JSON = first_existing(MACHINE_CODE_DIR / "区域映射表.json", MACHINE_CODE_DIR / "机具映射表.json")
REGION_MAPPING_TXT = first_existing(MACHINE_CODE_DIR / "区域映射表.txt")
MACHINE_MAPPING_TXT = first_existing(MACHINE_CODE_DIR / "机具映射表.txt")
PROJECT_INFO_XLSX = first_existing(
    DATA_ROOT / "0_项目基本信息情况" / "项目基本信息表.xlsx",
    DATA_ROOT / "0_项目基本信息" / "项目基本信息表.xlsx",
)
PROJECT_MATCH_DB = Path(MATCH_DB_PATH)

RUN_MODE = os.environ.get("MACHINE_RUN_MODE", "all")  # all | monthly_pair
EXCEL_FILE_OVERRIDE = os.environ.get("MACHINE_EXCEL_OVERRIDE", "")
OUTPUT_FILENAME_PREFIX = "机具费汇总"
OUTPUT_FILE_SUFFIX = "汇总"
RAW_MERGE_NAME = "机具费_raw合并.xlsx"
LEGACY_CLEAN_MERGE_NAME = "clean_机具费_raw合并.xlsx"  # 废弃，不再生成
PROJECT_STD_SUFFIX = "_机具费.xlsx"  # 按项目拆分后的文件名后缀
FULL_SUMMARY_BASENAME = f"{OUTPUT_FILENAME_PREFIX}_全量标准化"

SUMMARY_SHEET_NAME = "汇总"
PROJECT_SUMMARY_SHEET_NAME = "项目汇总机具费"
PROJECT_MONTHLY_CHANGE_SHEET_NAME = "项目月度金额增量"
PROJECT_MONTHLY_WIDE_SHEET_NAME = "项目月份金额宽表"
