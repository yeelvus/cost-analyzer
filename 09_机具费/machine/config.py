# -*- coding: utf-8 -*-
"""机具费统一配置。

  01_Raw/10_机具费/                 原始台账
  02_Std/10_机具费/
    B_机具费_raw_sum/              机具费_raw合并.xlsx（过程全量合并 + 项目匹配）
    C_机具费_raw_split/            原始拆分：{编码}_{项目}_机具费.xlsx
    E_机具费_std_split/            标准拆分：标准汇总按项目
  03_Sum/10_机具费/                 机具费_汇总.xlsx
分析报告 HTML 输出到 04_html/机具费分析/
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
MACHINE_CODE_DIR = PACKAGE_DIR.parent
CODE_ROOT = MACHINE_CODE_DIR.parent

for p in (str(MACHINE_CODE_DIR), str(CODE_ROOT)):
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from settings import DATA_ROOT, MACHINE_MAP_PATH, MATCH_DB_PATH, REPORT_ROOT  # type: ignore
    from common.paths import html_data_path, raw_data_path, std_data_path, sum_data_path  # type: ignore
except Exception:
    DATA_ROOT = Path(os.environ.get("COST_DATA_ROOT", str(Path.home() / "Desktop/邯郸中材/01成本统计")))
    MATCH_DB_PATH = DATA_ROOT / "config" / "common" / "project_name_match_db.json"
    MACHINE_MAP_PATH = DATA_ROOT / "config" / "machine" / "机具映射表.json"
    REPORT_ROOT = DATA_ROOT / "Z_report"

    def raw_data_path(*parts, legacy=None, create=True):  # type: ignore
        if legacy:
            return Path(legacy)
        return DATA_ROOT.joinpath("0_Raw", *parts)

    def html_data_path(*parts, create=True):  # type: ignore
        p = DATA_ROOT.joinpath("04_html", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

    def sum_data_path(*parts, create=True, legacy=None):  # type: ignore
        p = DATA_ROOT.joinpath("0_Sum", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

    def std_data_path(*parts, create=True, legacy=None):  # type: ignore
        if not create and legacy:
            lp = Path(legacy)
            if lp.exists():
                return lp
        p = DATA_ROOT.joinpath("0_Std", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p


def first_existing(*candidates: Path) -> Path:
    for p in candidates:
        if p and Path(p).exists():
            return Path(p)
    return Path(candidates[0]) if candidates else Path()


MACHINE_ROOT = first_existing(DATA_ROOT / "10_机具费🚜", DATA_ROOT / "10_机具费")

# ---------- 数据目录（A→F 顺序）----------
RAW_DIR = raw_data_path("机具费", legacy=MACHINE_ROOT / "A_机具费_raw")  # 原始台账 → 01_Raw
RAW_SUM_DIR = std_data_path(
    "机具费",
    "B_机具费_raw_sum",
    legacy=MACHINE_ROOT / "B_机具费_raw_sum",
)
RAW_SPLIT_DIR = std_data_path(
    "机具费",
    "C_机具费_raw_split",
    legacy=MACHINE_ROOT / "C_机具费_raw_split",
)
STD_SUM_DIR = sum_data_path("机具费", legacy=MACHINE_ROOT / "D_机具费_std_sum")
STD_SPLIT_DIR = std_data_path(
    "机具费",
    "E_机具费_std_split",
    legacy=MACHINE_ROOT / "E_机具费_std_split",
)
REPORT_DIR = html_data_path("机具费分析")  # HTML 分析报告 → 0_html
LEGACY_REPORT_DIR = MACHINE_ROOT / "F_机具费_report"    # 旧分析报告目录

# 兼容旧路径（回退查找 / 迁移源）
LEGACY_PATHS = {
    "RAW_SUM": [
        MACHINE_ROOT / "机具费_raw_sum",
        MACHINE_ROOT / "B_机具费_std",
    ],
    "RAW_SPLIT": [
        MACHINE_ROOT / "原始拆分",
        MACHINE_ROOT / "B_机具费_std",
    ],
    "STD_SUM": [
        MACHINE_ROOT / "标准汇总",
        MACHINE_ROOT / "B_机具费_sum",
    ],
    "STD_SPLIT": [
        MACHINE_ROOT / "标准拆分",
    ],
    "REPORT": [
        MACHINE_ROOT / "F_机具费_report",
        MACHINE_ROOT / "C_机具租赁分析报告",
    ],
}

LEGACY_STD_DIR = MACHINE_ROOT / "B_机具费_std"
LEGACY_SUM_DIR = MACHINE_ROOT / "B_机具费_sum"

# 别名
STD_DIR = RAW_SPLIT_DIR
SUM_DIR = STD_SUM_DIR


def ensure_data_dirs() -> None:
    """创建 A→E 流水线目录，以及 Z_report 分析报告目录。"""
    for d in (RAW_DIR, RAW_SUM_DIR, RAW_SPLIT_DIR, STD_SUM_DIR, STD_SPLIT_DIR, REPORT_DIR):
        d.mkdir(parents=True, exist_ok=True)


def migrate_legacy_dirs(*, move_files: bool = True) -> list[str]:
    """把旧目录中的文件迁移到新 A→F 目录（若新目录尚空且旧目录有内容）。"""
    ensure_data_dirs()
    logs: list[str] = []

    migrations = [
        (RAW_SUM_DIR, LEGACY_PATHS["RAW_SUM"], ("机具费_raw合并.xlsx", "clean_机具费_raw合并.xlsx")),
        (RAW_SPLIT_DIR, LEGACY_PATHS["RAW_SPLIT"], None),  # 全部 *_机具费.xlsx
        (STD_SUM_DIR, LEGACY_PATHS["STD_SUM"], None),
        (STD_SPLIT_DIR, LEGACY_PATHS["STD_SPLIT"], None),
        (REPORT_DIR, LEGACY_PATHS["REPORT"], None),
    ]

    for new_dir, old_dirs, only_names in migrations:
        new_dir.mkdir(parents=True, exist_ok=True)
        for old in old_dirs:
            if not old.exists() or old.resolve() == new_dir.resolve():
                continue
            # 原始拆分不要从 B_std 整包搬（含合并过程文件）
            if new_dir == RAW_SPLIT_DIR and old.name in ("B_机具费_std",):
                for f in old.glob("*_机具费.xlsx"):
                    if "合并" in f.name:
                        continue
                    dest = new_dir / f.name
                    if not dest.exists():
                        if move_files:
                            shutil.move(str(f), str(dest))
                        else:
                            shutil.copy2(f, dest)
                        logs.append(f"{f.name} → {new_dir.name}/")
                continue

            if only_names:
                for name in only_names:
                    src = old / name
                    dest = new_dir / name
                    if src.exists() and not dest.exists():
                        if move_files:
                            shutil.move(str(src), str(dest))
                        else:
                            shutil.copy2(src, dest)
                        logs.append(f"{name} → {new_dir.name}/")
            else:
                for f in old.iterdir():
                    if not f.is_file() or f.name.startswith("."):
                        continue
                    dest = new_dir / f.name
                    if not dest.exists():
                        if move_files:
                            shutil.move(str(f), str(dest))
                        else:
                            shutil.copy2(f, dest)
                        logs.append(f"{f.name} → {new_dir.name}/")
    return logs


MACHINE_MAPPING_JSON = first_existing(
    MACHINE_MAP_PATH,
    MACHINE_CODE_DIR / "机具映射表.json",
)
REGION_MAPPING_JSON = first_existing(
    MACHINE_MAP_PATH.parent / "区域映射表.json",
    MACHINE_CODE_DIR / "区域映射表.json",
    MACHINE_MAPPING_JSON,
)
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
LEGACY_CLEAN_MERGE_NAME = "clean_机具费_raw合并.xlsx"
PROJECT_STD_SUFFIX = "_机具费.xlsx"
FULL_SUMMARY_BASENAME = "机具费"
FULL_SUMMARY_FILENAME = f"{FULL_SUMMARY_BASENAME}_{OUTPUT_FILE_SUFFIX}.xlsx"

SUMMARY_SHEET_NAME = "汇总"
PROJECT_SUMMARY_SHEET_NAME = "项目汇总机具费"
PROJECT_MONTHLY_CHANGE_SHEET_NAME = "项目月度金额增量"
PROJECT_MONTHLY_WIDE_SHEET_NAME = "项目月份金额宽表"
PROCESS_REPORT_SHEET_NAME = "处理报告"
