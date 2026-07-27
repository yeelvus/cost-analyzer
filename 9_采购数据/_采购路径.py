# -*- coding: utf-8 -*-
"""
采购台账共享路径与轻量公共工具。

目录结构（代码）：
  9_采购数据/
    0_run_采购全流程.py   # 入口
    _采购路径.py          # 本文件
    流水线/               # 1→5 ETL
    规格标准化/           # 物品规格规则
    单位转换/             # 单位/数量换算
    内容标准化/           # 合并时内容标准化
    分析看板/             # 价格/成本分析
    工具/                 # 旁路工具

数据目录（磁盘 9_采购台账📦）：
  A_raw → A_提取 → B_std → D_sum
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

# 本包根 = 9_采购数据/
PACKAGE_ROOT = Path(__file__).resolve().parent
# monorepo code 根
_CODE_ROOT = PACKAGE_ROOT.parent
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))
if str(PACKAGE_ROOT) not in sys.path:
    sys.path.insert(0, str(PACKAGE_ROOT))

from common.paths import data_path, find_child_dir, find_script as find_script_file

_DATA_BASE = data_path("9_采购台账📦")
_SCRIPT_BASE = PACKAGE_ROOT

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
MONTH_CN_RE = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")
MONTH_ISO_RE = re.compile(r"((?:19|20)\d{2})-(0[1-9]|1[0-2])")


def find_data_dir(keyword: str) -> Path:
    return find_child_dir(_DATA_BASE, keyword)


def find_script(keyword: str) -> Path:
    return find_script_file(_SCRIPT_BASE, keyword)


# ── 代码模块目录 ──────────────────────────────────────────
PIPELINE_DIR = PACKAGE_ROOT / "流水线"
SPEC_DIR = PACKAGE_ROOT / "规格标准化"
UNIT_DIR = PACKAGE_ROOT / "单位转换"
CONTENT_DIR = PACKAGE_ROOT / "内容标准化"
ANALYSIS_CODE_DIR = PACKAGE_ROOT / "分析看板"
TOOLS_DIR = PACKAGE_ROOT / "工具"

CONTENT_STD_UTILS = CONTENT_DIR / "采购内容标准化_utils.py"
UNIT_PIPELINE = UNIT_DIR / "0_统一转换流水线.py"
SPEC_PIPELINE = SPEC_DIR / "0_统一标准化流水线.py"
ANALYSIS_BATCH = ANALYSIS_CODE_DIR / "0_汇总分析批量运行.py"

# ── 数据目录（9_采购台账📦 下，忽略 A_/B_/D_ 前缀）────────
RAW_DIR = find_data_dir("采购部采购台账_raw")
EXTRACT_DIR = find_data_dir("采购台账提取")
STD_DIR = find_data_dir("采购部采购台账_std")
SUM_DIR = find_data_dir("采购部采购台账_sum")
STD_SPLIT_DIR = find_data_dir("采购部采购台账_std_split")
SIMPLIFY_DIR = find_data_dir("采购部采购台账_simplify")

# ── 项目主数据 / 分析输出 ─────────────────────────────────
PROJECT_INFO_FILE = data_path("0_项目基本信息情况") / "项目基本信息表.xlsx"
try:
    from common.paths import report_path as _report_path

    ANALYSIS_DIR = _report_path("C_采购数据分析")
except Exception:
    ANALYSIS_DIR = data_path("Z_report") / "C_采购数据分析"

MERGED_PLAIN_NAME = "采购部采购台账_merged.xlsx"
MERGED_STD_NAME = "采购部采购台账_std_merged.xlsx"
MERGED_COMPARE_NAME = "采购部采购台账_std_merged_对比.xlsx"


def is_excel(path: Path) -> bool:
    return (
        path.is_file()
        and not path.name.startswith("~$")
        and not path.name.startswith(".")
        and path.suffix.lower() in EXCEL_EXTS
    )


def list_excels(folder: Path, *, recursive: bool = False) -> list[Path]:
    if not folder.is_dir():
        return []
    it = folder.rglob("*") if recursive else folder.iterdir()
    return sorted(p for p in it if is_excel(p))


def latest_sum_file(pattern: str = "*std_merged*.xlsx") -> Path | None:
    """SUM_DIR 下按修改时间取最新匹配文件。"""
    if not SUM_DIR.is_dir():
        return None
    cands = [p for p in SUM_DIR.glob(pattern) if is_excel(p)]
    if not cands:
        cands = [p for p in SUM_DIR.glob("*.xlsx") if is_excel(p)]
    return max(cands, key=lambda p: p.stat().st_mtime) if cands else None


def extract_project_name(filename: str) -> str:
    """从文件名取项目前缀：xxx_... 或 xxx项目..."""
    stem = Path(filename).stem.strip()
    if "_" in stem:
        return stem.split("_", 1)[0].strip()
    m = re.match(r"^(.+?项目)", stem)
    return m.group(1).strip() if m else stem


def extract_month_cn(text: str) -> str | None:
    """从路径/文件名提取「YYYY年MM月」。"""
    m = MONTH_CN_RE.search(text)
    if not m:
        return None
    return f"{int(m.group(1)):04d}年{int(m.group(2)):02d}月"


def extract_month_iso(text: str) -> str | None:
    """提取 YYYY-MM（优先中文年月，其次 ISO）。"""
    cn = extract_month_cn(text)
    if cn:
        y, rest = cn.split("年", 1)
        mm = rest.replace("月", "")
        return f"{int(y):04d}-{int(mm):02d}"
    m = MONTH_ISO_RE.search(text)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    m2 = re.search(r"(\d{2})年\s*(0?[1-9]|1[0-2])月", text)
    if m2:
        y = int(m2.group(1))
        y += 2000 if y < 70 else 1900
        return f"{y:04d}-{int(m2.group(2)):02d}"
    return None


def month_sort_key(month_iso: str) -> tuple[int, int]:
    y, m = month_iso.split("-")
    return int(y), int(m)


def ensure_package_path() -> Path:
    """保证 PACKAGE_ROOT / CODE_ROOT 在 sys.path，返回包根。"""
    if str(PACKAGE_ROOT) not in sys.path:
        sys.path.insert(0, str(PACKAGE_ROOT))
    if str(_CODE_ROOT) not in sys.path:
        sys.path.insert(0, str(_CODE_ROOT))
    return PACKAGE_ROOT
