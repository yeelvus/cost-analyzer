# -*- coding: utf-8 -*-
"""
采购台账共享路径与轻量公共工具。

目录结构（代码）：
  9_采购数据/
    0_run_采购全流程.py   # 入口（默认 04a→04→06）
    _采购路径.py          # 本文件
    流水线/               # 01–06 ETL
    规格标准化/ / 单位转换/ / 内容标准化/
    分析看板/ / 工具/

数据目录（磁盘 9_采购台账📦，前缀按流水线顺序）：
  A_raw → B_extract → C_std → D_raw_sum → E_sum → F_std_split
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

from common.paths import (
    data_path,
    find_child_dir,
    find_script as find_script_file,
    raw_data_path,
    std_data_path,
    sum_data_path,
)

_DATA_BASE = data_path("9_采购台账📦")
_SCRIPT_BASE = PACKAGE_ROOT

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
MONTH_CN_RE = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")
MONTH_ISO_RE = re.compile(r"((?:19|20)\d{2})-(0[1-9]|1[0-2])")

# 流水线目录正式名（前缀 A–F 与步骤顺序一致）
DIR_RAW = "A_采购部采购台账_raw"
DIR_EXTRACT = "B_采购部采购台账_extract"
DIR_STD = "C_采购部采购台账_std"
DIR_RAW_SUM = "D_采购部采购台账_raw_sum"
DIR_SUM = "E_采购部采购台账_sum"
DIR_STD_SPLIT = "F_采购部采购台账_std_split"


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


def _resolve_or_create_data_dir(
    keyword: str,
    create_name: str,
    *,
    legacy_names: tuple[str, ...] = (),
) -> Path:
    """
    优先匹配 keyword 目录；若正式前缀目录存在但名不一致则尽量迁到 create_name；
    再尝试 legacy 旧名迁移；否则创建 create_name。
    """
    # 1) 关键字命中且已是目标名 / 含关键字的现有目录
    found = find_data_dir(keyword)
    if found.is_dir() and keyword in found.name:
        target = _DATA_BASE / create_name
        # 旧前缀（如 D_…_sum）→ 新前缀（E_…_sum）
        if found.name != create_name and not target.exists():
            try:
                found.rename(target)
                print(f"已迁移目录: {found.name} → {create_name}")
                return target
            except OSError as exc:
                print(f"⚠️ 无法重命名 {found.name} → {create_name}: {exc}")
                return found
        return found

    # 2) 精确旧全名迁移
    for legacy in legacy_names:
        legacy_path = _DATA_BASE / legacy
        if not legacy_path.is_dir():
            continue
        target = _DATA_BASE / create_name
        if legacy_path.resolve() == target.resolve():
            return target
        if not target.exists():
            try:
                legacy_path.rename(target)
                print(f"已迁移目录: {legacy} → {create_name}")
                return target
            except OSError as exc:
                print(f"⚠️ 无法重命名 {legacy} → {create_name}: {exc}")
                return legacy_path
        return legacy_path

    # 3) 新建
    target = _DATA_BASE / create_name
    target.mkdir(parents=True, exist_ok=True)
    return target


# ── 数据目录（按流水线 A→F）────────────────────────────────
# A 原始 → B 去重提取 → C 格式标准 → D 纯合并 → E 标准汇总 → F 按项目拆分
# 原始台账在 0_Raw/9_采购台账（可带序号）；未迁移时回退旧目录
RAW_DIR = raw_data_path(
    "采购台账",
    legacy=_DATA_BASE / DIR_RAW,
)
EXTRACT_DIR = std_data_path(
    "采购台账",
    DIR_EXTRACT,
    legacy=_DATA_BASE / DIR_EXTRACT,
)
STD_DIR = std_data_path(
    "采购台账",
    DIR_STD,
    legacy=_DATA_BASE / DIR_STD,
)
# 注意：keyword 用 raw_sum / sum / std_split，避免「_std」误匹配 std_split
RAW_SUM_DIR = sum_data_path(
    "采购台账",
    legacy=_DATA_BASE / DIR_RAW_SUM,
)
SUM_DIR = sum_data_path(
    "采购台账",
    legacy=_DATA_BASE / DIR_SUM,
)
STD_SPLIT_DIR = std_data_path(
    "采购台账",
    DIR_STD_SPLIT,
    legacy=_DATA_BASE / DIR_STD_SPLIT,
)
SIMPLIFY_DIR = find_data_dir("采购部采购台账_simplify")

# ── 项目主数据 / 分析输出 ─────────────────────────────────
PROJECT_INFO_FILE = data_path("0_项目基本信息情况") / "项目基本信息表.xlsx"
# 月度汇率：年月 / 人民币对泰铢 / 美元对泰铢 / 人民币对美元
FX_RATE_FILE = data_path("0_项目基本信息情况") / "汇率三表.xlsx"
try:
    from common.paths import html_data_path as _html_data_path

    ANALYSIS_DIR = _html_data_path("采购数据分析")
except Exception:
    ANALYSIS_DIR = data_path("0_html") / "采购数据分析"

# 纯合并（C_std 直接拼表，无内容标准化）→ D_raw_sum
RAW_SUM_FILE_NAME = "采购部采购台账_raw_sum.xlsx"
# 兼容旧名引用
MERGED_PLAIN_NAME = RAW_SUM_FILE_NAME
# 标准化+对比（唯一 E_sum 主文件，数据源 D_raw_sum）
STD_SUM_FILE_NAME = "采购部采购台账_std_sum.xlsx"
MERGED_STD_NAME = STD_SUM_FILE_NAME
MERGED_COMPARE_NAME = STD_SUM_FILE_NAME  # 已与主文件合一，不再单独出「对比」文件
CATEGORY_SUMMARY_NAME = "项目采购分类汇总.xlsx"


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


def latest_sum_file(pattern: str = "*std_sum*.xlsx") -> Path | None:
    """SUM_DIR 下按修改时间取最新匹配文件。"""
    if not SUM_DIR.is_dir():
        return None
    cands = [p for p in SUM_DIR.glob(pattern) if is_excel(p)]
    if not cands:
        # 兼容旧名
        cands = [p for p in SUM_DIR.glob("*std_merged*.xlsx") if is_excel(p)]
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
