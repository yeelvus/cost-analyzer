"""
采购台账共享路径与轻量公共工具。

- 数据目录按关键词查找，忽略 A_/B_/C_ 等前缀
- 供 1 重命名 → 2 去重提取 → 3 标准化 共用
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import data_path, find_child_dir, find_script as find_script_file

_DATA_BASE = data_path("9_采购台账📦")
_SCRIPT_BASE = Path(__file__).parent

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
MONTH_CN_RE = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")
MONTH_ISO_RE = re.compile(r"((?:19|20)\d{2})-(0[1-9]|1[0-2])")


def find_data_dir(keyword: str) -> Path:
    return find_child_dir(_DATA_BASE, keyword)


def find_script(keyword: str) -> Path:
    return find_script_file(_SCRIPT_BASE, keyword)


# ── 数据目录 ──────────────────────────────────────────────
RAW_DIR = find_data_dir("采购部采购台账_raw")          # 月度原始收集
EXTRACT_DIR = find_data_dir("采购台账提取")            # 去重后每项目 1 文件
STD_DIR = find_data_dir("采购部采购台账_std")           # 标准化明细
SUM_DIR = find_data_dir("采购部采购台账_sum")           # 汇总输出
# 兼容旧名
SIMPLIFY_DIR = find_data_dir("采购部采购台账_simplify")

# ── 项目主数据 ────────────────────────────────────────────
PROJECT_INFO_FILE = data_path("0_项目基本信息情况") / "项目基本信息表.xlsx"


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
    # 25年3月
    m2 = re.search(r"(\d{2})年\s*(0?[1-9]|1[0-2])月", text)
    if m2:
        y = int(m2.group(1))
        y += 2000 if y < 70 else 1900
        return f"{y:04d}-{int(m2.group(2)):02d}"
    return None


def month_sort_key(month_iso: str) -> tuple[int, int]:
    y, m = month_iso.split("-")
    return int(y), int(m)
