"""
采购台账共享路径模块。
- 文件夹按关键词动态查找，忽略 A_ / B_ / C_ 等前缀字母，前缀变更后无需修改代码。
- 脚本按关键词动态查找，忽略 1_ / 2_ / 3_ 等前缀数字，步骤编号变更后无需修改代码。
"""
from __future__ import annotations
import re
from pathlib import Path

# 数据根目录（包含所有采购台账子文件夹的目录）
_DATA_BASE = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦")

# 当前脚本所在目录（各采购数据脚本目录）
_SCRIPT_BASE = Path(__file__).parent


def find_data_dir(keyword: str) -> Path:
    """在 _DATA_BASE 下查找名称包含 keyword 的目录，忽略前缀字母/数字。
    若未找到则返回 _DATA_BASE / keyword（方便调用方发现路径不存在）。
    """
    if _DATA_BASE.is_dir():
        for d in sorted(_DATA_BASE.iterdir()):
            if d.is_dir() and keyword in d.name:
                return d
    return _DATA_BASE / keyword


def find_script(keyword: str) -> Path:
    """在 _SCRIPT_BASE 下查找文件名包含 keyword 的 .py 脚本，忽略数字前缀。
    若未找到则返回 _SCRIPT_BASE / keyword。
    """
    if _SCRIPT_BASE.is_dir():
        for f in sorted(_SCRIPT_BASE.iterdir()):
            if f.is_file() and f.suffix == ".py" and keyword in f.name:
                return f
    return _SCRIPT_BASE / keyword


# ── 预解析常用路径常量（供各脚本直接 import 使用）──────────────────────────
RAW_DIR      = find_data_dir("采购部采购台账_raw")
STD_DIR      = find_data_dir("采购部采购台账_std")
SIMPLIFY_DIR = find_data_dir("采购部采购台账_simplify")
SUM_DIR      = find_data_dir("采购部采购台账_sum")
