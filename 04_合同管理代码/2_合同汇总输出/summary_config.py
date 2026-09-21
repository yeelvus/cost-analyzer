# -*- coding: utf-8 -*-
"""合同汇总公共配置。

路径按「_ 之后」的业务名匹配，前面的序号/字母可以改：
  5_合同管理📑 / 6_合同管理 / 合同管理📑  → 合同管理
  A_总包合同 / 总包合同                 → 总包合同
  B_分包合同                           → 分包合同
  0_Std / 5.1_分包合同                 → Std / 分包合同
项目文件夹 THCC-2024-001_胜途制造 不在此规则内（编码不是排序前缀）。
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[2]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import dir_key, find_named_dir, std_data_path  # noqa: E402
from settings import DATA_ROOT, MATCH_DB_PATH  # noqa: E402

# 排序前缀：5_ / 5.1_ / 00_ / A_ / B_
_SORT_PREFIX_RE = re.compile(r"^(?:\d+(?:\.\d+)?|[A-Za-z])[_-]")

# 合同根目录逻辑名（实际磁盘名可以带任意前缀，如 5_合同管理📑）
CONTRACT_ROOT_LOGICAL = "合同管理"
# 合同源文件目录：优先使用移动硬盘路径，未挂载时回退到 DATA_ROOT 下的目录。
CONTRACT_ROOT_OVERRIDE = Path("/Volumes/PC-SunYongtao/5_合同管理📑")
PROJECT_CODE_DB = str(MATCH_DB_PATH)

# 明细汇总（原 03）相关
TARGET_KEYWORDS = ["总包合同", "分包合同"]
DETAIL_OUTPUT_NAME = "总包合同+分包合同汇总.xlsx"

# 金额汇总（原 03.1）相关
GENERAL_CONTRACT_DIR = "总包合同"
SUBPACKAGE_CONTRACT_DIR = "分包合同"
AMOUNT_OUTPUT_NAME = "合同管理_总包分包汇总.xlsx"

# 合并输出
MERGED_OUTPUT_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/03_Sum/4_合同管理"
)
MERGED_OUTPUT_NAME = "合同管理_总包分包_汇总.xlsx"

# 分包台账逻辑名（实际可为 5.1_分包合同）
SUBPACKAGE_LEDGER_LOGICAL = "分包合同"
# 分包额对比容差（泰铢）；超过则着色标注（约等于原 0.01 万泰铢）
SUBPACKAGE_COMPARE_TOLERANCE_THB = 100.0

COLUMN_ORDER_FILE = str(Path(__file__).resolve().parent / "column_order.txt")

SETTLEMENT_CHECKS = [
    {
        "name": "竣工结算资料",
        "pattern": r"^H_[A-Z0-9-]+_竣工结算资料$",
        "empty_ok": False,
    },
]

OUTPUT_EXCLUDED_COLUMNS = {
    "对方单位公司名称",
    "对方公司名称",
    "竣工结算资料",
    "合同结算额",
    "结算金额",
    "合同签订日期",
    "计划开工时间",
    "计划竣工时间",
    "开工日期",
    "生效日期",
    "联系人及电话",
    "是否有业主结算资料",
    "付款方式",
    "关联主合同编号",
}


def strip_sort_prefix(name: str) -> str:
    """去掉名称里第一个 ``_`` / ``-`` 之前的序号或单字母前缀。"""
    text = str(name or "").strip()
    return _SORT_PREFIX_RE.sub("", text, count=1)


def logical_key(name: str) -> str:
    """用于比较的业务名：去排序前缀、去 emoji/符号。"""
    return dir_key(strip_sort_prefix(name)) or dir_key(name)


def find_named_folder(base: str | Path, logical: str) -> Path | None:
    """在 base 下找与 logical 匹配的子目录，忽略 ``_`` 之前的序号/字母。"""
    base_path = Path(base)
    if not base_path.is_dir() or not logical:
        return None

    found = find_named_dir(base_path, logical)
    if found is not None:
        return found

    target = logical_key(logical)
    if not target:
        return None

    exact: list[Path] = []
    suffix: list[Path] = []
    fuzzy: list[Path] = []
    for child in sorted(base_path.iterdir(), key=lambda p: p.name):
        if not child.is_dir() or child.name.startswith(".") or child.name.startswith("~$"):
            continue
        core = logical_key(child.name)
        if core == target:
            exact.append(child)
        elif child.name.endswith(logical) or strip_sort_prefix(child.name) == logical:
            suffix.append(child)
        elif target in core or core in target:
            fuzzy.append(child)
    if exact:
        return exact[0]
    if suffix:
        return suffix[0]
    if len(fuzzy) == 1:
        return fuzzy[0]
    return None


def resolve_named_dir(base: str | Path, logical: str) -> str:
    """返回匹配到的子目录名；找不到则退回 logical。"""
    found = find_named_folder(base, logical)
    return found.name if found is not None else logical


def resolve_existing_path(path: str, fallbacks: list[str] | None = None) -> str:
    """返回第一个存在的路径。"""
    candidates = [path] + (fallbacks or [])
    for p in candidates:
        if p and os.path.isdir(p):
            return p
    return path


def get_contract_root() -> str:
    if CONTRACT_ROOT_OVERRIDE.is_dir():
        return str(CONTRACT_ROOT_OVERRIDE)
    found = find_named_folder(DATA_ROOT, CONTRACT_ROOT_LOGICAL)
    if found is not None:
        return str(found)
    return resolve_existing_path(
        str(DATA_ROOT / f"5_{CONTRACT_ROOT_LOGICAL}📑"),
        [
            str(DATA_ROOT / f"5_{CONTRACT_ROOT_LOGICAL}"),
            str(DATA_ROOT / CONTRACT_ROOT_LOGICAL),
        ],
    )


def get_project_db() -> str:
    if os.path.isfile(PROJECT_CODE_DB):
        return PROJECT_CODE_DB
    alt = _CODE_ROOT / "common" / "project_name_match_db.json"
    return str(alt) if alt.is_file() else PROJECT_CODE_DB


def get_subpackage_std_dir() -> str:
    preferred = std_data_path(SUBPACKAGE_LEDGER_LOGICAL, create=False)
    if preferred.is_dir():
        return str(preferred)
    std_root = find_named_folder(DATA_ROOT, "Std")
    if std_root is not None:
        found = find_named_folder(std_root, SUBPACKAGE_LEDGER_LOGICAL)
        if found is not None:
            return str(found)
    return str(preferred)


def find_code_folder(logical: str) -> Path | None:
    """在代码根目录下按业务名找文件夹（如 5.1_分包合同台账处理）。"""
    return find_named_folder(_CODE_ROOT, logical)


def find_named_script(folder: Path, logical: str) -> Path | None:
    """在 folder 下找脚本，忽略 ``00_`` 这类前缀。"""
    if not folder.is_dir() or not logical:
        return None
    target = logical_key(Path(logical).stem)
    py_exact: list[Path] = []
    py_fuzzy: list[Path] = []
    for child in sorted(folder.iterdir(), key=lambda p: p.name):
        if not child.is_file() or child.suffix.lower() != ".py":
            continue
        if child.name.startswith("."):
            continue
        core = logical_key(child.stem)
        if core == target:
            py_exact.append(child)
        elif target and target in core:
            py_fuzzy.append(child)
    if py_exact:
        return py_exact[0]
    if len(py_fuzzy) == 1:
        return py_fuzzy[0]
    return None


if __name__ == "__main__":
    root = get_contract_root()
    print(f"合同根目录: {root}")
    print(f"  总包合同: {resolve_named_dir(root, GENERAL_CONTRACT_DIR)}")
    print(f"  分包合同: {resolve_named_dir(root, SUBPACKAGE_CONTRACT_DIR)}")
    print(f"分包台账std: {get_subpackage_std_dir()}")
    print(f"项目编码库: {get_project_db()}")
    ledger_code = find_code_folder("分包合同台账处理")
    print(f"台账处理脚本目录: {ledger_code}")
    if ledger_code:
        print(f"  全量合并: {find_named_script(ledger_code, '全量合并汇总')}")
