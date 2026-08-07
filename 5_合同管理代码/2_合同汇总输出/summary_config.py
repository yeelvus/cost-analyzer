# -*- coding: utf-8 -*-
"""合同汇总公共配置。"""

from __future__ import annotations

import os
from pathlib import Path

# 合同根目录 / 项目编码库（可按本机环境修改）
CONTRACT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑"
PROJECT_CODE_DB = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"

# 明细汇总（原 03）相关
TARGET_KEYWORDS = ["总包合同", "分包合同"]
DETAIL_OUTPUT_NAME = "总包合同+分包合同汇总.xlsx"

# 金额汇总（原 03.1）相关
GENERAL_CONTRACT_DIR = "总包合同"
SUBPACKAGE_CONTRACT_DIR = "分包合同"
AMOUNT_OUTPUT_NAME = "合同管理_总包分包汇总.xlsx"

# 合并输出
MERGED_OUTPUT_NAME = "合同管理_总包分包全量汇总.xlsx"

# 分包台账标准化目录（用于与文件夹汇总的分包额交叉核对）
SUBPACKAGE_STD_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/2_分包合同台账_std"
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


def resolve_existing_path(path: str, fallbacks: list[str] | None = None) -> str:
    """返回第一个存在的路径。"""
    candidates = [path] + (fallbacks or [])
    for p in candidates:
        if p and os.path.isdir(p):
            return p
    return path


def get_contract_root() -> str:
    return resolve_existing_path(
        CONTRACT_ROOT,
        [
            os.path.expanduser("~/Desktop/邯郸中材/01成本统计/5_合同管理📑"),
        ],
    )


def get_project_db() -> str:
    if os.path.isfile(PROJECT_CODE_DB):
        return PROJECT_CODE_DB
    alt = os.path.expanduser("~/Desktop/code/0_项目基本信息/project_name_match_db.json")
    return alt if os.path.isfile(alt) else PROJECT_CODE_DB


def get_subpackage_std_dir() -> str:
    return resolve_existing_path(
        SUBPACKAGE_STD_DIR,
        [
            os.path.expanduser("~/Desktop/邯郸中材/01成本统计/5.1_分包合同/2_分包合同台账_std"),
        ],
    )
