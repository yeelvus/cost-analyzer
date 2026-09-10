# -*- coding: utf-8 -*-
"""分包合同台账处理 - 公共路径配置。"""

from __future__ import annotations

import sys
from pathlib import Path

# 代码根目录（含 common / settings）
CODE_ROOT = Path(__file__).resolve().parents[1]
if str(CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(CODE_ROOT))

from settings import DATA_ROOT, MATCH_DB_PATH  # noqa: E402
from common.paths import raw_data_path  # noqa: E402

# 业务目录
SUBPACKAGE_ROOT = DATA_ROOT / "5.1_分包合同"
RAW_DIR = raw_data_path("分包合同", legacy=SUBPACKAGE_ROOT / "1_分包合同台账_raw")
STD_DIR = SUBPACKAGE_ROOT / "2_分包合同台账_std"
SUM_DIR = SUBPACKAGE_ROOT / "3_分包合同_sum"

# 项目基本信息
PROJECT_INFO_XLSX = DATA_ROOT / "0_项目基本信息情况" / "项目基本信息表.xlsx"

# 输出文件名
MERGED_OUTPUT_NAME = "分包合同台账_全量汇总.xlsx"
COMPARISON_ONLY_NAME = "分包合同汇总逐月对比汇总.xlsx"
PAID_ONLY_SUFFIX = "_项目汇总.xlsx"

# 对比容差（泰铢）
AMOUNT_TOLERANCE = 1.0
