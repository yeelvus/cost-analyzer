# -*- coding: utf-8 -*-
"""财务付款一体化默认路径配置（可用 CLI 覆盖）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

CODE_DIR = Path(__file__).resolve().parent.parent  # 财务付款数据/
DESKTOP_COST = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
PAY_ROOT = DESKTOP_COST / "12-财务数据💰" / "A_财务付款数据"


@dataclass
class Paths:
    """全流程路径。"""

    raw_dir: Path = field(default_factory=lambda: PAY_ROOT / "A_财务付款_raw")
    std_dir: Path = field(default_factory=lambda: PAY_ROOT / "B_财务付款_std")

    sum_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT / "C_财务付款_sum" / "财务付款数据汇总.xlsx"
    )
    sum_xlsx_z: Path = field(
        default_factory=lambda: DESKTOP_COST / "Z_report" / "财务付款数据汇总.xlsx"
    )

    cashflow_html: Path = field(
        default_factory=lambda: PAY_ROOT
        / "C_财务付款_sum"
        / "财务付款项目现金流量分析.html"
    )
    project_info_xlsx: Path = field(
        default_factory=lambda: DESKTOP_COST
        / "0_项目基本信息情况"
        / "项目基本信息表.xlsx"
    )

    classify_data: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data"
    )
    review_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "待复核池.xlsx"
    )
    rules_suggest_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "规则建议.xlsx"
    )

    def ensure_output_parents(self) -> None:
        self.std_dir.mkdir(parents=True, exist_ok=True)
        self.sum_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.sum_xlsx_z.parent.mkdir(parents=True, exist_ok=True)
        self.cashflow_html.parent.mkdir(parents=True, exist_ok=True)
        self.classify_data.mkdir(parents=True, exist_ok=True)
