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
    # 分类总表：一份 Excel 调全部分类（也复制到 C_sum 便于查找）
    class_master_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "付款分类总表.xlsx"
    )
    class_master_xlsx_sum: Path = field(
        default_factory=lambda: PAY_ROOT / "C_财务付款_sum" / "付款分类总表.xlsx"
    )
    taxonomy_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "科目体系.xlsx"
    )
    # 每次 apply-class 生成的可读学习报告
    learn_log_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "学习记录.xlsx"
    )
    learn_log_xlsx_sum: Path = field(
        default_factory=lambda: PAY_ROOT / "C_财务付款_sum" / "学习记录.xlsx"
    )
    # 分类质量/修正报告（含仍待人工_高频）
    class_fix_report_xlsx: Path = field(
        default_factory=lambda: CODE_DIR / "learn_classify" / "data" / "分类修正记录.xlsx"
    )
    class_fix_report_xlsx_sum: Path = field(
        default_factory=lambda: PAY_ROOT / "C_财务付款_sum" / "分类修正记录.xlsx"
    )

    # 学习池：汇总 / 按项目 / 按月份 + 已确认总表
    learn_pool_root: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池"
    )
    learn_pool_all: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "00_汇总"
    )
    learn_pool_all_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "00_汇总"
        / "全部摘要_学习池.xlsx"
    )
    learn_pool_by_project: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "01_按项目"
    )
    learn_pool_by_month: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "02_按月份"
    )
    learn_pool_confirmed: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "03_已确认"
        / "已确认分类.xlsx"
    )
    learn_pool_snapshot: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "_export_snapshot.json"
    )
    learn_pool_conflicts: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "03_已确认"
        / "冲突记录.xlsx"
    )

    def ensure_output_parents(self) -> None:
        self.std_dir.mkdir(parents=True, exist_ok=True)
        self.sum_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.sum_xlsx_z.parent.mkdir(parents=True, exist_ok=True)
        self.cashflow_html.parent.mkdir(parents=True, exist_ok=True)
        self.classify_data.mkdir(parents=True, exist_ok=True)
        self.class_master_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.class_master_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
        self.learn_log_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.learn_log_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
        self.learn_pool_all.mkdir(parents=True, exist_ok=True)
        self.learn_pool_by_project.mkdir(parents=True, exist_ok=True)
        self.learn_pool_by_month.mkdir(parents=True, exist_ok=True)
        self.learn_pool_confirmed.parent.mkdir(parents=True, exist_ok=True)
