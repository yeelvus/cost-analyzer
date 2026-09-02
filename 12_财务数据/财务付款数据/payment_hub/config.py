# -*- coding: utf-8 -*-
"""财务付款一体化默认路径配置（可用 CLI 覆盖）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

CODE_DIR = Path(__file__).resolve().parent.parent  # 财务付款数据/
DESKTOP_COST = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
PAY_ROOT = DESKTOP_COST / "12-财务数据💰" / "A_财务付款数据"
REPORT_ROOT = DESKTOP_COST / "Z_report"


@dataclass
class Paths:
    """全流程路径。"""

    raw_dir: Path = field(default_factory=lambda: PAY_ROOT / "A_财务付款_raw")
    # raw 标准化合并结果（第一步 raw_std）
    raw_std_dir: Path = field(
        default_factory=lambda: PAY_ROOT / "A1_财务付款_raw_std"
    )
    # 摘要分类结果（第二步 classify，供 sum / std 复用）
    classified_dir: Path = field(
        default_factory=lambda: PAY_ROOT / "A2_财务付款_classified"
    )
    std_dir: Path = field(default_factory=lambda: PAY_ROOT / "B_财务付款_std")

    sum_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT / "C_财务付款_sum" / "财务付款数据汇总.xlsx"
    )
    amount_mismatch_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT
        / "C_财务付款_sum"
        / "财务付款_应付与预扣实付不平明细.xlsx"
    )

    cashflow_html: Path = field(
        default_factory=lambda: REPORT_ROOT
        / "财务付款"
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
    # 账簿：已填一/二/三级 → learn-ledger 写入记忆
    ledger_xlsx: Path = field(
        default_factory=lambda: Path(
            "/Users/wyattsun/Desktop/财务数据/账簿汇总_v7_分类整理2026-08-29_副本.xlsx"
        )
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
    # 兼容旧路径；00_汇总现拆为 4 个：有/无编码 × 来源分组
    learn_pool_all_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "00_汇总"
        / "全部摘要_学习池.xlsx"
    )
    learn_pool_by_project: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "01_按项目"
    )
    # 无正式项目编码：合并待核对，不与正式项目混排
    learn_pool_no_code: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "01b_待核对_无项目编码"
    )
    learn_pool_no_code_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "01b_待核对_无项目编码"
        / "待核对_无项目编码汇总.xlsx"
    )
    learn_pool_by_month: Path = field(
        default_factory=lambda: PAY_ROOT / "D_财务付款_学习池" / "02_按月份"
    )
    learn_pool_priority_xlsx: Path = field(
        default_factory=lambda: PAY_ROOT
        / "D_财务付款_学习池"
        / "00_汇总"
        / "00_处理优先级_总览.xlsx"
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

    @property
    def sum_dir(self) -> Path:
        """C_财务付款_sum 根目录（分组输出写在其下一级子目录）。"""
        return self.sum_xlsx.parent

    def group_raw_std_dir(self, group: str) -> Path:
        """A1_raw_std/{组名}/"""
        name = (group or "").strip() or "未分组"
        return self.raw_std_dir / name

    def group_raw_std_xlsx(self, group: str) -> Path:
        """A1_raw_std/{组名}/{组名}_财务付款_raw_std.xlsx"""
        name = (group or "").strip() or "未分组"
        return self.group_raw_std_dir(name) / f"{name}_财务付款_raw_std.xlsx"

    def group_classified_dir(self, group: str) -> Path:
        """A2_classified/{组名}/"""
        name = (group or "").strip() or "未分组"
        return self.classified_dir / name

    def group_classified_xlsx(self, group: str) -> Path:
        """A2_classified/{组名}/{组名}_classified.xlsx"""
        name = (group or "").strip() or "未分组"
        return self.group_classified_dir(name) / f"{name}_classified.xlsx"

    def group_std_dir(self, group: str) -> Path:
        """B_std/{组名}/"""
        name = (group or "").strip() or "未分组"
        return self.std_dir / name

    def group_sum_dir(self, group: str) -> Path:
        """C_sum/{组名}/"""
        name = (group or "").strip() or "未分组"
        return self.sum_dir / name

    def group_sum_xlsx(self, group: str) -> Path:
        """C_sum/{组名}/{组名}_财务付款数据汇总.xlsx"""
        name = (group or "").strip() or "未分组"
        return self.group_sum_dir(name) / f"{name}_财务付款数据汇总.xlsx"

    def group_cashflow_html(self, group: str) -> Path:
        name = (group or "").strip() or "未分组"
        # 与汇总命名一致：文件夹名_原文件名
        stem = Path(self.cashflow_html).stem
        suffix = Path(self.cashflow_html).suffix
        return Path(self.cashflow_html).parent / f"{name}_{stem}{suffix}"

    def group_amount_mismatch_xlsx(self, group: str) -> Path:
        name = (group or "").strip() or "未分组"
        stem = Path(self.amount_mismatch_xlsx).stem
        suffix = Path(self.amount_mismatch_xlsx).suffix
        return self.group_sum_dir(name) / f"{name}_{stem}{suffix}"

    def ensure_output_parents(self) -> None:
        self.raw_std_dir.mkdir(parents=True, exist_ok=True)
        self.classified_dir.mkdir(parents=True, exist_ok=True)
        self.std_dir.mkdir(parents=True, exist_ok=True)
        self.sum_dir.mkdir(parents=True, exist_ok=True)
        self.cashflow_html.parent.mkdir(parents=True, exist_ok=True)
        self.classify_data.mkdir(parents=True, exist_ok=True)
        self.class_master_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.class_master_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
        self.learn_log_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.learn_log_xlsx_sum.parent.mkdir(parents=True, exist_ok=True)
        self.learn_pool_all.mkdir(parents=True, exist_ok=True)
        self.learn_pool_by_project.mkdir(parents=True, exist_ok=True)
        self.learn_pool_no_code.mkdir(parents=True, exist_ok=True)
        self.learn_pool_by_month.mkdir(parents=True, exist_ok=True)
        self.learn_pool_confirmed.parent.mkdir(parents=True, exist_ok=True)
