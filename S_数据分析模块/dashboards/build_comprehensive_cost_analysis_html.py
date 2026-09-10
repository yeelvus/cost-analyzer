#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目成本统计综合分析 HTML 报告生成器（多层级钻取版）

================================================================================
成本数据收集与核算整体思路（重要背景，请维护此说明）
================================================================================
1. 原始来源（双轨并行）：
   a. 项目每月填写的「成本统计模板」（2_成本统计收集/A_成本统计_raw/2026年XX月/*_成本统计_*.xlsm）
      - 固定 sheet 结构：概况统计、1分包款、2材料汇总、2-1采购部采购、2-2项目部采购、
        3材料租赁、4，5工资、6机具费、7人员调遣、8其他、9临建费用、业主预计结算额 等。
      - 模板里有「实际已发生」+「后续预估发生」两套口径，便于全周期成本控制。
   b. 各专业部门独立维护的原始台账（采购明细、机具使用记录、人工费台账、财务付款流水、
      合同/分包合同文件、搅拌站混凝土等），经过各自标准化管道产出「B_xxx_sum」或「按项目拆分」下的 per-project 文件。

2. 中央聚合与对账（核心核算逻辑）：
   - 2_成本代码/成本_汇总按月收集成本统计.py 使用「声明式 CONFIG」+ ProjectProcessor.evaluate_expression
     （openpyxl 直接读 cell 或 sum 列），从每个项目的成本统计模板里「刮取」关键指标。
   - 同时加载各领域**最新**汇总文件（load_project_amount_map + find_latest_*），按项目编码/标准名 join。
   - 形成「项目统计数据汇总_全月份.xlsx」（master）：既有「项目自己报的数」（预计总成本、毛利、各项支出、后续预估），
     又有「台账客观汇总」（采购数据汇总、机具费汇总、人工费汇总、财务分包已付款、总包/分包合同额等）。
   - 成本_数据分析.py 进一步做差异核对（采购部采购金额 vs 采购数据汇总、工资+调遣 vs 人工费汇总、机具支出 vs 机具汇总、
     预估无变化提醒等），暴露需要人工复核的地方。

3. 人/材/机具体数据：
   - 「人」：11_人工费.../B_人工费台账_std/按项目拆分/*_人工费.xlsx （月度汇总 sheet 含 中方正式工铢、中方劳务工铢、
     子公司属地工铢、劳务公司铢、签证费、总人工费 等明细 + 原始数据）。
   - 「材」：9_采购台账.../E_采购部采购台账_sum/采购部采购台账_std_merged.xlsx （项目汇总 总采购金额） + master 中的
     材料采购实际已发生 / 采购部采购金额 / 项目部采购金额 + 模板里的 2材料汇总 行项目。
   - 「机」：10_机具费.../B_机具费_sum/*_机具费_汇总.xlsx（费用合计、使用次数、机具种类数、出租单位） + 模板 6机具费 sheet
     的类目/机具名称/型号/已发生/后续预估。

4. 标准化层：0_项目基本信息/project_name_matcher + JSON DB，把各种「万源通」「糖厂项目_成本统计」等别名映射到标准名+编码。

5. 报告定位：本脚本读取 master + 上述 per-project 明细文件，生成**单一多层级交互 HTML**：
   - 第 0 层：全局概览 + 可筛选项目总表 + 趋势/结构图。
   - 第 1 层：点击任意项目 → 进入「项目深度分析面板」（支持切换不同项目）。
   - 第 2 层：面板内多 Tab（概况 / 人工(人) / 采购(材) / 机具(机) / 合同财务 / 历史），每个 Tab 展示该维度的具体数据 + 可视化。
   这解决了「html不只是一个层级」和「选择项目后要看到具体的人材机数据」的需求。

输出默认：0_html/项目成本统计综合分析报告.html
可被 2_成本代码/0_一键运行_成本流程.py 自动调用（在三个成本脚本之后）。

运行示例见文件末尾。
依赖：pandas, openpyxl（通常已有）。Chart.js 通过 CDN。
================================================================================
"""


from __future__ import annotations

# --- path bootstrap (S_数据分析模块) ---
import sys
from pathlib import Path as _PathForBootstrap
try:
    from S_数据分析模块._bootstrap import ensure_code_root, ensure_script_dir
    ensure_code_root(_PathForBootstrap(__file__))
    ensure_script_dir(__file__)
except Exception:
    _p = _PathForBootstrap(__file__).resolve()
    for _c in [_p.parent, *_p.parents]:
        if (_c / "settings.py").exists():
            if str(_c) not in sys.path:
                sys.path.insert(0, str(_c))
            break
    if str(_p.parent) not in sys.path:
        sys.path.insert(0, str(_p.parent))
# --- end bootstrap ---

import argparse
import html as html_escape
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# 复用项目设置与路径
try:
    from settings import DATA_ROOT, REPORT_ROOT, data_path, report_path
    from common.paths import html_data_path
    from common.paths import find_latest_excel_file as _common_find_latest  # type: ignore
except Exception:
    DATA_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
    REPORT_ROOT = DATA_ROOT / "Z_report"
    data_path = lambda *p: DATA_ROOT.joinpath(*p)
    report_path = lambda *p: REPORT_ROOT.joinpath(*p)
    _common_find_latest = None

    def html_data_path(*parts, create=True):
        p = DATA_ROOT.joinpath("0_html", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p


# ==================== 默认路径（与现有成本流程对齐） ====================
DEFAULT_SUMMARY_DIR = DATA_ROOT / "2_成本统计收集" / "B_成本统计_sum"
DEFAULT_INPUT_GLOB = "*项目统计数据汇总*全月份*.xlsx"
DEFAULT_OUTPUT_HTML = html_data_path() / "项目成本统计综合分析报告.html"

# 数值列（与成本_数据分析.py 保持一致，扩展）
NUMERIC_COLS = [
    "预计总成本", "计算总成本", "毛利", "毛利率", "实际已发生金额", "业主预计结算额",
    "总包合同额", "业主签证", "分包结算额", "分包合同额", "分包签证", "财务分包已付款",
    "材料采购实际已发生", "采购部采购金额", "采购数据汇总", "项目部采购金额",
    "商混站混凝土汇总", "材料租赁费用", "工资支出", "人员调遣支出", "人工费汇总",
    "机具费支出", "机具费汇总", "其他费用支出", "临建费用支出", "后续预估汇总",
    "材料采购后续预估发生", "材料租赁后续预估", "机具费后续预估",
    "人员调遣后续预估", "其他费用后续预估", "临建费用后续预估",
    "工程部记录分包已付款",
]

KEY_CATEGORY_COLS = [
    "材料采购实际已发生", "采购数据汇总", "项目部采购金额",
    "机具费支出", "机具费汇总",
    "工资支出", "人工费汇总",
    "财务分包已付款", "工程部记录分包已付款",
    "商混站混凝土汇总",
    "其他费用支出", "临建费用支出",
]


@dataclass
class ProjectLatest:
    name: str
    code: str
    month: str
    est_cost: float
    gross_profit: float
    gross_margin: float
    settlement: float
    contract_total: float
    subcontract_contract: float
    finance_paid: float
    procurement: float
    machinery: float
    labor: float
    actual_occurred: float
    forecast_total: float


def safe_float(v: Any) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).strip().replace(",", "").replace("，", "").replace("%", "")
        if not s or s.lower() in ("nan", "none", "-"):
            return 0.0
        return float(s)
    except Exception:
        return 0.0


def normalize_margin(v: Any) -> Optional[float]:
    val = safe_float(v)
    if val == 0.0:
        return None
    if 1 < abs(val) <= 100:
        val = val / 100.0
    return val


def parse_month_sort_key(m: str) -> Tuple[int, int]:
    import re
    mth = re.match(r"^(\d{4})年(\d{1,2})月$", str(m).strip())
    if mth:
        return (int(mth.group(1)), int(mth.group(2)))
    return (0, 0)


def find_latest_master_excel(explicit: Optional[str] = None) -> Path:
    if explicit:
        p = Path(explicit).expanduser().resolve()
        if p.exists():
            return p
    # 优先 SUMMARY_DIR
    candidates: List[Path] = []
    for base in [DEFAULT_SUMMARY_DIR, REPORT_ROOT, DATA_ROOT]:
        if base.exists():
            for pat in [DEFAULT_INPUT_GLOB, "*项目统计数据汇总*.xlsx"]:
                candidates.extend([f for f in base.rglob(pat) if f.is_file() and not f.name.startswith("~$")])
    if not candidates:
        # 回退：任意 2_成本统计收集 下的
        for base in [DATA_ROOT / "2_成本统计收集", DATA_ROOT]:
            if base.exists():
                candidates.extend([f for f in base.rglob("*.xlsx") if "项目统计数据汇总" in f.name and not f.name.startswith("~$")])
    if not candidates:
        raise FileNotFoundError(f"未找到主汇总Excel，请指定 --input-excel 或确认路径: {DEFAULT_SUMMARY_DIR}")
    # 最新修改时间优先
    latest = max(candidates, key=lambda p: p.stat().st_mtime)
    print(f"✓ 自动选择主数据文件: {latest}")
    return latest


def load_and_prepare(input_excel: Path) -> Tuple[pd.DataFrame, List[str], str]:
    df = pd.read_excel(input_excel, engine="openpyxl")
    # 标准化列名
    df.columns = [str(c).strip() for c in df.columns]

    # 数值化
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = df[col].apply(safe_float)

    if "毛利率" in df.columns:
        df["毛利率_norm"] = df["毛利率"].apply(normalize_margin)
    else:
        df["毛利率_norm"] = None

    # 月份排序键
    df["_month_key"] = df.get("月份", "").apply(parse_month_sort_key)

    months = sorted(df["月份"].dropna().unique().tolist(), key=parse_month_sort_key, reverse=True)
    latest_month = months[0] if months else ""

    # 确保关键列存在
    for need in ["项目名称", "项目编码", "月份"]:
        if need not in df.columns:
            df[need] = ""

    return df, months, latest_month


def build_project_latest(df: pd.DataFrame, latest_month: str) -> List[ProjectLatest]:
    latest_df = df[df["月份"] == latest_month].copy()
    if latest_df.empty and not df.empty:
        # 取每个项目最近一条
        latest_df = df.sort_values("_month_key", ascending=False).groupby("项目名称", as_index=False).first()

    projs: List[ProjectLatest] = []
    for _, r in latest_df.iterrows():
        projs.append(ProjectLatest(
            name=str(r.get("项目名称", "")).strip(),
            code=str(r.get("项目编码", "")).strip(),
            month=str(r.get("月份", "")),
            est_cost=safe_float(r.get("预计总成本")),
            gross_profit=safe_float(r.get("毛利")),
            gross_margin=safe_float(r.get("毛利率_norm")) or safe_float(r.get("毛利率")),
            settlement=safe_float(r.get("业主预计结算额")),
            contract_total=safe_float(r.get("总包合同额")),
            subcontract_contract=safe_float(r.get("分包合同额")),
            finance_paid=safe_float(r.get("财务分包已付款")) or safe_float(r.get("工程部记录分包已付款")),
            procurement=safe_float(r.get("采购数据汇总")) or safe_float(r.get("采购部采购金额")),
            machinery=safe_float(r.get("机具费汇总")) or safe_float(r.get("机具费支出")),
            labor=safe_float(r.get("人工费汇总")) or safe_float(r.get("工资支出")),
            actual_occurred=safe_float(r.get("实际已发生金额")),
            forecast_total=safe_float(r.get("后续预估汇总")),
        ))
    # 去重保序
    seen = set()
    uniq = []
    for p in projs:
        if p.name and p.name not in seen:
            seen.add(p.name)
            uniq.append(p)
    return uniq


def compute_global_stats(df: pd.DataFrame, latest_month: str, projects: List[ProjectLatest]) -> Dict[str, Any]:
    latest_rows = df[df["月份"] == latest_month] if latest_month else pd.DataFrame()
    total_est = sum(p.est_cost for p in projects)
    total_settle = sum(p.settlement for p in projects)
    total_profit = sum(p.gross_profit for p in projects)
    margins = [p.gross_margin for p in projects if p.gross_margin is not None and p.gross_margin != 0]
    avg_margin = sum(margins) / len(margins) if margins else 0.0
    total_finance = sum(p.finance_paid for p in projects)
    total_proc = sum(p.procurement for p in projects)
    total_mach = sum(p.machinery for p in projects)
    total_labor = sum(p.labor for p in projects)

    # 历史整体毛利率趋势（简单月均）
    margin_by_month: Dict[str, List[float]] = {}
    for _, row in df.iterrows():
        m = str(row.get("月份", ""))
        mv = row.get("毛利率_norm") if "毛利率_norm" in df.columns else normalize_margin(row.get("毛利率"))
        if mv is None or mv == 0:
            continue
        margin_by_month.setdefault(m, []).append(float(mv))
    month_margin_avg = {m: sum(v) / len(v) for m, v in margin_by_month.items() if v}

    return {
        "project_count": len(projects),
        "latest_month": latest_month,
        "total_est_cost_wan": round(total_est, 2),
        "total_settlement_wan": round(total_settle, 2),
        "total_gross_profit_wan": round(total_profit, 2),
        "avg_margin": round(avg_margin, 4),
        "total_finance_paid_wan": round(total_finance, 2),
        "total_procurement_wan": round(total_proc, 2),
        "total_machinery_wan": round(total_mach, 2),
        "total_labor_wan": round(total_labor, 2),
        "month_margin_avg": month_margin_avg,
        "data_rows": len(df),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def prepare_trend_data(df: pd.DataFrame, months: List[str]) -> Dict[str, Any]:
    """为多项目趋势准备数据（限制项目数避免过大）"""
    # 取有毛利率的项目，最近有数据的优先
    proj_month_margin: Dict[str, Dict[str, Optional[float]]] = {}
    proj_month_cost: Dict[str, Dict[str, Optional[float]]] = {}

    for _, row in df.iterrows():
        p = str(row.get("项目名称", "")).strip()
        if not p:
            continue
        m = str(row.get("月份", ""))
        if m not in months:
            continue
        mv = row.get("毛利率_norm") if "毛利率_norm" in df.columns else normalize_margin(row.get("毛利率"))
        proj_month_margin.setdefault(p, {})[m] = mv
        proj_month_cost.setdefault(p, {})[m] = safe_float(row.get("预计总成本")) or safe_float(row.get("实际已发生金额"))

    # 选代表性项目：数据最全 + 有最近月份
    def score(name: str) -> Tuple[int, int]:
        mm = proj_month_margin.get(name, {})
        non_null = sum(1 for v in mm.values() if v is not None)
        has_latest = 1 if (months and mm.get(months[0])) else 0
        return (-non_null, -has_latest)  # 排序用，越小越好

    ordered = sorted(proj_month_margin.keys(), key=score)[:12]  # 最多12个项目，避免图例爆炸

    labels = [m for m in sorted(months, key=parse_month_sort_key)]  # 时间正序
    series_margin = {}
    series_cost = {}
    for p in ordered:
        series_margin[p] = [proj_month_margin[p].get(m) for m in labels]
        series_cost[p] = [proj_month_cost[p].get(m) for m in labels]

    return {
        "labels": labels,
        "margin": series_margin,
        "cost": series_cost,
    }


def prepare_category_breakdown(projects: List[ProjectLatest]) -> Dict[str, Any]:
    cats = {
        "采购": sum(p.procurement for p in projects),
        "机具": sum(p.machinery for p in projects),
        "人工": sum(p.labor for p in projects),
        "财务分包付款": sum(p.finance_paid for p in projects),
        "其他/临建": sum(p.est_cost * 0.1 for p in projects),  # 兜底简化
    }
    # 更精确：用最新行的各列求和（如果有）
    total = sum(cats.values()) or 1.0
    return {
        "labels": list(cats.keys()),
        "values": [round(v, 2) for v in cats.values()],
        "percent": [round(v / total * 100, 1) for v in cats.values()],
    }


def prepare_project_cards_data(df: pd.DataFrame, latest_month: str, max_projects: int = 20) -> Tuple[List[str], Dict[str, Any], List[str]]:
    """准备类似人工费/公司月报风格的 per-project 卡片数据"""
    cards = []
    chart_payload: Dict[str, Any] = {}
    toc: List[str] = []

    # 取最新 + 历史
    proj_groups = df.groupby("项目名称")
    ordered = []
    for name, g in proj_groups:
        g = g.sort_values("_month_key", ascending=False)
        if not g.empty:
            ordered.append((name, g))

    # 优先显示有最新月份或数据较全的
    ordered.sort(key=lambda x: (x[1].iloc[0]["月份"] != latest_month, -len(x[1])), reverse=False)

    for idx, (pname, g) in enumerate(ordered[:max_projects], 1):
        g_sorted = g.sort_values("_month_key")  # 时间升序用于图表
        months_local = g_sorted["月份"].tolist()
        margin_s = [normalize_margin(v) for v in g_sorted.get("毛利率", [])]
        cost_s = [safe_float(v) for v in g_sorted.get("预计总成本", g_sorted.get("实际已发生金额", []))]
        settle_s = [safe_float(v) for v in g_sorted.get("业主预计结算额", [])]
        profit_s = [safe_float(v) for v in g_sorted.get("毛利", [])]

        card_id = f"card_{idx}"
        toc.append(pname)
        chart_payload[card_id] = {
            "months": months_local,
            "margin": margin_s,
            "cost": cost_s,
            "settlement": settle_s,
            "profit": profit_s,
        }

        latest = g.iloc[0]
        cards.append({
            "id": card_id,
            "name": pname,
            "code": str(latest.get("项目编码", "")),
            "latest_month": str(latest.get("月份", "")),
            "est_cost": safe_float(latest.get("预计总成本")),
            "margin": normalize_margin(latest.get("毛利率")),
            "settlement": safe_float(latest.get("业主预计结算额")),
            "finance_paid": safe_float(latest.get("财务分包已付款")) or safe_float(latest.get("工程部记录分包已付款")),
            "rows_html": _build_mini_table_rows(g_sorted.tail(6)),  # 最近6期
        })

    return toc, chart_payload, cards


def _build_mini_table_rows(g_tail: pd.DataFrame) -> str:
    rows = []
    for _, r in g_tail.iterrows():
        m = str(r.get("月份", ""))
        est = safe_float(r.get("预计总成本"))
        mar = normalize_margin(r.get("毛利率"))
        mar_str = f"{mar*100:.1f}%" if mar is not None else "-"
        settle = safe_float(r.get("业主预计结算额"))
        rows.append(
            f"<tr><td>{m}</td>"
            f"<td class='num'>{est:,.1f}</td>"
            f"<td class='num'>{mar_str}</td>"
            f"<td class='num'>{settle:,.1f}</td></tr>"
        )
    return "".join(rows)


# =============================================================================
# 新增：领域明细加载器（支持多层级 HTML 的「人/材/机」钻取数据）
# 这些函数让报告不再只是 master 聚合层，而是能进入具体的人工月度构成、采购金额、机具使用等。
# =============================================================================

def find_latest_in_dir(dir_path: Path, patterns: List[str], prefer_keywords: List[str] = None) -> Optional[Path]:
    """通用最新文件查找（按修改时间 + 名称时间戳优先）。"""
    if not dir_path.exists():
        return None
    files = []
    for pat in patterns:
        files.extend([p for p in dir_path.rglob(pat) if p.is_file() and not p.name.startswith("~$")])
    if not files:
        return None
    # 简单按 mtime 排序（collector 里有更复杂的 ts 解析，可按需加强）
    return max(files, key=lambda p: p.stat().st_mtime)


def normalize_project_code(value):
    """小工具：项目编码归一（复制自 collector，避免依赖）。"""
    if value is None or pd.isna(value):
        return None
    code = str(value).strip()
    if not code:
        return None
    if re.fullmatch(r"\d+\.0+", code):
        code = code.split('.', 1)[0]
    return code

def load_procurement_map() -> Dict[str, float]:
    """加载最新采购项目汇总（材）。优先 E_采购部采购台账_sum 下的 merged/标准版。"""
    base = DATA_ROOT / "9_采购台账📦" / "E_采购部采购台账_sum"
    candidates = [
        base / "采购部采购台账_std_merged.xlsx",
        base / "采购部采购台账_标准版.xlsx",
        base / "采购部采购台账_merged.xlsx",
    ]
    for cand in candidates:
        if cand.exists():
            try:
                df = pd.read_excel(cand, sheet_name="项目汇总")
                df.columns = [str(c).strip() for c in df.columns]
                code_col = next((c for c in ["项目编码", "标准项目编码", "项目代码"] if c in df.columns), None)
                amt_col = next((c for c in ["总采购金额", "采购金额", "金额"] if c in df.columns), None)
                if code_col and amt_col:
                    m = {}
                    for _, r in df.iterrows():
                        code = normalize_project_code(r.get(code_col))
                        if code:
                            m[code] = safe_float(r.get(amt_col)) / 10000.0  # 转万
                    print(f"✓ 已加载采购汇总（材）: {cand.name}，{len(m)} 项目")
                    return m
            except Exception as e:
                print(f"  采购文件加载失败 {cand.name}: {e}")
    print("⚠️ 未找到或无法解析采购项目汇总（材数据将退化到 master 字段）")
    return {}


def find_per_project_labor_file(project_code: str, project_name: str) -> Optional[Path]:
    """在人工费 按项目拆分 目录找匹配的 per-project 明细文件（人）。"""
    labor_dir = DATA_ROOT / "11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std/按项目拆分"
    if not labor_dir.exists():
        return None
    # 优先用编码匹配
    if project_code and project_code != "未匹配":
        for p in labor_dir.glob(f"*{project_code}*_人工费.xlsx"):
            return p
    # 再用名称关键词
    safe_name = project_name.replace("/", "_").replace(" ", "")
    for p in labor_dir.glob(f"*{safe_name}*_人工费.xlsx"):
        return p
    # 宽松包含
    for p in labor_dir.glob("*_人工费.xlsx"):
        if project_code and project_code in p.name:
            return p
        if any(tok in p.name for tok in safe_name.split("项目")[:1] if len(tok) > 1):
            return p
    return None


def load_labor_detail_for_project(project_code: str, project_name: str) -> Dict[str, Any]:
    """读取单个项目的劳动费月度明细（人）。返回 {months: [], breakdown: {type: [vals]}, total: [vals]}"""
    f = find_per_project_labor_file(project_code, project_name)
    if not f or not f.exists():
        return {}
    try:
        xls = pd.ExcelFile(f)
        sheet = "月度汇总" if "月度汇总" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(f, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        # 常见列
        month_col = next((c for c in ["年-月", "月份", "年月", "month"] if c in df.columns), None)
        if not month_col:
            return {}
        types = ["中方正式工铢", "中方劳务工铢", "子公司属地工铢", "劳务公司铢", "签证费", "总人工费"]
        avail_types = [t for t in types if t in df.columns]
        months = []
        series: Dict[str, List[float]] = {t: [] for t in avail_types}
        for _, r in df.iterrows():
            m = str(r.get(month_col, "")).strip()
            if not m or m.lower() == "nan":
                continue
            months.append(m)
            for t in avail_types:
                series[t].append(safe_float(r.get(t)))
        print(f"  ✓ 加载人工明细（人）: {f.name}")
        return {"file": str(f.name), "months": months, "breakdown": series}
    except Exception as e:
        print(f"  加载人工明细失败 {project_name}: {e}")
        return {}


def find_per_project_machine_file(project_code: str, project_name: str) -> Optional[Path]:
    """找机具 per-project 汇总（机）。"""
    mach_dir = DATA_ROOT / "10_机具费🚜/B_机具费_sum"
    if not mach_dir.exists():
        return None
    if project_code and project_code != "未匹配":
        for p in mach_dir.glob(f"*{project_code}*_机具费_汇总.xlsx"):
            return p
    for p in mach_dir.glob("*_机具费_汇总.xlsx"):
        if project_code and project_code in p.name:
            return p
        if project_name and project_name.split("项目")[0] in p.name:
            return p
    return None


def load_machine_detail_for_project(project_code: str, project_name: str) -> Dict[str, Any]:
    f = find_per_project_machine_file(project_code, project_name)
    if not f or not f.exists():
        return {}
    try:
        df = pd.read_excel(f)
        df.columns = [str(c).strip() for c in df.columns]
        # 这些文件通常是单行汇总（费用合计 + 次数等）
        row = df.iloc[0] if len(df) > 0 else {}
        detail = {
            "file": str(f.name),
            "total_thb": safe_float(row.get("费用合计(泰铢)")),
            "use_count": safe_float(row.get("使用次数")),
            "kind_count": safe_float(row.get("机具种类数")),
            "supplier_count": safe_float(row.get("出租单位数")),
            "months_covered": safe_float(row.get("覆盖月份数")),
        }
        print(f"  ✓ 加载机具明细（机）: {f.name}")
        return detail
    except Exception as e:
        print(f"  加载机具明细失败 {project_name}: {e}")
        return {}


def attach_details_to_latest_projects(projects: List[ProjectLatest], master_df: pd.DataFrame, latest_month: str) -> Dict[str, Dict[str, Any]]:
    """为最新月份的项目批量附加 人/材/机 明细。返回 {项目名: details_dict}"""
    details: Dict[str, Dict[str, Any]] = {}
    proc_map = load_procurement_map()

    for p in projects:
        d: Dict[str, Any] = {
            "procurement_wan": proc_map.get(p.code, 0.0),
            "labor": load_labor_detail_for_project(p.code, p.name),
            "machinery": load_machine_detail_for_project(p.code, p.name),
        }
        # 也可以从 master 里取该项目最近几期的完整行，供历史 Tab 使用
        proj_hist = master_df[(master_df["项目名称"] == p.name)].sort_values("_month_key", ascending=False).head(6)
        d["history_rows"] = [
            {
                "month": str(r.get("月份", "")),
                "est": safe_float(r.get("预计总成本")),
                "margin": normalize_margin(r.get("毛利率")),
                "settle": safe_float(r.get("业主预计结算额")),
                "actual": safe_float(r.get("实际已发生金额")),
                "forecast": safe_float(r.get("后续预估汇总")),
                "labor_master": safe_float(r.get("人工费汇总")) or safe_float(r.get("工资支出")),
                "mach_master": safe_float(r.get("机具费汇总")) or safe_float(r.get("机具费支出")),
                "proc_master": safe_float(r.get("采购数据汇总")) or safe_float(r.get("采购部采购金额")),
            }
            for _, r in proj_hist.iterrows()
        ]
        details[p.name] = d
    return details


def make_html(stats: Dict[str, Any],
              projects_latest: List[ProjectLatest],
              trend: Dict[str, Any],
              category: Dict[str, Any],
              toc: List[str],
              cards_payload: Dict[str, Any],
              cards: List[Dict[str, Any]],
              source_file: str,
              project_details: Dict[str, Dict[str, Any]] = None) -> str:
    """生成多层级综合 HTML（干净占位符替换版，避免转义问题）。"""
    if project_details is None:
        project_details = {}

    # 序列化
    months_json = json.dumps(trend.get("labels", []), ensure_ascii=False)
    margin_series_json = json.dumps(trend.get("margin", {}), ensure_ascii=False)
    cost_series_json = json.dumps(trend.get("cost", {}), ensure_ascii=False)
    cat_labels_json = json.dumps(category.get("labels", []), ensure_ascii=False)
    cat_values_json = json.dumps(category.get("values", []), ensure_ascii=False)
    cards_json = json.dumps(cards_payload, ensure_ascii=False)
    projects_json = json.dumps([{
        "name": p.name, "code": p.code, "month": p.month,
        "est": round(p.est_cost, 2), "profit": round(p.gross_profit, 2),
        "margin": round(p.gross_margin * 100, 2) if p.gross_margin else None,
        "settle": round(p.settlement, 2), "finance": round(p.finance_paid, 2),
        "proc": round(p.procurement, 2), "mach": round(p.machinery, 2), "labor": round(p.labor, 2),
    } for p in projects_latest], ensure_ascii=False)
    details_json = json.dumps(project_details, ensure_ascii=False)

    # 计算决策支持字段
    def get_status(margin, variance=0, stale=False):
        if margin is None or margin < 0:
            return "danger", "负毛利/风险"
        if margin < 0.03 or abs(variance) > 500 or stale:
            return "warning", "需关注"
        return "good", "健康"

    project_cards = []
    risk_items = []
    for p in projects_latest:
        margin = p.gross_margin or 0
        # 简单方差示例 (计算总成本 vs 预计)
        variance = (p.est_cost or 0) - (p.actual_occurred or 0)  # 简化
        stale = False  # 可后续从 history 判断预估是否更新
        status, status_text = get_status(margin, variance, stale)

        payment_pct = 0
        if p.settlement and p.finance_paid:
            payment_pct = min(100, round(p.finance_paid / p.settlement * 100, 1))

        alerts = []
        if margin < 0: alerts.append("负毛利")
        if abs(variance) > 300: alerts.append("成本偏差大")
        if p.procurement and abs(p.procurement - (p.est_cost or 0)*0.3) > 200: alerts.append("采购异常")

        safe_name = html_escape.escape(p.name)
        card_html = f"""
<div class="proj-card {status}" data-project-name="{safe_name}">
  <div class="card-head">
    <div>
      <span class="code">{html_escape.escape(p.code or '-')}</span>
      <strong class="pname">{safe_name}</strong>
    </div>
    <span class="status {status}">{status_text}</span>
  </div>
  <div class="big-metric">
    <span class="metric-val {status}">{margin*100:.1f}%</span>
    <span class="metric-label">毛利率</span>
  </div>
  <div class="mini-stats">
    <div>预计 <strong>{p.est_cost:,.0f}</strong>万</div>
    <div>结算 <strong>{p.settlement:,.0f}</strong>万</div>
    <div>已付 <strong>{p.finance_paid:,.0f}</strong>万</div>
  </div>
  <div class="progress"><div class="bar" style="width:{payment_pct}%"></div></div>
  <div class="progress-label">付款进度 {payment_pct}%</div>
  <div class="alerts">{" ".join([f'<span class="alert">{a}</span>' for a in alerts]) or '<span class="no-alert">无明显异常</span>'}</div>
  <button class="detail-btn" data-project-name="{safe_name}">查看详情与人/材/机</button>
</div>"""
        project_cards.append(card_html)

        if status in ("danger", "warning"):
            risk_items.append({
                "name": p.name,
                "status": status_text,
                "margin": f"{margin*100:.1f}%",
                "alerts": ", ".join(alerts) or "成本结构需复核"
            })

    cards_grid_html = "\n".join(project_cards) if project_cards else "<p>无项目数据</p>"

    # 风险列表 HTML
    risk_html = ""
    if risk_items:
        risk_rows = "".join([f"<tr><td>{r['name']}</td><td><span class='status {r['status']}'>{r['status']}</span></td><td>{r['margin']}</td><td>{r['alerts']}</td></tr>" for r in risk_items])
        risk_html = f"""
<div class="risk-section">
  <h3>⚠️ 本期风险预警（需决策关注）</h3>
  <table class="risk-table">
    <thead><tr><th>项目</th><th>状态</th><th>毛利率</th><th>主要问题</th></tr></thead>
    <tbody>{risk_rows}</tbody>
  </table>
</div>"""

    # 表格行（保留兼容， data 属性）
    table_rows = []
    for p in projects_latest:
        margin_pct = f"{p.gross_margin*100:.1f}%" if p.gross_margin else "-"
        margin_class = "good" if (p.gross_margin or 0) >= 0.05 else ("warn" if (p.gross_margin or 0) > 0 else "bad")
        safe_name = html_escape.escape(p.name)
        table_rows.append(
            f"<tr data-project-name=\"{safe_name}\" data-margin=\"{p.gross_margin or 0}\" style=\"cursor:pointer\">"
            f"<td class='proj'><span class='code'>{html_escape.escape(p.code or '-')}</span><br><strong>{safe_name}</strong></td>"
            f"<td>{p.month}</td>"
            f"<td class='num'>{p.est_cost:,.1f}</td>"
            f"<td class='num'>{p.gross_profit:,.1f}</td>"
            f"<td class='num {margin_class}'>{margin_pct}</td>"
            f"<td class='num'>{p.settlement:,.1f}</td>"
            f"<td class='num'>{p.finance_paid:,.1f}</td>"
            f"<td class='num small'>{p.procurement:,.0f} / {p.machinery:,.0f} / {p.labor:,.0f}</td>"
            f"<td><button class='row-btn' data-project-name=\"{safe_name}\">进入详情 ▶</button></td>"
            f"</tr>"
        )
    table_html = "\n".join(table_rows) if table_rows else "<tr><td colspan='9'>暂无数据</td></tr>"

    # 卡片 sections（快速入口，使用 data 属性）
    card_sections = []
    for c in cards:
        safe_c = html_escape.escape(c['name'])
        card_sections.append(f"""
<section class="project-card" id="{c['id']}">
  <div class="card-header">
    <h3>{safe_c} <span class="code">({html_escape.escape(c['code'] or '-')})</span></h3>
    <span class="meta">{c['latest_month']}</span>
    <button class="export-btn" data-project-name="{safe_c}">进入多层级详情 ▶</button>
  </div>
  <div class="kpi-mini">
    <div>预计总成本 <strong>{c['est_cost']:,.1f}</strong> 万</div>
    <div>毛利率 <strong class="{'good' if (c['margin'] or 0)>=0.05 else 'bad'}">{(c['margin']*100) if c['margin'] else '-':.1f}%</strong></div>
    <div>业主结算 <strong>{c['settlement']:,.1f}</strong> 万</div>
    <div>分包财务付 <strong>{c['finance_paid']:,.1f}</strong> 万</div>
  </div>
  <div class="chart-grid">
    <div class="chart-wrap"><canvas id="c_{c['id']}_m"></canvas></div>
    <div class="chart-wrap"><canvas id="c_{c['id']}_a"></canvas></div>
  </div>
  <table class="mini"><thead><tr><th>月份</th><th>预计成本</th><th>毛利率</th><th>结算额</th></tr></thead><tbody>{c['rows_html']}</tbody></table>
</section>
""")
    cards_html = "\n".join(card_sections)

    generated = stats["generated_at"]
    src_name = Path(source_file).name

    # 完整 HTML 模板（普通字符串 + 清晰占位符）
    html_content = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>项目成本统计综合分析报告（多层级）</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root { --bg:#f6f8f5; --card:#fff; --line:#e2e8e3; --txt:#1f2a24; --muted:#5c6b63; --accent:#2b8a6e; --good:#2b8a6e; --warn:#c48a2f; --bad:#c44a3f; }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--txt); font-family:"PingFang SC","Noto Sans SC","Microsoft YaHei",system-ui,sans-serif; }
.container { max-width:1320px; margin:0 auto; padding:16px 20px 40px; }
.hero { background:linear-gradient(135deg,#2b8a6e 0%,#1f6348 100%); color:#fff; padding:22px 26px; border-radius:16px; margin-bottom:16px; box-shadow:0 10px 30px rgba(43,138,110,.25); }
h1 { margin:0 0 4px; font-size:24px; }
.subtitle { opacity:.92; margin:0; font-size:13px; }
.metrics { display:grid; grid-template-columns:repeat(auto-fit,minmax(148px,1fr)); gap:10px; margin:12px 0; }
.metric { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px 14px; }
.metric .val { font-size:20px; font-weight:700; }
.metric .lbl { font-size:11px; color:var(--muted); }
.section { background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px; margin:14px 0; box-shadow:0 3px 10px rgba(0,0,0,.03); }}
h2 { margin:0 0 10px; font-size:17px; border-left:4px solid var(--accent); padding-left:8px; }
table { width:100%; border-collapse:collapse; font-size:12.5px; }
th,td { padding:7px 6px; border-bottom:1px solid var(--line); text-align:left; }
th { background:#f1f5f2; color:#31453a; font-weight:600; position:sticky; top:0; }
.num { text-align:right; font-variant-numeric:tabular-nums; }
.good { color:var(--good); font-weight:600; } .warn { color:var(--warn); font-weight:600; } .bad { color:var(--bad); font-weight:600; }
.code { font-size:10px; color:#666; background:#f4f7f5; padding:0 4px; border-radius:2px; }
.row-btn { font-size:11px; padding:2px 8px; border:none; background:#2b8a6e; color:#fff; border-radius:4px; cursor:pointer; }
#detail-panel { border:2px solid #2b8a6e; border-radius:14px; padding:14px; margin:16px 0; background:#fafcf9; display:none; }
#detail-panel.open { display:block; }
.detail-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:8px; }
.detail-header h3 { margin:0; font-size:18px; }
.tabbar { display:flex; flex-wrap:wrap; gap:4px; margin:8px 0; border-bottom:1px solid var(--line); padding-bottom:4px; }
.tab-btn { padding:5px 12px; border:1px solid var(--line); background:#fff; border-radius:6px 6px 0 0; cursor:pointer; font-size:12.5px; }
.tab-btn.active { background:#2b8a6e; color:#fff; border-color:#2b8a6e; }
.tab-content { min-height:220px; }
@media print { .row-btn, .tab-btn { display:none !important; } #detail-panel { border:1px solid #ccc; } }
</style>
</head>
<body>
<div class="container">
  <div class="hero">
    <h1>📊 项目成本统计综合分析报告 <span style="font-size:13px;opacity:.85;">（多层级钻取版）</span></h1>
    <p class="subtitle">数据来源：{{SRC_NAME}} | 最新月份：{{LATEST_MONTH}} | 生成：{{GENERATED}}</p>
    <p style="margin:4px 0 0;opacity:.9;font-size:12.5px;">点击表格任意行或「进入详情」→ 进入项目多层级面板，可查看概况 + 具体的人(人工明细)/材(采购)/机(机具)等数据。</p>
  </div>

  <!-- 全局 KPI -->
  <div class="metrics">
    <div class="metric"><div class="val">{{PROJECT_COUNT}}</div><div class="lbl">项目数</div></div>
    <div class="metric"><div class="val">{{TOTAL_EST}}</div><div class="lbl">预计总成本(万)</div></div>
    <div class="metric"><div class="val">{{TOTAL_SETTLE}}</div><div class="lbl">业主结算额(万)</div></div>
    <div class="metric"><div class="val">{{AVG_MARGIN}}</div><div class="lbl">平均毛利率</div></div>
    <div class="metric"><div class="val">{{TOTAL_FINANCE}}</div><div class="lbl">财务分包已付(万)</div></div>
    <div class="metric"><div class="val">{{TOTAL_PROC}} / {{TOTAL_MACH}}</div><div class="lbl">采购 / 机具(万)</div></div>
  </div>

  <!-- 概览表 -->
  <div class="section">
    <h2>🔍 项目总览表（点击行进入详情）</h2>
    <div class="controls" style="margin-bottom:8px;">
      <input id="search" type="text" placeholder="搜索项目..." oninput="filterTable()" style="flex:1;min-width:200px;">
      <button onclick="resetFilters()" style="padding:5px 10px;border:1px solid var(--line);background:#fff;border-radius:6px;font-size:12px;">重置</button>
      <button onclick="window.print()" style="padding:5px 10px;background:#2b8a6e;color:#fff;border:none;border-radius:6px;font-size:12px;">打印/PDF</button>
    </div>
    <div style="max-height:380px;overflow:auto;">
      <table id="mainTable">
        <thead><tr>
          <th>项目</th><th>月份</th><th>预计总成本</th><th>毛利</th><th>毛利率</th>
          <th>业主结算</th><th>财务已付</th><th>采购/机具/人工</th><th></th>
        </tr></thead>
        <tbody>{{TABLE_HTML}}</tbody>
      </table>
    </div>
    <div style="font-size:11px;color:var(--muted);margin-top:4px;">提示：表格行可直接点击进入该项目的「人/材/机」多层级详情面板。</div>
  </div>

  <!-- 趋势 & 结构 -->
  <div class="section">
    <h2>📈 趋势 & 成本结构（概览层）</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      <div><canvas id="trendMargin" style="height:240px;"></canvas></div>
      <div><canvas id="trendCost" style="height:240px;"></canvas></div>
    </div>
    <div style="margin-top:12px;display:flex;gap:16px;align-items:center;flex-wrap:wrap;">
      <div style="width:260px;height:200px;"><canvas id="catPie"></canvas></div>
      <div style="font-size:12px;color:#33443b;max-width:420px;">
        成本结构基于最新 master 聚合。 更细的人/材/机明细请在下方「项目详情面板」里选择具体项目查看。
      </div>
    </div>
  </div>

  <!-- 快速卡片入口 -->
  <div class="section">
    <h2>🧩 快速项目卡片（点击按钮进入多层级人/材/机详情）</h2>
    {{CARDS_HTML}}
  </div>

  <!-- 多层级详情面板 -->
  <div id="detail-panel">
    <div class="detail-header">
      <h3 id="detail-title">项目深度分析</h3>
      <div>
        <button onclick="closeDetailPanel()" style="padding:4px 10px;border:1px solid #ccc;background:#fff;border-radius:4px;font-size:12px;">关闭面板</button>
        <button onclick="window.print()" style="padding:4px 10px;background:#2b8a6e;color:#fff;border:none;border-radius:4px;font-size:12px;">打印当前</button>
      </div>
    </div>

    <div style="margin-bottom:8px;font-size:12px;">
      快速切换：<select id="proj-selector" onchange="selectProject(this.value)" style="padding:3px 6px;"></select>
    </div>

    <div class="tabbar">
      <button class="tab-btn active" data-tab="overview" onclick="switchTab('overview')">概况 & 预估</button>
      <button class="tab-btn" data-tab="ren" onclick="switchTab('ren')">👷 人工 (人)</button>
      <button class="tab-btn" data-tab="cai" onclick="switchTab('cai')">📦 采购材料 (材)</button>
      <button class="tab-btn" data-tab="ji" onclick="switchTab('ji')">🚜 机具设备 (机)</button>
      <button class="tab-btn" data-tab="history" onclick="switchTab('history')">📜 历史 & 合同财务</button>
    </div>

    <div id="tab-content" class="tab-content"></div>

    <div style="font-size:11px;color:#666;margin-top:8px;">
      数据来自 master + 专业 per-project 明细（人工按项目拆分、机具汇总、采购 D_sum 等）。Tab 为空表示暂无对应明细文件。
    </div>
  </div>

  <div class="section" style="background:#fffef5;border-color:#f0e6c8;">
    <h2>⚠️ 核算思路简要 & 数据说明</h2>
    <div style="font-size:12px;color:#534a2f;line-height:1.6;">
      本报告是「项目自己填的成本统计模板刮取值」+「各专业台账客观汇总」的对账视图。<br>
      点击项目后可在「人工(人)」Tab 看到中方正式工/劳务/属地工等具体构成（来自按项目拆分人工费文件）；「采购(材)」和「机具(机)」同理。
    </div>
  </div>

  <div style="text-align:center;margin:24px 0 8px;color:#8a968f;font-size:11px;">
    泰国子公司成本统计系统 · 多层级综合分析报告
  </div>
</div>

<script>
const months = __MONTHS_JSON__;
const marginData = __MARGIN_JSON__;
const costData = __COST_JSON__;
const catLabels = __CAT_LABELS__;
const catValues = __CAT_VALUES__;
const cardsData = __CARDS_JSON__;
const projectsData = __PROJECTS_JSON__;
const allDetails = __DETAILS_JSON__;

let currentProj = null;
let detailCharts = [];

function filterTable() {
  const q = (document.getElementById('search').value || '').toLowerCase().trim();
  const minM = parseFloat(document.getElementById('minMargin').value || '-1');
  document.querySelectorAll('#mainTable tbody tr').forEach(r => {
    const name = (r.getAttribute('data-name') || '').toLowerCase();
    const m = parseFloat(r.getAttribute('data-margin') || '0');
    r.style.display = (!q || name.includes(q)) && (m >= minM/100) ? '' : 'none';
  });
}
function resetFilters() { document.getElementById('search').value=''; document.getElementById('minMargin').value='-1'; filterTable(); }

function populateProjectSelector() {
  const sel = document.getElementById('proj-selector');
  if (!sel) return;
  sel.innerHTML = '<option value="">-- 选择项目 --</option>';
  projectsData.forEach(p => {
    const o = document.createElement('option'); o.value = p.name; o.textContent = p.name + ' (' + (p.code||'') + ')'; sel.appendChild(o);
  });
}

function selectProject(name) {
  if (!name) return;
  currentProj = name;
  const panel = document.getElementById('detail-panel');
  panel.classList.add('open');
  panel.scrollIntoView({behavior:'smooth', block:'start'});
  const sel = document.getElementById('proj-selector'); if (sel) sel.value = name;
  switchTab('overview');
}

function closeDetailPanel() {
  document.getElementById('detail-panel').classList.remove('open');
  currentProj = null;
}

function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  renderTabContent(tab);
}

function destroyDetailCharts() { detailCharts.forEach(c => { try{c.destroy();}catch(e){} }); detailCharts = []; }

function renderTabContent(tab) {
  const container = document.getElementById('tab-content');
  if (!container || !currentProj) return;
  destroyDetailCharts();

  const proj = projectsData.find(x => x.name === currentProj) || {};
  const det = allDetails[currentProj] || {};
  const labor = det.labor || {};
  const mach = det.machinery || {};
  const hist = det.history_rows || [];

  let html = '';
  if (tab === 'overview') {
    html = `<div class="kpi-strip">
      <div><strong>预计总成本</strong> ${proj.est||0} 万</div>
      <div><strong>毛利</strong> ${proj.profit||0} 万 &nbsp; 毛利率 <span class="${(proj.margin||0)>=5?'good':'bad'}">${proj.margin||'-'}%</span></div>
      <div><strong>业主结算</strong> ${proj.settle||0} 万</div>
      <div><strong>财务分包已付</strong> ${proj.finance||0} 万</div>
    </div>
    <div style="font-size:12.5px;margin-top:6px;"><strong>核对提示</strong>：master 最新记录。 具体构成见其他 Tab。</div>`;
  } else if (tab === 'ren') {
    const lmonths = labor.months || [];
    const br = labor.breakdown || {};
    html = `<div><strong>人工费明细（人）</strong> — ${labor.file || '未找到按项目拆分文件'}</div>`;
    if (lmonths.length) {
      html += `<table class="mini"><thead><tr><th>年-月</th>${Object.keys(br).map(k=>`<th>${k}</th>`).join('')}</tr></thead><tbody>`;
      for (let i=0; i<lmonths.length; i++) {
        html += `<tr><td>${lmonths[i]}</td>${Object.keys(br).map(k => `<td class="num">${(br[k][i]||0).toLocaleString()}</td>`).join('')}</tr>`;
      }
      html += `</tbody></table><div style="margin-top:6px;height:150px;"><canvas id="laborChart"></canvas></div>`;
    } else {
      html += `<div style="color:#c44a3f;">暂无人工 per-project 明细。master 人工费汇总：${proj.labor||0}。</div>`;
    }
  } else if (tab === 'cai') {
    html = `<div><strong>采购材料（材）</strong></div>
      <div class="kpi-strip">
        <div>master 采购数据汇总：<strong>${proj.proc||0}</strong> 万</div>
        <div>采购台账总采购金额：<strong>${det.procurement_wan||0}</strong> 万</div>
      </div>
      <div style="font-size:12px;color:#555;">更细子目见原始模板「2材料汇总」或采购标准化台账。</div>`;
  } else if (tab === 'ji') {
    html = `<div><strong>机具设备（机）</strong> — ${mach.file || '未找到机具 per-project 汇总'}</div>
      <div class="kpi-strip">
        <div>费用合计(泰铢)：<strong>${(mach.total_thb||0).toLocaleString()}</strong></div>
        <div>使用次数：${mach.use_count||'-'} &nbsp; 机具种类：${mach.kind_count||'-'}</div>
      </div>
      <div style="font-size:12px;color:#555;">详细记录在 10_机具费 的处理报告中。master 机具费汇总：${proj.mach||0}。</div>`;
  } else if (tab === 'history') {
    html = `<div><strong>历史轨迹（来自 master）</strong></div>
      <table class="mini"><thead><tr><th>月份</th><th>预计成本</th><th>毛利率</th><th>结算额</th><th>人工(master)</th><th>机具(master)</th></tr></thead><tbody>
      ${hist.map(h => `<tr><td>${h.month}</td><td class="num">${h.est}</td><td class="num">${h.margin? (h.margin*100).toFixed(1)+'%':'-'}</td><td class="num">${h.settle}</td><td class="num">${h.labor_master||'-'}</td><td class="num">${h.mach_master||'-'}</td></tr>`).join('')}
      </tbody></table>`;
  }

  container.innerHTML = html;

  if (tab === 'ren' && labor.months && labor.months.length) {
    const el = document.getElementById('laborChart');
    if (el) {
      const br = labor.breakdown || {};
      const ds = Object.keys(br).map((k,i) => ({label:k, data:br[k], borderWidth:1.5, tension:.2}));
      const c = new Chart(el, {type:'line', data:{labels:labor.months, datasets:ds}, options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9}}}}}} });
      detailCharts.push(c);
    }
  }
}

function initOverviewCharts() {
  const mEl = document.getElementById('trendMargin');
  const cEl = document.getElementById('trendCost');
  if (mEl && marginData) {
    const labels = months;
    const mds = Object.keys(marginData).map((k,i)=>({label:k,data:marginData[k],borderColor:['#2b8a6e','#3f6fcf','#c48a2f'][i%3],borderWidth:2,tension:.25,spanGaps:true}));
    new Chart(mEl,{type:'line',data:{labels,datasets:mds},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>(v*100).toFixed(0)+'%'}}}}});
  }
  if (cEl && costData) {
    const labels = months;
    const cds = Object.keys(costData).map((k,i)=>({label:k,data:costData[k],borderColor:['#2b8a6e','#3f6fcf'][i%2],borderWidth:2,tension:.2,spanGaps:true}));
    new Chart(cEl,{type:'line',data:{labels,datasets:cds},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}}}});
  }
  const pie = document.getElementById('catPie');
  if (pie && catLabels.length) {
    new Chart(pie,{type:'doughnut',data:{labels:catLabels,datasets:[{data:catValues,backgroundColor:['#2b8a6e','#3f6fcf','#c48a2f','#8e4db4']}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'}}}});
  }
}

function initAll() {
  populateProjectSelector();
  initOverviewCharts();

  // 附加点击事件（使用 data-project-name，更可靠，无 inline onclick 转义问题）
  document.querySelectorAll('#mainTable tbody tr').forEach(tr => {
    tr.addEventListener('click', function(e) {
      const name = tr.getAttribute('data-project-name');
      if (name) selectProject(name);
    });
  });

  document.querySelectorAll('.project-card .export-btn, #mainTable .row-btn').forEach(btn => {
    btn.addEventListener('click', function(e) {
      e.stopImmediatePropagation();
      const name = btn.getAttribute('data-project-name');
      if (name) selectProject(name);
    });
  });

  // 初始过滤一次
  setTimeout(() => { const s=document.getElementById('search'); if(s) s.oninput=filterTable; }, 20);

  console.log('%c[多层级成本分析] 报告已就绪 - 点击表格行或按钮可打开项目详情面板', 'color:#2b8a6e');
}
window.addEventListener('load', initAll);
</script>
</body>
</html>
"""

    # 占位符替换
    html_content = (html_content
        .replace("{{SRC_NAME}}", html_escape.escape(src_name))
        .replace("{{LATEST_MONTH}}", stats.get("latest_month", ""))
        .replace("{{GENERATED}}", generated)
        .replace("{{PROJECT_COUNT}}", str(stats.get("project_count", 0)))
        .replace("{{TOTAL_EST}}", f"{stats.get('total_est_cost_wan', 0):,.0f}")
        .replace("{{TOTAL_SETTLE}}", f"{stats.get('total_settlement_wan', 0):,.0f}")
        .replace("{{AVG_MARGIN}}", f"{(stats.get('avg_margin') or 0)*100:.1f}%")
        .replace("{{TOTAL_FINANCE}}", f"{stats.get('total_finance_paid_wan', 0):,.0f}")
        .replace("{{TOTAL_PROC}}", f"{stats.get('total_procurement_wan', 0):,.0f}")
        .replace("{{TOTAL_MACH}}", f"{stats.get('total_machinery_wan', 0):,.0f}")
        .replace("{{TABLE_HTML}}", table_html or "")
        .replace("{{CARDS_HTML}}", cards_html or "")
        .replace("__MONTHS_JSON__", months_json)
        .replace("__MARGIN_JSON__", margin_series_json)
        .replace("__COST_JSON__", cost_series_json)
        .replace("__CAT_LABELS__", cat_labels_json)
        .replace("__CAT_VALUES__", cat_values_json)
        .replace("__CARDS_JSON__", cards_json)
        .replace("__PROJECTS_JSON__", projects_json)
        .replace("__DETAILS_JSON__", details_json)
    )

    return html_content
    """生成多层级综合 HTML（概览 + 可选择项目的 人/材/机 钻取详情面板）"""
    if project_details is None:
        project_details = {}

    # 序列化数据
    months_json = json.dumps(trend.get("labels", []), ensure_ascii=False)
    margin_series_json = json.dumps(trend.get("margin", {}), ensure_ascii=False)
    cost_series_json = json.dumps(trend.get("cost", {}), ensure_ascii=False)
    cat_labels_json = json.dumps(category.get("labels", []), ensure_ascii=False)
    cat_values_json = json.dumps(category.get("values", []), ensure_ascii=False)
    cards_json = json.dumps(cards_payload, ensure_ascii=False)
    projects_json = json.dumps([
        {
            "name": p.name,
            "code": p.code,
            "month": p.month,
            "est": round(p.est_cost, 2),
            "profit": round(p.gross_profit, 2),
            "margin": round(p.gross_margin * 100, 2) if p.gross_margin else None,
            "settle": round(p.settlement, 2),
            "finance": round(p.finance_paid, 2),
            "proc": round(p.procurement, 2),
            "mach": round(p.machinery, 2),
            "labor": round(p.labor, 2),
        } for p in projects_latest
    ], ensure_ascii=False)
    details_json = json.dumps(project_details, ensure_ascii=False)

    # 表格行（最新）—— 点击即可进入下一层（多层级核心）
    table_rows = []
    for p in projects_latest:
        margin_pct = f"{p.gross_margin*100:.1f}%" if p.gross_margin else "-"
        margin_class = "good" if (p.gross_margin or 0) >= 0.05 else ("warn" if (p.gross_margin or 0) > 0 else "bad")
        safe_name = html_escape.escape(p.name)
        table_rows.append(
            f"<tr data-name='{safe_name}' data-margin='{p.gross_margin or 0}' onclick=\"selectProject('{safe_name}')\" style='cursor:pointer'>"
            f"<td class='proj'><span class='code'>{html_escape.escape(p.code or '-')}</span><br><strong>{safe_name}</strong></td>"
            f"<td>{p.month}</td>"
            f"<td class='num'>{p.est_cost:,.1f}</td>"
            f"<td class='num'>{p.gross_profit:,.1f}</td>"
            f"<td class='num {margin_class}'>{margin_pct}</td>"
            f"<td class='num'>{p.settlement:,.1f}</td>"
            f"<td class='num'>{p.finance_paid:,.1f}</td>"
            f"<td class='num small'>{p.procurement:,.0f} / {p.machinery:,.0f} / {p.labor:,.0f}</td>"
            f"<td><button class='row-btn' onclick=\"event.stopImmediatePropagation(); selectProject('{safe_name}')\">进入详情 ▶</button></td>"
            f"</tr>"
        )
    table_html = "\n".join(table_rows) if table_rows else "<tr><td colspan='9'>暂无数据</td></tr>"

    # 保留旧卡片（快速入口）
    card_sections = []
    for c in cards:
        card_sections.append(f"""
<section class="project-card" id="{c['id']}">
  <div class="card-header">
    <h3>{html_escape.escape(c['name'])} <span class="code">({html_escape.escape(c['code'] or '-')})</span></h3>
    <span class="meta">{c['latest_month']}</span>
    <button class="export-btn" onclick="event.stopImmediatePropagation(); selectProject('{html_escape.escape(c['name'])}')">进入多层级详情 ▶</button>
  </div>
  <div class="kpi-mini">
    <div>预计总成本 <strong>{c['est_cost']:,.1f}</strong> 万</div>
    <div>毛利率 <strong class="{'good' if (c['margin'] or 0)>=0.05 else 'bad'}">{(c['margin']*100) if c['margin'] else '-':.1f}%</strong></div>
    <div>业主结算 <strong>{c['settlement']:,.1f}</strong> 万</div>
    <div>分包财务付 <strong>{c['finance_paid']:,.1f}</strong> 万</div>
  </div>
  <div class="chart-grid">
    <div class="chart-wrap"><canvas id="c_{c['id']}_m"></canvas></div>
    <div class="chart-wrap"><canvas id="c_{c['id']}_a"></canvas></div>
  </div>
  <table class="mini"><thead><tr><th>月份</th><th>预计成本</th><th>毛利率</th><th>结算额</th></tr></thead><tbody>{c['rows_html']}</tbody></table>
</section>
""")
    cards_html = "\n".join(card_sections)

    generated = stats["generated_at"]
    src_name = Path(source_file).name

    # ========== 多层级 HTML（概览 + 项目选择后的人/材/机 Tab 面板） ==========
    # Use regular string (not f) to avoid JS { } being interpreted as f-string expressions
    html_content = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>项目成本统计综合分析报告（多层级）</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root {{ --bg:#f6f8f5; --card:#fff; --line:#e2e8e3; --txt:#1f2a24; --muted:#5c6b63; --accent:#2b8a6e; --good:#2b8a6e; --warn:#c48a2f; --bad:#c44a3f; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--txt); font-family:"PingFang SC","Noto Sans SC","Microsoft YaHei",system-ui,sans-serif; }}
.container {{ max-width:1320px; margin:0 auto; padding:16px 20px 40px; }}
.hero {{ background:linear-gradient(135deg,#2b8a6e 0%,#1f6348 100%); color:#fff; padding:22px 26px; border-radius:16px; margin-bottom:16px; box-shadow:0 10px 30px rgba(43,138,110,.25); }}
h1 {{ margin:0 0 4px; font-size:24px; }}
.subtitle {{ opacity:.92; margin:0; font-size:13px; }}
.metrics {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(148px,1fr)); gap:10px; margin:12px 0; }}
.metric {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px 14px; }}
.metric .val {{ font-size:20px; font-weight:700; }}
.metric .lbl {{ font-size:11px; color:var(--muted); }}
.section {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px; margin:14px 0; box-shadow:0 3px 10px rgba(0,0,0,.03); }}
h2 {{ margin:0 0 10px; font-size:17px; border-left:4px solid var(--accent); padding-left:8px; }}
table {{ width:100%; border-collapse:collapse; font-size:12.5px; }}
th,td {{ padding:7px 6px; border-bottom:1px solid var(--line); text-align:left; }}
th {{ background:#f1f5f2; color:#31453a; font-weight:600; position:sticky; top:0; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
.good {{ color:var(--good); font-weight:600; }} .warn {{ color:var(--warn); font-weight:600; }} .bad {{ color:var(--bad); font-weight:600; }}
.code {{ font-size:10px; color:#666; background:#f4f7f5; padding:0 4px; border-radius:2px; }}
.row-btn {{ font-size:11px; padding:2px 8px; border:none; background:#2b8a6e; color:#fff; border-radius:4px; cursor:pointer; }}
#detail-panel {{ border:2px solid #2b8a6e; border-radius:14px; padding:14px; margin:16px 0; background:#fafcf9; display:none; }}
#detail-panel.open {{ display:block; }}
.detail-header {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:8px; }}
.detail-header h3 {{ margin:0; font-size:18px; }}
.tabbar {{ display:flex; flex-wrap:wrap; gap:4px; margin:8px 0; border-bottom:1px solid var(--line); padding-bottom:4px; }}
.tab-btn {{ padding:5px 12px; border:1px solid var(--line); background:#fff; border-radius:6px 6px 0 0; cursor:pointer; font-size:12.5px; }}
.tab-btn.active {{ background:#2b8a6e; color:#fff; border-color:#2b8a6e; }}
.tab-content {{ min-height:220px; }}
@media print {{ .row-btn, .tab-btn {{ display:none !important; }} #detail-panel {{ border:1px solid #ccc; }} }}
</style>
</head>
<body>
<div class="container">
  <div class="hero">
    <h1>📊 项目成本统计综合分析报告 <span style="font-size:13px;opacity:.85;">（多层级钻取版）</span></h1>
    <p class="subtitle">数据来源：{html_escape.escape(src_name)} | 最新月份：{stats['latest_month']} | 生成：{generated}</p>
    <p style="margin:4px 0 0;opacity:.9;font-size:12.5px;">点击表格任意行或「进入详情」→ 进入项目多层级面板，可查看概况 + 具体的人(人工明细)/材(采购)/机(机具)等数据。</p>
  </div>

  <!-- 全局 KPI -->
  <div class="metrics">
    <div class="metric"><div class="val">{stats['project_count']}</div><div class="lbl">项目数</div></div>
    <div class="metric"><div class="val">{stats['total_est_cost_wan']:,.0f}</div><div class="lbl">预计总成本(万)</div></div>
    <div class="metric"><div class="val">{stats['total_settlement_wan']:,.0f}</div><div class="lbl">业主结算额(万)</div></div>
    <div class="metric"><div class="val">{stats['avg_margin']*100:.1f}%</div><div class="lbl">平均毛利率</div></div>
    <div class="metric"><div class="val">{stats['total_finance_paid_wan']:,.0f}</div><div class="lbl">财务分包已付(万)</div></div>
    <div class="metric"><div class="val">{stats['total_procurement_wan']:,.0f} / {stats['total_machinery_wan']:,.0f}</div><div class="lbl">采购 / 机具(万)</div></div>
  </div>

  <!-- 概览表（可点击进入下一层） -->
  <div class="section">
    <h2>🔍 项目总览表（点击行进入详情）</h2>
    <div class="controls" style="margin-bottom:8px;">
      <input id="search" type="text" placeholder="搜索项目..." oninput="filterTable()" style="flex:1;min-width:200px;">
      <button onclick="resetFilters()" style="padding:5px 10px;border:1px solid var(--line);background:#fff;border-radius:6px;font-size:12px;">重置</button>
      <button onclick="window.print()" style="padding:5px 10px;background:#2b8a6e;color:#fff;border:none;border-radius:6px;font-size:12px;">打印/PDF</button>
    </div>
    <div style="max-height:380px;overflow:auto;">
      <table id="mainTable">
        <thead><tr>
          <th>项目</th><th>月份</th><th>预计总成本</th><th>毛利</th><th>毛利率</th>
          <th>业主结算</th><th>财务已付</th><th>采购/机具/人工</th><th></th>
        </tr></thead>
        <tbody>{table_html}</tbody>
      </table>
    </div>
    <div style="font-size:11px;color:var(--muted);margin-top:4px;">提示：表格行可直接点击进入该项目的「人/材/机」多层级详情面板。</div>
  </div>

  <!-- 趋势 & 结构（保留概览层） -->
  <div class="section">
    <h2>📈 趋势 & 成本结构（概览层）</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      <div><canvas id="trendMargin" style="height:240px;"></canvas></div>
      <div><canvas id="trendCost" style="height:240px;"></canvas></div>
    </div>
    <div style="margin-top:12px;display:flex;gap:16px;align-items:center;flex-wrap:wrap;">
      <div style="width:260px;height:200px;"><canvas id="catPie"></canvas></div>
      <div style="font-size:12px;color:#33443b;max-width:420px;">
        成本结构基于最新 master 聚合。 更细的人/材/机明细请在下方「项目详情面板」里选择具体项目查看。
      </div>
    </div>
  </div>

  <!-- 旧卡片快速入口 -->
  <div class="section">
    <h2>🧩 快速项目卡片（点击按钮进入多层级人/材/机详情）</h2>
    {cards_html}
  </div>

  <!-- ==================== 多层级详情面板（核心） ==================== -->
  <div id="detail-panel">
    <div class="detail-header">
      <h3 id="detail-title">项目深度分析</h3>
      <div>
        <button onclick="closeDetailPanel()" style="padding:4px 10px;border:1px solid #ccc;background:#fff;border-radius:4px;font-size:12px;">关闭面板</button>
        <button onclick="window.print()" style="padding:4px 10px;background:#2b8a6e;color:#fff;border:none;border-radius:4px;font-size:12px;">打印当前</button>
      </div>
    </div>

    <div style="margin-bottom:8px;font-size:12px;">
      快速切换：<select id="proj-selector" onchange="selectProject(this.value)" style="padding:3px 6px;"></select>
    </div>

    <div class="tabbar">
      <button class="tab-btn active" data-tab="overview" onclick="switchTab('overview')">概况 & 预估</button>
      <button class="tab-btn" data-tab="ren" onclick="switchTab('ren')">👷 人工 (人)</button>
      <button class="tab-btn" data-tab="cai" onclick="switchTab('cai')">📦 采购材料 (材)</button>
      <button class="tab-btn" data-tab="ji" onclick="switchTab('ji')">🚜 机具设备 (机)</button>
      <button class="tab-btn" data-tab="history" onclick="switchTab('history')">📜 历史 & 合同财务</button>
    </div>

    <div id="tab-content" class="tab-content"></div>

    <div style="font-size:11px;color:#666;margin-top:8px;">
      数据来自 master + 专业 per-project 明细（人工按项目拆分、机具汇总、采购 D_sum 等）。Tab 为空表示暂无对应明细文件。
    </div>
  </div>

  <div class="section" style="background:#fffef5;border-color:#f0e6c8;">
    <h2>⚠️ 核算思路简要 & 数据说明</h2>
    <div style="font-size:12px;color:#534a2f;line-height:1.6;">
      本报告是「项目自己填的成本统计模板刮取值」+「各专业台账客观汇总」的对账视图。<br>
      点击项目后可在「人工(人)」Tab 看到中方正式工/劳务/属地工等具体构成（来自按项目拆分人工费文件）；「采购(材)」和「机具(机)」同理。<br>
      更完整的明细仍建议参考各 B_xxx_sum 目录下的专业报告。
    </div>
  </div>

  <div style="text-align:center;margin:24px 0 8px;color:#8a968f;font-size:11px;">
    泰国子公司成本统计系统 · 多层级综合分析报告 · 基于现有成本收集核算流程
  </div>
</div>

<script>
const months = {months_json};
const marginData = {margin_series_json};
const costData = {cost_series_json};
const catLabels = {cat_labels_json};
const catValues = {cat_values_json};
const cardsData = {cards_json};
const projectsData = {projects_json};
const allDetails = {details_json};

let currentProj = null;
let detailCharts = [];

function filterTable() {{
  const q = (document.getElementById('search').value || '').toLowerCase().trim();
  const minM = parseFloat(document.getElementById('minMargin').value || '-1');
  document.querySelectorAll('#mainTable tbody tr').forEach(r => {{
    const name = (r.getAttribute('data-name') || '').toLowerCase();
    const m = parseFloat(r.getAttribute('data-margin') || '0');
    r.style.display = (!q || name.includes(q)) && (m >= minM/100) ? '' : 'none';
  }});
}}
function resetFilters() {{ document.getElementById('search').value=''; document.getElementById('minMargin').value='-1'; filterTable(); }}

function populateProjectSelector() {{
  const sel = document.getElementById('proj-selector');
  if (!sel) return;
  sel.innerHTML = '<option value="">-- 选择项目 --</option>';
  projectsData.forEach(p => {{
    const o = document.createElement('option'); o.value = p.name; o.textContent = p.name + ' (' + (p.code||'') + ')'; sel.appendChild(o);
  }});
}}

function selectProject(name) {{
  if (!name) return;
  currentProj = name;
  const panel = document.getElementById('detail-panel');
  panel.classList.add('open');
  panel.scrollIntoView({{behavior:'smooth', block:'start'}});
  const sel = document.getElementById('proj-selector'); if (sel) sel.value = name;
  switchTab('overview');
}}

function closeDetailPanel() {{
  document.getElementById('detail-panel').classList.remove('open');
  currentProj = null;
}}

function switchTab(tab) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  renderTabContent(tab);
}}

function destroyDetailCharts() {{ detailCharts.forEach(c => {{ try{{c.destroy();}}catch(e){{}} }}); detailCharts = []; }}

function renderTabContent(tab) {{
  const container = document.getElementById('tab-content');
  if (!container || !currentProj) return;
  destroyDetailCharts();

  const proj = projectsData.find(x => x.name === currentProj) || {{}};
  const det = allDetails[currentProj] || {{}};
  const labor = det.labor || {{}};
  const mach = det.machinery || {{}};
  const hist = det.history_rows || [];

  let html = '';
  if (tab === 'overview') {{
    html = `<div class="kpi-strip">
      <div><strong>预计总成本</strong> ${{proj.est||0}} 万</div>
      <div><strong>毛利</strong> ${{proj.profit||0}} 万 &nbsp; 毛利率 <span class="${{(proj.margin||0)>=5?'good':'bad'}}">${{proj.margin||'-'}}%</span></div>
      <div><strong>业主结算</strong> ${{proj.settle||0}} 万</div>
      <div><strong>财务分包已付</strong> ${{proj.finance||0}} 万</div>
    </div>
    <div style="font-size:12.5px;margin-top:6px;"><strong>核对提示</strong>：master 最新记录（模板刮取 + 台账 join）。具体构成见其他 Tab。</div>`;
  }} else if (tab === 'ren') {{
    const lmonths = labor.months || [];
    const br = labor.breakdown || {{}};
    html = `<div><strong>人工费明细（人）</strong> — ${{labor.file || '未找到按项目拆分文件'}}</div>`;
    if (lmonths.length) {{
      html += `<table class="mini"><thead><tr><th>年-月</th>${{Object.keys(br).map(k=>`<th>${{k}}</th>`).join('')}}</tr></thead><tbody>`;
      for (let i=0; i<lmonths.length; i++) {{
        html += `<tr><td>${{lmonths[i]}}</td>${{Object.keys(br).map(k => `<td class="num">${{(br[k][i]||0).toLocaleString()}}</td>`).join('')}}</tr>`;
      }}
      html += `</tbody></table><div style="margin-top:6px;height:150px;"><canvas id="laborChart"></canvas></div>`;
    }} else {{
      html += `<div style="color:#c44a3f;">暂无人工 per-project 明细。master 人工费汇总：${{proj.labor||0}}。</div>`;
    }}
  }} else if (tab === 'cai') {{
    html = `<div><strong>采购材料（材）</strong></div>
      <div class="kpi-strip">
        <div>master 采购数据汇总：<strong>${{proj.proc||0}}</strong> 万</div>
        <div>采购台账总采购金额：<strong>${{det.procurement_wan||0}}</strong> 万</div>
      </div>
      <div style="font-size:12px;color:#555;">更细子目见原始模板「2材料汇总」或采购标准化台账。</div>`;
  }} else if (tab === 'ji') {{
    html = `<div><strong>机具设备（机）</strong> — ${{mach.file || '未找到机具 per-project 汇总'}}</div>
      <div class="kpi-strip">
        <div>费用合计(泰铢)：<strong>${{(mach.total_thb||0).toLocaleString()}}</strong></div>
        <div>使用次数：${{mach.use_count||'-'}} &nbsp; 机具种类：${{mach.kind_count||'-'}}</div>
      </div>
      <div style="font-size:12px;color:#555;">详细记录在 10_机具费 的处理报告中。master 机具费汇总：${{proj.mach||0}}。</div>`;
  }} else if (tab === 'history') {{
    html = `<div><strong>历史轨迹（来自 master）</strong></div>
      <table class="mini"><thead><tr><th>月份</th><th>预计成本</th><th>毛利率</th><th>结算额</th><th>人工(master)</th><th>机具(master)</th></tr></thead><tbody>
      ${{hist.map(h => `<tr><td>${{h.month}}</td><td class="num">${{h.est}}</td><td class="num">${{h.margin? (h.margin*100).toFixed(1)+'%':'-'}}</td><td class="num">${{h.settle}}</td><td class="num">${{h.labor_master||'-'}}</td><td class="num">${{h.mach_master||'-'}}</td></tr>`).join('')}}
      </tbody></table>`;
  }}

  container.innerHTML = html;

  if (tab === 'ren' && labor.months && labor.months.length) {{
    const el = document.getElementById('laborChart');
    if (el) {{
      const br = labor.breakdown || {{}};
      const ds = Object.keys(br).map((k,i) => ({{label:k, data:br[k], borderWidth:1.5, tension:.2}}));
      const c = new Chart(el, {{type:'line', data:{{labels:labor.months, datasets:ds}}, options:{{responsive:true, maintainAspectRatio:false, plugins:{{legend:{{position:'bottom',labels:{{font:{{size:9}}}}}}}} }});
      detailCharts.push(c);
    }}
  }}
}}

function initOverviewCharts() {
  const mEl = document.getElementById('trendMargin');
  const cEl = document.getElementById('trendCost');
  if (mEl && marginData) {
    const labels = months;
    const mds = Object.keys(marginData).map((k,i)=>({label:k,data:marginData[k],borderColor:['#2b8a6e','#3f6fcf','#c48a2f'][i%3],borderWidth:2,tension:.25,spanGaps:true}));
    new Chart(mEl,{type:'line',data:{labels,datasets:mds},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>(v*100).toFixed(0)+'%'}}}}});
  }
  if (cEl && costData) {
    const labels = months;
    const cds = Object.keys(costData).map((k,i)=>({label:k,data:costData[k],borderColor:['#2b8a6e','#3f6fcf'][i%2],borderWidth:2,tension:.2,spanGaps:true}));
    new Chart(cEl,{type:'line',data:{labels,datasets:cds},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}}}});
  }
  const pie = document.getElementById('catPie');
  if (pie && catLabels.length) {
    new Chart(pie,{type:'doughnut',data:{labels:catLabels,datasets:[{data:catValues,backgroundColor:['#2b8a6e','#3f6fcf','#c48a2f','#8e4db4']}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'}}}});
  }
}

function initAll() {{
  populateProjectSelector();
  initOverviewCharts();
  setTimeout(() => {{ const s=document.getElementById('search'); if(s) s.oninput=filterTable; }},20);
  console.log('%c[多层级成本分析] 报告已就绪', 'color:#2b8a6e');
}}
window.addEventListener('load', initAll);
</script>
</body>
</html>
"""
    return html_content
    table_rows = []
    for p in projects_latest:
        margin_pct = f"{p.gross_margin*100:.1f}%" if p.gross_margin else "-"
        margin_class = "good" if (p.gross_margin or 0) >= 0.05 else ("warn" if (p.gross_margin or 0) > 0 else "bad")
        table_rows.append(
            f"<tr data-name='{html_escape.escape(p.name)}' data-margin='{p.gross_margin or 0}'>"
            f"<td class='proj'><span class='code'>{html_escape.escape(p.code or '-')}</span><br><strong>{html_escape.escape(p.name)}</strong></td>"
            f"<td>{p.month}</td>"
            f"<td class='num'>{p.est_cost:,.1f}</td>"
            f"<td class='num'>{p.gross_profit:,.1f}</td>"
            f"<td class='num {margin_class}'>{margin_pct}</td>"
            f"<td class='num'>{p.settlement:,.1f}</td>"
            f"<td class='num'>{p.finance_paid:,.1f}</td>"
            f"<td class='num small'>{p.procurement:,.0f} / {p.machinery:,.0f} / {p.labor:,.0f}</td>"
            f"<td><button class='row-btn' onclick=\"scrollToCard('{[k for k,v in cards_payload.items() if any(c['name']==p.name for c in cards)] or ['card_1'][0]}')\">详情</button></td>"
            f"</tr>"
        )
    table_html = "\n".join(table_rows) if table_rows else "<tr><td colspan='9'>暂无数据</td></tr>"

    # 卡片 sections
    card_sections = []
    for c in cards:
        card_sections.append(f"""
<section class="project-card" id="{c['id']}">
  <div class="card-header">
    <h3>{html_escape.escape(c['name'])} <span class="code">({html_escape.escape(c['code'] or '-')})</span></h3>
    <span class="meta">{c['latest_month']}</span>
    <button class="export-btn" onclick="exportCardPDF('{c['id']}', '{html_escape.escape(c['name'])}')">导出PDF</button>
  </div>
  <div class="kpi-mini">
    <div>预计总成本 <strong>{c['est_cost']:,.1f}</strong> 万</div>
    <div>毛利率 <strong class="{'good' if (c['margin'] or 0)>=0.05 else 'bad'}">{(c['margin']*100) if c['margin'] else '-':.1f}%</strong></div>
    <div>业主结算 <strong>{c['settlement']:,.1f}</strong> 万</div>
    <div>分包财务付 <strong>{c['finance_paid']:,.1f}</strong> 万</div>
  </div>
  <div class="chart-grid">
    <div class="chart-wrap"><canvas id="c_{c['id']}_m"></canvas></div>
    <div class="chart-wrap"><canvas id="c_{c['id']}_a"></canvas></div>
  </div>
  <table class="mini"><thead><tr><th>月份</th><th>预计成本</th><th>毛利率</th><th>结算额</th></tr></thead><tbody>{c['rows_html']}</tbody></table>
</section>
""")
    cards_html = "\n".join(card_sections)

    toc_links = "".join([f'<a class="toc-link" href="#{cid}">{html_escape.escape(t)}</a>' for cid, t in zip([f"card_{i+1}" for i in range(len(cards))], toc[:len(cards)]) ])

    generated = stats["generated_at"]
    src_name = Path(source_file).name

    html_content = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>项目成本统计综合分析报告</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg: #f6f8f5;
  --card: #fff;
  --line: #e2e8e3;
  --txt: #1f2a24;
  --muted: #5c6b63;
  --accent: #2b8a6e;
  --good: #2b8a6e;
  --warn: #c48a2f;
  --bad: #c44a3f;
}}
* {{ box-sizing: border-box; }}
body {{ margin:0; background:var(--bg); color:var(--txt); font-family: "PingFang SC","Noto Sans SC","Microsoft YaHei",system-ui,sans-serif; }}
.container {{ max-width: 1280px; margin: 0 auto; padding: 20px; }}
.hero {{ background: linear-gradient(135deg,#2b8a6e 0%, #1f6348 100%); color:#fff; padding: 24px 28px; border-radius: 16px; margin-bottom: 18px; box-shadow: 0 10px 30px rgba(43,138,110,0.25); }}
h1 {{ margin:0 0 6px; font-size: 26px; }}
.subtitle {{ opacity:0.95; margin:0; }}
.metrics {{ display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin:16px 0; }}
.metric {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:14px 16px; }}
.metric .val {{ font-size:22px; font-weight:700; color:#1f2a24; }}
.metric .lbl {{ font-size:12px; color:var(--muted); margin-top:2px; }}
.section {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:18px; margin:16px 0; box-shadow:0 4px 12px rgba(0,0,0,0.04); }}
h2 {{ margin:0 0 12px; font-size:18px; border-left:4px solid var(--accent); padding-left:10px; }}
.controls {{ display:flex; flex-wrap:wrap; gap:10px; margin-bottom:12px; align-items:center; }}
input[type="text"] {{ padding:8px 12px; border:1px solid var(--line); border-radius:8px; min-width:220px; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
th,td {{ padding:9px 8px; border-bottom:1px solid var(--line); text-align:left; vertical-align:middle; }}
th {{ background:#f1f5f2; color:#31453a; font-weight:600; position:sticky; top:0; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
.good {{ color:var(--good); font-weight:600; }}
.warn {{ color:var(--warn); font-weight:600; }}
.bad {{ color:var(--bad); font-weight:600; }}
.proj strong {{ font-size:14px; }}
.code {{ font-size:11px; color:#666; background:#f4f7f5; padding:1px 5px; border-radius:3px; }}
.project-card {{ border:1px solid var(--line); border-radius:12px; padding:14px 16px; margin:12px 0; }}
.card-header {{ display:flex; align-items:center; gap:12px; margin-bottom:8px; }}
.card-header h3 {{ margin:0; flex:1; font-size:15px; }}
.meta {{ font-size:12px; color:var(--muted); }}
.chart-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
.chart-wrap {{ position:relative; height:220px; }}
.mini {{ font-size:12px; margin-top:8px; }}
.mini th {{ background:#f8faf8; }}
.kpi-mini {{ display:flex; gap:16px; flex-wrap:wrap; font-size:12px; margin:6px 0 10px; color:#33443b; }}
.kpi-mini strong {{ color:#1f2a24; }}
.row-btn {{ font-size:11px; padding:3px 9px; border:none; background:#2b8a6e; color:#fff; border-radius:5px; cursor:pointer; }}
.toc {{ display:flex; flex-wrap:wrap; gap:6px; margin:8px 0 16px; }}
.toc-link {{ font-size:12px; background:#f1f5f2; padding:4px 9px; border-radius:999px; text-decoration:none; color:#1f2a24; }}
.export-btn, .export-all-btn {{ background:#2b8a6e; color:#fff; border:none; border-radius:6px; padding:5px 11px; font-size:12px; cursor:pointer; }}
.export-all-btn {{ margin:10px auto; display:block; }}
.filter-row {{ display:flex; gap:8px; flex-wrap:wrap; }}
@media (max-width:900px) {{
  .chart-grid {{ grid-template-columns:1fr; }}
  .metrics {{ grid-template-columns: repeat(2,1fr); }}
}}
@media print {{
  .export-btn, .export-all-btn, .row-btn {{ display:none !important; }}
  .project-card {{ break-inside:avoid; page-break-inside:avoid; }}
}}
</style>
</head>
<body>
<div class="container">
  <div class="hero">
    <h1>📊 项目成本统计综合分析报告</h1>
    <p class="subtitle">数据来源：{html_escape.escape(src_name)} &nbsp;|&nbsp; 最新月份：{stats['latest_month']} &nbsp;|&nbsp; 生成：{generated}</p>
    <p style="margin:6px 0 0;opacity:0.9;font-size:13px;">一站式整合成本收集、毛利率、采购/机具/人工/财务/分包等多维度数据，支持交互筛选与趋势分析</p>
  </div>

  <!-- 全局指标 -->
  <div class="metrics">
    <div class="metric"><div class="val">{stats['project_count']}</div><div class="lbl">统计项目数</div></div>
    <div class="metric"><div class="val">{stats['total_est_cost_wan']:,.0f}</div><div class="lbl">最新预计总成本（万）</div></div>
    <div class="metric"><div class="val">{stats['total_settlement_wan']:,.0f}</div><div class="lbl">业主预计结算额（万）</div></div>
    <div class="metric"><div class="val">{stats['total_gross_profit_wan']:,.0f}</div><div class="lbl">毛利合计（万）</div></div>
    <div class="metric"><div class="val">{stats['avg_margin']*100:.1f}%</div><div class="lbl">平均毛利率</div></div>
    <div class="metric"><div class="val">{stats['total_finance_paid_wan']:,.0f}</div><div class="lbl">财务分包已付款（万）</div></div>
    <div class="metric"><div class="val">{stats['total_procurement_wan']:,.0f}</div><div class="lbl">采购金额（万）</div></div>
    <div class="metric"><div class="val">{stats['total_machinery_wan']:,.0f} / {stats['total_labor_wan']:,.0f}</div><div class="lbl">机具 / 人工（万）</div></div>
  </div>

  <!-- 控制栏 -->
  <div class="section">
    <h2>🔍 项目数据总览（最新月份）</h2>
    <div class="controls">
      <input id="search" type="text" placeholder="搜索项目名称或编码..." oninput="filterTable()">
      <div class="filter-row">
        <label style="font-size:12px;color:var(--muted);">毛利率 ≥ <input id="minMargin" type="number" step="0.01" value="-1" style="width:70px" oninput="filterTable()">%</label>
        <button onclick="resetFilters()" style="padding:6px 10px;border:1px solid var(--line);background:#fff;border-radius:6px;">重置筛选</button>
        <button onclick="window.print()" style="padding:6px 10px;background:#2b8a6e;color:#fff;border:none;border-radius:6px;">打印 / 导出PDF</button>
      </div>
    </div>

    <div style="overflow:auto; max-height:420px;">
      <table id="mainTable">
        <thead>
          <tr>
            <th style="min-width:160px">项目</th>
            <th>月份</th>
            <th>预计总成本</th>
            <th>毛利</th>
            <th>毛利率</th>
            <th>业主结算额</th>
            <th>财务分包已付</th>
            <th>采购/机具/人工</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {table_html}
        </tbody>
      </table>
    </div>
    <div style="font-size:12px;color:var(--muted);margin-top:6px;">提示：点击「详情」可快速定位下方项目卡片。表格支持实时搜索与毛利率阈值过滤。</div>
  </div>

  <!-- 趋势 -->
  <div class="section">
    <h2>📈 多项目趋势分析</h2>
    <div class="chart-grid">
      <div class="chart-wrap"><canvas id="trendMargin"></canvas></div>
      <div class="chart-wrap"><canvas id="trendCost"></canvas></div>
    </div>
    <div style="font-size:12px;color:var(--muted);">展示数据较全的代表性项目（最多12个）。时间轴按月份升序。</div>
  </div>

  <!-- 结构 -->
  <div class="section">
    <h2>📊 成本结构概览（最新）</h2>
    <div style="display:flex;gap:20px;flex-wrap:wrap;align-items:center;">
      <div style="flex:1;min-width:280px;height:260px;"><canvas id="catPie"></canvas></div>
      <div style="flex:1;min-width:260px;font-size:13px;line-height:1.7;color:#33443b;">
        说明：基于最新月份各项目「采购 / 机具 / 人工 / 财务分包付款」等关键模块金额汇总。<br>
        实际以各专业台账为准，此处为快速横向对比参考。
      </div>
    </div>
  </div>

  <!-- 项目详情卡片 -->
  <div class="section">
    <h2>🧩 项目详情卡片（可展开趋势 + 最近期次）</h2>
    <div class="toc">{toc_links}</div>
    <div style="text-align:center;margin:6px 0 12px;">
      <button class="export-all-btn" onclick="exportAllCardsPDF()">📦 导出全部项目卡片（打印）</button>
    </div>
    {cards_html}
  </div>

  <!-- 数据质量与说明 -->
  <div class="section" style="background:#fffef5;border-color:#f0e6c8;">
    <h2>⚠️ 数据质量与使用说明</h2>
    <ul style="margin:6px 0 0 18px;font-size:13px;line-height:1.65;color:#534a2f;">
      <li>主数据来自成本统计月度收集（{html_escape.escape(src_name)}），已通过项目名称匹配库标准化。</li>
      <li>「采购/机具/人工」列优先使用对应「汇总」字段，回退到支出字段。</li>
      <li>毛利率颜色：绿色≥5%、橙色>0、红色≤0（仅供快速识别）。</li>
      <li>建议结合各专业模块详细报告使用：人工费HTML、机具租赁分析、财务现金流、采购分析看板等（位于 Z_report 及各 B_sum 目录）。</li>
      <li>生成时间：{generated}。如需更新，先运行成本一键流程，再重新生成本报告。</li>
    </ul>
  </div>

  <div style="text-align:center;margin:30px 0 10px;color:#8a968f;font-size:12px;">
    泰国子公司成本统计系统 · 综合分析报告 · Powered by Grok + 项目现有流程
  </div>
</div>

<script>
const months = {months_json};
const marginData = {margin_series_json};
const costData = {cost_series_json};
const catLabels = {cat_labels_json};
const catValues = {cat_values_json};
const cardsData = {cards_json};
const projectsData = {projects_json};

function filterTable() {{
  const q = (document.getElementById('search').value || '').toLowerCase().trim();
  const minM = parseFloat(document.getElementById('minMargin').value || '-1');
  const rows = document.querySelectorAll('#mainTable tbody tr');
  rows.forEach(r => {{
    const name = (r.getAttribute('data-name') || '').toLowerCase();
    const m = parseFloat(r.getAttribute('data-margin') || '0');
    const hit = (!q || name.includes(q)) && (m >= minM / 100);
    r.style.display = hit ? '' : 'none';
  }});
}}
function resetFilters() {{
  document.getElementById('search').value = '';
  document.getElementById('minMargin').value = '-1';
  filterTable();
}}
function scrollToCard(cardId) {{
  const el = document.getElementById(cardId);
  if (el) {{
    el.scrollIntoView({{ behavior:'smooth', block:'center' }});
    el.style.transition = 'box-shadow .2s';
    el.style.boxShadow = '0 0 0 3px rgba(43,138,110,0.35)';
    setTimeout(() => {{ el.style.boxShadow = '0 4px 12px rgba(0,0,0,0.04)'; }}, 1200);
  }}
}}

let charts = {{}};
function initTrendCharts() {{
  const mEl = document.getElementById('trendMargin');
  const cEl = document.getElementById('trendCost');
  if (!mEl || !cEl) return;

  const labels = months;
  const mDatasets = Object.keys(marginData).map((k, i) => ({{
    label: k,
    data: marginData[k],
    borderColor: ['#2b8a6e','#3f6fcf','#c48a2f','#8e4db4','#667eea','#c44a3f'][i % 6],
    backgroundColor: 'transparent',
    borderWidth: 2,
    tension: 0.25,
    spanGaps: true,
    pointRadius: 2,
  }}));
  new Chart(mEl, {{
    type: 'line',
    data: {{ labels, datasets: mDatasets }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins: {{ legend:{{position:'bottom', labels:{{boxWidth:10, font:{{size:10}}}} }} }},
      scales: {{ y:{{ ticks:{{ callback:v => (v*100).toFixed(0)+'%' }} }} }}
    }}
  }});

  const cDatasets = Object.keys(costData).map((k, i) => ({{
    label: k,
    data: costData[k],
    borderColor: ['#2b8a6e','#3f6fcf','#c48a2f','#8e4db4','#667eea','#c44a3f'][i % 6],
    backgroundColor: 'rgba(43,138,110,0.08)',
    borderWidth: 2,
    tension: 0.2,
    spanGaps: true,
  }}));
  new Chart(cEl, {{
    type: 'line',
    data: {{ labels, datasets: cDatasets }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins: {{ legend:{{position:'bottom', labels:{{boxWidth:10, font:{{size:10}}}} }} }},
      scales: {{ y:{{ ticks:{{ callback:v => Number(v).toLocaleString() }} }} }}
    }}
  }});
}}

function initCatChart() {{
  const el = document.getElementById('catPie');
  if (!el || !catLabels.length) return;
  new Chart(el, {{
    type: 'doughnut',
    data: {{
      labels: catLabels,
      datasets: [{{ data: catValues, backgroundColor: ['#2b8a6e','#3f6fcf','#c48a2f','#8e4db4','#667eea','#c44a3f'] }}]
    }},
    options: {{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{position:'right'}} }} }}
  }});
}}

function initCardCharts() {{
  Object.entries(cardsData).forEach(([cid, series]) => {{
    const mEl = document.getElementById('c_' + cid + '_m');
    const aEl = document.getElementById('c_' + cid + '_a');
    if (!mEl || !aEl) return;

    new Chart(mEl, {{
      type:'line',
      data: {{
        labels: series.months,
        datasets: [
          {{ label:'毛利率', data:series.margin, borderColor:'#3f6fcf', yAxisID:'yR', borderWidth:2, tension:0.25, spanGaps:true }},
          {{ label:'毛利', data:series.profit, borderColor:'#2b8a6e', yAxisID:'yA', borderWidth:2, tension:0.2, spanGaps:true }}
        ]
      }},
      options: {{
        responsive:true, maintainAspectRatio:false,
        scales: {{
          yR: {{ type:'linear', position:'right', ticks:{{ callback:v=>(v*100).toFixed(0)+'%' }} }},
          yA: {{ type:'linear', position:'left', ticks:{{ callback:v=>Number(v).toLocaleString() }} }}
        }},
        plugins: {{ legend:{{ display:false }} }}
      }}
    }});

    new Chart(aEl, {{
      type:'bar',
      data: {{
        labels: series.months,
        datasets: [
          {{ label:'预计总成本', data:series.cost, backgroundColor:'rgba(43,138,110,0.75)' }},
          {{ label:'业主结算额', data:series.settlement, backgroundColor:'rgba(63,111,207,0.65)' }}
        ]
      }},
      options: {{
        responsive:true, maintainAspectRatio:false,
        plugins: {{ legend:{{ position:'bottom', labels:{{ font:{{size:10}} }} }} }},
        scales: {{ y:{{ ticks:{{ callback:v=>Number(v).toLocaleString() }} }} }}
      }}
    }});
  }});
}}

function exportCardPDF(cardId, projectName) {{
  const el = document.getElementById(cardId);
  if (!el) return;
  const all = document.querySelectorAll('.project-card');
  all.forEach(c => c.style.display = (c.id === cardId ? 'block' : 'none'));
  const origTitle = document.title;
  document.title = projectName + ' - 成本分析';
  setTimeout(() => {{
    window.print();
    all.forEach(c => c.style.display = 'block');
    document.title = origTitle;
  }}, 80);
}}

function exportAllCardsPDF() {{
  const btns = document.querySelectorAll('.export-btn, .export-all-btn');
  btns.forEach(b => b.style.display = 'none');
  setTimeout(() => {{
    window.print();
    btns.forEach(b => b.style.display = '');
  }}, 80);
}}

function initAll() {{
  initTrendCharts();
  initCatChart();
  initCardCharts();
  // 初始过滤一次
  setTimeout(filterTable, 30);
  console.log('%c[综合成本分析] HTML 报告已就绪', 'color:#2b8a6e');
}}
window.addEventListener('load', initAll);
</script>
</body>
</html>
"""
    return html_content


def main() -> None:
    parser = argparse.ArgumentParser(description="生成项目成本统计综合分析 HTML（多层级钻取版，支持人/材/机具体数据）")
    parser.add_argument("--input-excel", default=None, help="主汇总Excel路径（项目统计数据汇总_全月份.xlsx）")
    parser.add_argument("--output-html", default=None, help="输出HTML完整路径")
    parser.add_argument("--max-cards", type=int, default=18, help="详情卡片最多显示的项目数（概览用）")
    args = parser.parse_args()

    input_path = find_latest_master_excel(args.input_excel)
    df, all_months, latest_month = load_and_prepare(input_path)

    projects = build_project_latest(df, latest_month)
    stats = compute_global_stats(df, latest_month, projects)
    trend = prepare_trend_data(df, all_months)
    category = prepare_category_breakdown(projects)
    toc, cards_payload, cards = prepare_project_cards_data(df, latest_month, max_projects=args.max_cards)

    # === 新：加载人/材/机明细，支撑多层级钻取 ===
    project_details = attach_details_to_latest_projects(projects, df, latest_month)

    html_str = make_html(
        stats=stats,
        projects_latest=projects,
        trend=trend,
        category=category,
        toc=toc,
        cards_payload=cards_payload,
        cards=cards,
        source_file=str(input_path),
        project_details=project_details,   # 新增：传递给 HTML 做 Tab 内容
    )

    out_path = Path(args.output_html) if args.output_html else DEFAULT_OUTPUT_HTML
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html_str, encoding="utf-8")

    print(f"\n✅ 多层级综合分析HTML已生成：{out_path}")
    print(f"   项目数：{stats['project_count']}，最新月份：{stats['latest_month']}")
    print(f"   已尝试为最新项目附加 人工(人)/采购(材)/机具(机) 明细数据（见详情面板 Tab）。")
    print(f"   同时可参考 Z_report 下的其他专项报告（人工费、机具、预算看板等）。")


if __name__ == "__main__":
    main()
