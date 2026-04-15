#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【已修改版】项目成本分析报告生成器
新增功能：一个汇总报告 + 增强每个项目的详细报告
作者：Grok 为你修改
"""

import pandas as pd
from pathlib import Path
import re
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ==================== 配置区 ====================
BASE_DIR = Path(__file__).resolve().parent
SUMMARY_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集/B_成本统计_sum")
INPUT_EXCEL = SUMMARY_DIR / "项目统计数据汇总_全月份.xlsx"
OUTPUT_DIR = SUMMARY_DIR / "项目成本分析报告_TEXT"

# 需要转为数值的列
EXCEL_NUMERIC_COLUMNS = [
    "预计总成本", "计算总成本", "毛利", "毛利率", "实际已发生金额", "业主预计结算额",
    "总包合同额", "业主签证", "分包结算额", "分包合同额", "分包签证", "财务分包已付款",
    "材料采购实际已发生", "采购部采购金额", "采购数据汇总", "项目部采购金额",
    "商混站混凝土汇总", "材料租赁费用", "工资支出", "人员调遣支出", "人工费汇总",
    "机具费支出", "机具费汇总", "其他费用支出", "临建费用支出", "后续预估汇总",
    "材料采购后续预估发生", "材料租赁后续预估", "机具费后续预估",
    "人员调遣后续预估", "其他费用后续预估", "临建费用后续预估",
]

REQUIRED_BASE_COLUMNS = ["月份", "项目编码", "项目名称"]


def load_summary_dataframe(input_excel: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(input_excel, engine="openpyxl", engine_kwargs={"data_only": True})
    except TypeError:
        return pd.read_excel(input_excel, engine="openpyxl")


def resolve_input_excel() -> Path:
    if INPUT_EXCEL.exists():
        return INPUT_EXCEL
    candidates = sorted(
        [p for p in SUMMARY_DIR.glob("*项目统计数据汇总*.xlsx") if p.is_file() and not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]
    raise FileNotFoundError(f"未找到汇总文件，请检查目录: {SUMMARY_DIR}")


def safe_to_float(val):
    if pd.isna(val) or val is None:
        return 0.0
    try:
        return float(val)
    except:
        return 0.0


def clean_filename(name):
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name[:100]


def is_significant_change(val1, val2, threshold=5.0):
    return abs(val1 - val2) >= threshold


def get_metric(row, *candidates):
    for col in candidates:
        if col in row.index:
            return safe_to_float(row.get(col, 0))
    return 0.0


def get_total_cost(row):
    total = get_metric(row, "计算总成本", "计算总成本（万元）", "计算总成本(万元)")
    if total != 0:
        return total
    occurred = get_metric(row, "实际已发生金额")
    forecast = get_metric(row, "后续预估汇总")
    return occurred + forecast


# ==================== 新增：生成汇总报告 ====================
def generate_summary_report(df: pd.DataFrame, output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)
    summary_file = output_dir / f"全项目成本汇总分析报告_最新月份_{datetime.now().strftime('%Y%m%d')}.txt"
    
    # 按项目分组，取每个项目最新的2条记录（当前 + 上期）
    project_groups = df.groupby("项目名称")
    
    txt = []
    txt.append("=== 全项目成本汇总分析报告（最新月份）===")
    txt.append(f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    txt.append(f"共统计 {len(project_groups)} 个项目\n")
    txt.append("=" * 100)

    for project_name, group in project_groups:
        # 按月份降序，取最新和上期
        sorted_group = group.sort_values(by='月份', ascending=False)
        current = sorted_group.iloc[0]
        prev = sorted_group.iloc[1] if len(sorted_group) >= 2 else None
        
        month = current.get('月份', 'N/A')
        code = current.get('项目编码', 'N/A')
        
        # 关键指标
        est_total = get_metric(current, "预计总成本")
        calc_total = get_total_cost(current)
        diff_total = calc_total - est_total
        
        # 毛利率
        current_margin = safe_to_float(current.get('毛利率', 0))
        prev_margin = safe_to_float(prev.get('毛利率', 0)) if prev is not None else 0
        
        txt.append(f"\n【{project_name}】  项目编码：{code}   最新月份：{month}")
        txt.append(f"  预计总成本：{est_total:,.2f} 万元")
        txt.append(f"  计算总成本：{calc_total:,.2f} 万元")
        if abs(diff_total) > 0:
            txt.append(f"  → 差异：{diff_total:+,.2f} 万元   【需要核对确认】")
        else:
            txt.append("  → 预计总成本与计算总成本一致")
        
        # 毛利率变化
        if prev is not None and current_margin != prev_margin:
            margin_chg = current_margin - prev_margin
            direction = "降低" if margin_chg < 0 else "升高"
            txt.append(f"  毛利率较上期{direction}：{margin_chg:+.2f}%")
            if margin_chg < 0:
                txt.append("     建议重点检查材料成本、人工费、机具费是否超支")
        
        # 后续预估无变化提醒
        forecast_cols = ["后续预估汇总", "材料采购后续预估发生", "材料租赁后续预估",
                         "机具费后续预估", "人员调遣后续预估", "其他费用后续预估", "临建费用后续预估"]
        no_change_list = []
        for col in forecast_cols:
            if col in current.index and prev is not None and col in prev.index:
                if abs(safe_to_float(current[col]) - safe_to_float(prev[col])) < 1:
                    no_change_list.append(col)
        if no_change_list:
            txt.append(f"  ⚠️ 后续预估以下字段较上期无变化：{', '.join(no_change_list)}")
        
        # 其他关键对比（简略版）
        purchase_dev = get_metric(current, "采购部采购金额") - get_metric(current, "采购数据汇总")
        if abs(purchase_dev) >= 10:
            txt.append(f"  采购部采购金额 vs 采购数据汇总差异：{purchase_dev:+,.2f} 万元 【需核对】")
        
        labor_dev = (get_metric(current, "工资支出") + get_metric(current, "人员调遣支出")) - get_metric(current, "人工费汇总")
        if abs(labor_dev) >= 10:
            txt.append(f"  人工费内部差异（工资+调遣 vs 人工费汇总）：{labor_dev:+,.2f} 万元")
        
        tool_dev = get_metric(current, "机具费支出") - get_metric(current, "机具费汇总")
        if abs(tool_dev) >= 10:
            txt.append(f"  机具费内部差异：{tool_dev:+,.2f} 万元")
        
        txt.append("-" * 80)
    
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write("\n".join(txt))
    
    print(f"✅ 汇总报告已生成：{summary_file.name}")


# ==================== 增强版：单个项目详细报告 ====================
def generate_project_text_report(df_project, project_name):
    df_proj_sorted = df_project.sort_values(by='月份', ascending=False)
    current = df_proj_sorted.iloc[0]
    prev = df_proj_sorted.iloc[1] if len(df_proj_sorted) >= 2 else None
    
    txt = []
    txt.append(f"{project_name} 项目成本详细对比分析报告")
    txt.append(f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    txt.append(f"当前月份：{current.get('月份', 'N/A')}")
    txt.append("=" * 90)
    
    # 1. 预计 vs 计算总成本
    est = get_metric(current, "预计总成本")
    calc = get_total_cost(current)
    diff = calc - est
    txt.append("1. 预计总成本 vs 计算总成本")
    txt.append(f"   预计总成本：{est:,.2f} 万元")
    txt.append(f"   计算总成本：{calc:,.2f} 万元")
    if abs(diff) > 0:
        txt.append(f"   → 差异：{diff:+,.2f} 万元   【需要核对确认】")
    else:
        txt.append("   → 两者一致")
    txt.append("")
    
    # 2. 毛利率变化
    curr_margin = safe_to_float(current.get('毛利率', 0))
    if prev is not None:
        prev_margin = safe_to_float(prev.get('毛利率', 0))
        chg = curr_margin - prev_margin
        txt.append(f"2. 毛利率较上期：{chg:+.2f}%")
        if chg < 0:
            txt.append("   ⚠️ 毛利率降低，建议重点核查材料成本、人工费、机具费是否超支")
        txt.append("")
    
    # 3. 采购、人工、机具对比
    purchase_dev = get_metric(current, "采购部采购金额") - get_metric(current, "采购数据汇总")
    if abs(purchase_dev) >= 10:
        txt.append(f"3. 采购部采购金额 vs 采购数据汇总 差异：{purchase_dev:+,.2f} 万元 【需核对】")
    
    labor_dev = (get_metric(current, "工资支出") + get_metric(current, "人员调遣支出")) - get_metric(current, "人工费汇总")
    if abs(labor_dev) >= 10:
        txt.append(f"   人工费内部差异：{labor_dev:+,.2f} 万元")
    
    tool_dev = get_metric(current, "机具费支出") - get_metric(current, "机具费汇总")
    if abs(tool_dev) >= 10:
        txt.append(f"   机具费内部差异：{tool_dev:+,.2f} 万元")
    txt.append("")
    
    # 4. 后续预估无变化提醒
    forecast_cols = ["后续预估汇总", "材料采购后续预估发生", "材料租赁后续预估",
                     "机具费后续预估", "人员调遣后续预估", "其他费用后续预估", "临建费用后续预估"]
    no_change = []
    if prev is not None:
        for col in forecast_cols:
            if col in current.index and col in prev.index:
                if abs(safe_to_float(current[col]) - safe_to_float(prev[col])) < 1:
                    no_change.append(col)
    if no_change:
        txt.append(f"⚠️ 后续预估以下字段较上期无变化：{', '.join(no_change)}，请确认是否需要调整预估")
    
    txt.append("=" * 90)
    return "\n".join(txt)


# ==================== 主程序 ====================
def main():
    input_excel = resolve_input_excel()
    print(f"正在读取汇总文件: {input_excel}")
    
    df = load_summary_dataframe(input_excel)
    
    # 数据清洗
    for col in EXCEL_NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(safe_to_float)
    
    # 新增：生成汇总报告
    generate_summary_report(df, OUTPUT_DIR)
    
    # 原有：生成每个项目的详细报告（已增强）
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    project_groups = df.groupby("项目名称")
    print(f"共处理 {len(project_groups)} 个项目...")
    
    for project_name, group in project_groups:
        clean_name = clean_filename(project_name)
        txt_filename = f"{clean_name}_成本对比分析报告.txt"
        txt_path = OUTPUT_DIR / txt_filename
        
        report_text = generate_project_text_report(group, project_name)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(report_text)
        print(f"✓ 已生成详细报告：{txt_filename}")
    
    print(f"\n🎉 全部完成！")
    print(f"汇总报告路径：{OUTPUT_DIR}")
    print(f"详细报告路径：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()