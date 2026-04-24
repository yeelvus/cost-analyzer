import os
import re
import sys
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent

# ====================== 配置区 ======================
class Config:
    # 输入
    INPUT_SOURCE = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_1_人工费台账/A_人工费台账_raw"
    INPUT_GLOB = "泰国子公司项目人工费统计台账*.xlsx"
    
    # 项目匹配
    MAPPING_EXCEL_PATH = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
    MATCH_DB_PATH = r"/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"
    
    # 输出
    OUTPUT_STD = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std"
    OUTPUT_Z_REPORT = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report"
    
    # HTML
    HTML_TEMPLATE_PATH = SCRIPT_DIR / "人工费汇总_template.html"
    
    MONEY_COLUMNS = ['中方正式工铢', '中方劳务工铢', '子公司属地工铢', '劳务公司铢', '签证费']


# ====================== 路径与模块初始化 ======================
MATCHER_DIR = SCRIPT_DIR.parent / "0_项目基本信息"
LEGACY_MATCHER_DIR = SCRIPT_DIR.parent / "19_代码"

for p in [MATCHER_DIR, LEGACY_MATCHER_DIR]:
    if p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))

from project_name_matcher import (
    build_standard_projects_from_dataframe,
    is_unmatched_alias,
    load_or_init_db,
    match_project,
    save_db,
)
from excel_formatting_utils import apply_common_excel_layout


# ====================== 辅助函数 ======================
def _parse_date_key_from_name(file_name: str):
    stem = Path(file_name).stem
    m = re.search(r'(20\d{2})[.年\-/](\d{1,2})(?:[.月\-/](\d{1,2}))?', stem)
    if not m:
        return None
    y = int(m.group(1))
    mm = int(m.group(2))
    dd = int(m.group(3)) if m.group(3) else 1
    if mm < 1 or mm > 12 or dd < 1 or dd > 31:
        return None
    return y, mm, dd


def resolve_input_file() -> Path:
    p = Path(Config.INPUT_SOURCE).expanduser()
    if p.is_file():
        return p.resolve()
    if p.is_dir():
        candidates = [x for x in p.glob(Config.INPUT_GLOB) if x.is_file()]
        if not candidates:
            raise FileNotFoundError(f"未找到匹配文件：{Config.INPUT_GLOB}")

        def sort_key(fp: Path):
            parsed = _parse_date_key_from_name(fp.name)
            return (1, parsed or (0, 0, 0), fp.stat().st_mtime)

        return sorted(candidates, key=sort_key)[-1].resolve()
    raise FileNotFoundError(f"输入路径不存在：{p}")


def parse_money(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace('\u3000', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.replace('，', '', regex=False)
        .str.replace('（', '(', regex=False)
        .str.replace('）', ')', regex=False)
        .str.replace('－', '-', regex=False)
        .str.replace('–', '-', regex=False)
        .str.replace('—', '-', regex=False)
        .str.replace('THB', '', regex=False)
        .str.replace('thb', '', regex=False)
    )

    def to_num(val):
        if pd.isna(val):
            return 0
        text = str(val).strip()
        if not text:
            return 0
        is_negative = text.startswith('(') and text.endswith(')')
        text = text.strip('()')
        text = re.sub(r'[^0-9.-]', '', text)
        if text.count('-') > 1:
            text = text.replace('-', '', text.count('-') - 1)
        try:
            num = float(text)
        except:
            num = 0
        return -num if is_negative else num

    return cleaned.apply(to_num)


def render_html_from_template(template_path: Path, replacements: dict) -> str:
    if not template_path.exists():
        raise FileNotFoundError(f"HTML模板不存在：{template_path}")
    html = template_path.read_text(encoding='utf-8')
    for key, value in replacements.items():
        html = html.replace(key, value)
    return html


# ====================== 主程序 ======================
def main():
    print("🚀 开始处理人工费台账...")

    excel_file = resolve_input_file()
    print(f"✅ 使用文件：{excel_file.name}")

    df = pd.read_excel(excel_file, sheet_name="基础数据")
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    df.columns = df.columns.str.strip()

    # 列名映射
    column_mapping = {
        '项目名称': ['项目名称'],
        '地区': ['地区'],
        '中方正式工铢': ['中方正式工铢'],
        '中方劳务工铢': ['中方劳务工铢', '中方劳务工株'],
        '子公司属地工铢': ['子公司属地工铢'],
        '劳务公司铢': ['劳务公司铢'],
        '签证费': ['签证费'],
        '年': ['年'],
        '月': ['月']
    }

    found_columns = {}
    for standard, variants in column_mapping.items():
        for col in df.columns:
            if any(v in col for v in variants):
                found_columns[standard] = col
                break

    required = ['项目名称', '中方正式工铢', '中方劳务工铢', '子公司属地工铢', '劳务公司铢', '年', '月']
    missing = [col for col in required if col not in found_columns]
    if missing:
        print(f"\n错误：以下必要列未找到：{missing}")
        sys.exit(1)

    df = df.rename(columns={v: k for k, v in found_columns.items()})

    # 金额转换
    money_columns = Config.MONEY_COLUMNS
    for col in money_columns:
        if col in df.columns:
            df[col] = parse_money(df[col])
        else:
            df[col] = 0.0

    # 数据清洗
    INVALID_PROJECT_NAMES = {"", "nan", "none", "null", "n/a", "na", "-", "--", "/"}
    df = df[df['项目名称'].apply(lambda x: str(x).strip().lower() not in INVALID_PROJECT_NAMES)]
    df = df[~df['项目名称'].astype(str).str.contains('合计|总计|小计|Grand', case=False, na=False)]
    df.reset_index(drop=True, inplace=True)

    # ========== 项目名称标准化 ==========
    print("\n🔄 开始项目名称标准化...")
    if os.path.exists(Config.MAPPING_EXCEL_PATH):
        mapping_xls = pd.ExcelFile(Config.MAPPING_EXCEL_PATH)
        mapping_df = pd.read_excel(mapping_xls, sheet_name=mapping_xls.sheet_names[0])
        mapping_df.columns = [str(c).strip() for c in mapping_df.columns]

        # 列选择逻辑（保留原始完整逻辑）
        def pick_first_existing_column(columns, candidates):
            for name in candidates:
                if name in columns:
                    return name
            return None

        code_col = pick_first_existing_column(mapping_df.columns, ['项目编码', '项目代码', '项目编号', '编码', 'project_code'])
        name_col = pick_first_existing_column(mapping_df.columns, [
            '项目名称（合同管理）', '项目名称(合同管理)', '合同管理项目名称',
            '项目名称（成本梳理）', '项目名称(成本梳理)', '成本梳理项目名称',
            '项目名称', '标准项目名称', 'project_name'
        ])
        status_col = pick_first_existing_column(mapping_df.columns, ['备注', '项目状态', '状态', 'target_status'])

        matcher_info_df = pd.DataFrame({
            '项目名称': mapping_df[name_col],
            '项目编码': mapping_df[code_col] if code_col else '',
            '备注': mapping_df[status_col] if status_col else '需手动核对',
        })

        standard_projects = build_standard_projects_from_dataframe(matcher_info_df)
        match_db = load_or_init_db(Config.MATCH_DB_PATH, standard_projects)

        def standardize_project_name(raw_name):
            if str(raw_name).strip().lower() in INVALID_PROJECT_NAMES:
                return ("", str(raw_name), "")
            raw_name = str(raw_name).strip()
            hit = match_project(raw_name, match_db) if match_db else None
            if hit:
                return (
                    hit.get('project_code', '') or '',
                    hit.get('project_name', raw_name) or raw_name,
                    hit.get('target_status', '') or '',
                )
            return ("", raw_name, "")

        standardized = df['项目名称'].apply(standardize_project_name)
        df['项目编码'] = standardized.apply(lambda x: x[0])
        df['项目名称'] = standardized.apply(lambda x: x[1])
        df['项目状态'] = standardized.apply(lambda x: x[2])

        save_db(Config.MATCH_DB_PATH, match_db)
        print("✅ 项目名称标准化完成（共享匹配库）")
    else:
        print("⚠️ 映射表不存在，使用原始项目名称")
        df['项目编码'] = ""
        df['项目状态'] = ""

    # ========== 聚合 ==========
    grouped = df.groupby(['年', '月', '项目名称', '项目编码'])[money_columns].sum().reset_index()
    grouped['年'] = grouped['年'].astype(int)
    grouped['月'] = grouped['月'].astype(int)
    grouped['年-月'] = grouped.apply(lambda r: f"{int(r['年'])}-{int(r['月']):02d}", axis=1)
    grouped['总人工费'] = grouped[money_columns].sum(axis=1)

    summary_df = df.groupby(['项目编码', '项目名称'], as_index=False)[money_columns].sum()
    summary_df['总人工费'] = summary_df[money_columns].sum(axis=1)
    summary_df = summary_df[['项目编码', '项目名称'] + money_columns + ['总人工费']]

    summary_wan_df = summary_df.copy()
    summary_wan_df[money_columns + ['总人工费']] /= 10000.0
    summary_wan_df = summary_wan_df.rename(columns={c: f"{c}(万泰铢)" for c in money_columns + ['总人工费']})

    # 项目求和
    sum_components = money_columns[:]
    proj_summary = df.groupby(['项目编码', '项目名称'], as_index=False)[sum_components].sum()
    proj_summary['人工费汇总金额'] = proj_summary[sum_components].sum(axis=1)
    proj_summary = proj_summary[['项目编码', '项目名称', '人工费汇总金额']]

    # ====================== 输出 ======================
    output_files = [
        Path(Config.OUTPUT_STD) / "人工费汇总.xlsx",
        Path(Config.OUTPUT_Z_REPORT) / "人工费汇总.xlsx"
    ]

    for out_path in output_files:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        base_name = Path(excel_file).stem

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            proj_summary.to_excel(writer, sheet_name='项目求和', index=False)
            summary_df.to_excel(writer, sheet_name='按项目汇总', index=False)
            summary_wan_df.to_excel(writer, sheet_name='按项目汇总(万泰铢)', index=False)

            raw_output_df = df.drop(columns=['项目状态', '原始项目名称'], errors='ignore')
            raw_output_df.to_excel(writer, sheet_name='原始数据', index=False)

            # 各指标按月透视表
            metrics = money_columns + ['总人工费']
            for metric in metrics:
                pivot = grouped.pivot_table(
                    index=['项目编码', '项目名称'],
                    columns='年-月',
                    values=metric,
                    aggfunc='sum',
                    fill_value=0
                )
                cols_sorted = sorted(pivot.columns, key=lambda s: (int(s.split('-')[0]), int(s.split('-')[1])))
                pivot = pivot.reindex(columns=cols_sorted)
                sheet_name = metric if len(metric) <= 31 else metric[:31]
                pivot.to_excel(writer, sheet_name=sheet_name)

        # 应用样式
        wb = load_workbook(out_path)
        skip_headers = {'项目名称', '地区', '年', '月', '年-月', '项目编码'}
        for ws in wb.worksheets:
            headers = [cell.value for cell in ws[1]]
            numeric_start = next((i+1 for i, h in enumerate(headers) if h and str(h) not in skip_headers), 3)
            apply_common_excel_layout(
                ws,
                numeric_start_col=numeric_start,
                freeze_panes='B2',
                thousand_number_format='#,##0.00'
            )
        wb.save(out_path)
        print(f"✅ 已输出：{out_path}")

    # ====================== HTML 报告 ======================
    print("\n📊 生成交互式 HTML 图表...")
    months_sorted = sorted(grouped['年-月'].unique(), key=lambda s: (int(s.split('-')[0]), int(s.split('-')[1])))
    project_status_map = df.groupby('项目名称')['项目状态'].first().fillna('未知').to_dict()
    project_list = sorted(str(k) for k in project_status_map.keys())

    html_metrics = [m for m in metrics if m != '签证费']
    metrics_data = {}
    for metric in html_metrics:
        pivot = grouped.pivot_table(index='年-月', columns='项目名称', values=metric, aggfunc='sum', fill_value=0)
        pivot = pivot.reindex(index=months_sorted, fill_value=0)
        metric_data = {}
        for project in project_list:
            series = pivot.get(project, pd.Series([0] * len(months_sorted), index=months_sorted))
            metric_data[project] = {
                'x': list(series.index),
                'y': [float(v) if pd.notna(v) else 0.0 for v in series.values]
            }
        metrics_data[metric] = metric_data

    html_file = Path(Config.OUTPUT_Z_REPORT) / f'人工费汇总分析_{base_name}.html'
    html_template = render_html_from_template(
        Config.HTML_TEMPLATE_PATH,
        {
            '__METRICS__': json.dumps(html_metrics, ensure_ascii=False),
            '__METRICS_DATA__': json.dumps(metrics_data, ensure_ascii=False),
            '__PROJECT_STATUS__': json.dumps(project_status_map, ensure_ascii=False),
            '__MONTHS__': json.dumps(months_sorted, ensure_ascii=False),
            '__PROJECTS__': json.dumps(project_list, ensure_ascii=False),
        }
    )
    html_file.write_text(html_template, encoding='utf-8')
    print(f"✅ HTML 图表生成完成：{html_file}")

    print("\n🎉 处理完成！")
    print(f"共记录项目：{len(summary_df)} 个")
    print(f"Excel 输出位置：{Config.OUTPUT_STD} 和 {Config.OUTPUT_Z_REPORT}")


if __name__ == "__main__":
    main()