import pandas as pd
import json
from openpyxl import load_workbook
import os
import sys
import re
from pathlib import Path
from datetime import datetime

SCRIPT_DIR = Path(__file__).resolve().parent
# 项目名称匹配统一从 0_项目基本信息情况 加载，确保与其他模块共用同一 matcher 与数据库。
MATCHER_DIR = SCRIPT_DIR.parent / "/Users/wyattsun/Desktop/code/0_项目基本信息"
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

# ================= 项目名称标准化配置 =================
MAPPING_EXCEL_PATH = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
MATCH_DB_PATH = r"/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"
INPUT_SOURCE = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_1_人工费台账/A_人工费台账_raw"
INPUT_GLOB = "泰国子公司项目人工费统计台账*.xlsx"
OUTPUT_ROOT = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std"
HTML_OUTPUT_ROOT = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/C_数据分析"
HTML_TEMPLATE_PATH = SCRIPT_DIR / "人工费汇总_template.html"
# ====================================================

def _parse_date_key_from_name(file_name: str):
    """从文件名中提取日期键，用于选择最新月份/最新文件。"""
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


def resolve_input_file(source_path: str) -> str:
    """输入可为文件或目录；目录时自动选择最新月份文件。"""
    p = Path(source_path).expanduser()
    if p.is_file():
        return str(p.resolve())

    if p.is_dir():
        candidates = [x for x in p.glob(INPUT_GLOB) if x.is_file()]
        if not candidates:
            raise FileNotFoundError(f"输入目录下未找到匹配文件：{p / INPUT_GLOB}")

        def sort_key(fp: Path):
            parsed = _parse_date_key_from_name(fp.name)
            if parsed is not None:
                return (1, parsed, fp.stat().st_mtime)
            return (0, (0, 0, 0), fp.stat().st_mtime)

        latest = sorted(candidates, key=sort_key)[-1]
        return str(latest.resolve())

    raise FileNotFoundError(f"输入路径不存在：{source_path}")


def render_html_from_template(template_path: Path, replacements: dict) -> str:
    if not template_path.exists():
        raise FileNotFoundError(f"HTML模板不存在：{template_path}")

    html = template_path.read_text(encoding='utf-8')
    for key, value in replacements.items():
        html = html.replace(key, value)
    return html


try:
    excel_file = resolve_input_file(INPUT_SOURCE)
except Exception as e:
    print(f"❌ 自动识别输入文件失败：{e}")
    print("请检查 INPUT_SOURCE 或输入目录中的文件命名后重试。")
    sys.exit(1)

print(f"已选择/指定文件：{excel_file}")

# 读取 "基础数据" sheet
try:
    df = pd.read_excel(excel_file, sheet_name="基础数据")
except ValueError:
    print("错误：文件中没有找到名为 '基础数据' 的工作表！")
    sys.exit(1)
except Exception as e:
    print(f"读取Excel失败：{e}")
    sys.exit(1)

print(f"成功读取 '基础数据' sheet，形状：{df.shape}")
print("项目匹配输入列：基础数据 -> 项目名称")

# 清理空行空列
df.dropna(how='all', inplace=True)
df.dropna(axis=1, how='all', inplace=True)

print("\n识别到的列名：")
print(list(df.columns))

# 标准化列名
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
    print("当前列名：", list(df.columns))
    input("\n按回车退出...")
    exit()

print("列名匹配成功：")
for k, v in found_columns.items():
    print(f"  {k} → '{v}'")

# 重命名列
df = df.rename(columns={v: k for k, v in found_columns.items()})

# 转换费用列为数字（关键修复）
money_columns = ['中方正式工铢', '中方劳务工铢', '子公司属地工铢', '劳务公司铢', '签证费']


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

  def to_num(val: str):
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
    except Exception:
      num = 0
    return -num if is_negative else num

  return cleaned.apply(to_num)


INVALID_PROJECT_NAMES = {"", "nan", "none", "null", "n/a", "na", "-", "--", "/"}


def is_valid_project_name(value) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return text not in INVALID_PROJECT_NAMES


def normalize_text_column(df: pd.DataFrame, col_name: str, default: str = "") -> None:
    if col_name in df.columns:
        df[col_name] = df[col_name].apply(lambda x: str(x).strip() if pd.notna(x) else default)
    else:
        df[col_name] = default


def collect_valid_project_names(series: pd.Series) -> list[str]:
    return sorted({
        str(n).strip()
        for n in series.tolist()
        if is_valid_project_name(n)
    })


for col in money_columns:
  if col in df.columns:
    df[col] = parse_money(df[col])

# 进一步清洗文本列，避免前后空格导致的重复或绘图异常
normalize_text_column(df, '项目名称', '')
normalize_text_column(df, '地区', '未知')
if '签证费' not in df.columns:
    df['签证费'] = 0.0

# 过滤无效行
df = df[df['项目名称'].apply(is_valid_project_name)]
df = df[~df['项目名称'].astype(str).str.contains('合计|总计|小计|Grand', case=False, na=False)]
df.reset_index(drop=True, inplace=True)

# ========== 项目名称标准化（在聚合前） ==========
print("\n🔄 开始项目名称标准化...")

standard_projects = []
match_db = None


def pick_first_existing_column(columns, candidates):
    for name in candidates:
        if name in columns:
            return name
    return None

if os.path.exists(MAPPING_EXCEL_PATH):
    print(f"📋 读取映射表：{MAPPING_EXCEL_PATH}")
    mapping_xls = pd.ExcelFile(MAPPING_EXCEL_PATH)
    mapping_sheet_name = mapping_xls.sheet_names[0]
    mapping_df = pd.read_excel(mapping_xls, sheet_name=mapping_sheet_name)
    print(f"✅ 映射表工作表：{mapping_sheet_name}")

    # 清理列名空格，并按固定优先级选择列名
    mapping_df.columns = [str(c).strip() for c in mapping_df.columns]
    code_col = pick_first_existing_column(
        mapping_df.columns,
        ['项目编码', '项目代码', '项目编号', '编码', 'project_code']
    )
    name_col = pick_first_existing_column(
        mapping_df.columns,
        ['项目名称（合同管理）', '项目名称(合同管理)', '合同管理项目名称', '项目名称（成本梳理）', '项目名称(成本梳理)', '成本梳理项目名称', '项目名称', '标准项目名称', 'project_name']
    )
    status_col = pick_first_existing_column(
        mapping_df.columns,
        ['备注', '项目状态', '状态', 'target_status']
    )

    if not name_col:
        print("❌ 映射表缺少必要列：项目名称")
        print("可接受的示例列名：项目名称（合同管理）/项目名称（成本梳理）/项目长名称/项目名称")
        print(f"当前列名：{list(mapping_df.columns)}")
        sys.exit(1)

    matcher_info_df = pd.DataFrame({
        '项目名称': mapping_df[name_col],
        '项目编码': mapping_df[code_col] if code_col else '',
        '备注': mapping_df[status_col] if status_col else '需手动核对',
    })

    standard_projects = build_standard_projects_from_dataframe(matcher_info_df)
    match_db = load_or_init_db(MATCH_DB_PATH, standard_projects)
    print(
        f"✅ 共享匹配库已加载：标准项目 {len(standard_projects)} 个，"
        f"别名 {len(match_db.get('alias_to_standard', {}))} 个，"
        f"未匹配留痕 {len(match_db.get('unmatched_aliases', []))} 个"
    )
else:
    print(f"⚠️  映射表文件不存在：{MAPPING_EXCEL_PATH}")
    print("将使用原始项目名称")


def standardize_project_name(raw_name):
    if not is_valid_project_name(raw_name):
        return ("", raw_name, "")
    raw_name = str(raw_name).strip()
    if match_db is not None:
        hit = match_project(raw_name, match_db)
        if hit:
            return (
                hit.get('project_code', '') or '',
                hit.get('project_name', raw_name) or raw_name,
                hit.get('target_status', '') or '',
            )
    return ("", raw_name, "")


# 1) 使用共享 matcher 对基础数据 sheet 中项目名称做自动匹配
raw_project_names = collect_valid_project_names(df['项目名称'])
unmatched_projects = []

if match_db is not None:
    print(f"\n🤖 开始共享匹配库自动匹配，待匹配项目：{len(raw_project_names)} 个")
    for raw_name in raw_project_names:
        if is_unmatched_alias(match_db, raw_name):
            continue
        hit = match_project(raw_name, match_db)
        if not hit:
            unmatched_projects.append(raw_name)

    # 2) 自动匹配失败时，不再弹窗；直接保留原名并记录到未匹配留痕
    if unmatched_projects:
        print(f"\n⚠️ 自动匹配未命中 {len(unmatched_projects)} 个项目，已保留原名称（无弹窗）。")
        existing_unmatched = set(match_db.get('unmatched_aliases', []))
        existing_unmatched.update(unmatched_projects)
        match_db['unmatched_aliases'] = sorted(existing_unmatched)
    else:
        print("✅ 所有项目均已自动匹配")

    save_db(MATCH_DB_PATH, match_db)
    print(f"✅ 本次匹配结果已写入共享匹配库：{MATCH_DB_PATH}")


standardized = df['项目名称'].apply(standardize_project_name)
df['项目编码'] = standardized.apply(lambda x: x[0])
df['项目名称'] = standardized.apply(lambda x: x[1])
df['项目状态'] = standardized.apply(lambda x: x[2])

print("✅ 项目名称标准化完成")

remaining_unmatched = []
if match_db is not None:
    for n in raw_project_names:
        if not match_project(n, match_db):
            remaining_unmatched.append(n)
if remaining_unmatched:
    print(f"\n⚠️  仍有 {len(remaining_unmatched)} 个项目未匹配，已保持原名称")
    for name in remaining_unmatched:
        print(f"   - {name}")
elif match_db is not None:
    print("\n✅ 所有项目都已成功匹配")

print("✅ 项目名称标准化完成\n")

# ========== 聚合 ==========
random_cols = ['年', '月', '项目名称', '项目编码']
agg_cols = money_columns
grouped = df.groupby(random_cols)[agg_cols].sum().reset_index()
grouped['年'] = grouped['年'].astype(int)
grouped['月'] = grouped['月'].astype(int)
grouped['年-月'] = grouped.apply(lambda r: f"{int(r['年'])}-{int(r['月']):02d}", axis=1)
grouped['总人工费'] = grouped[agg_cols].sum(axis=1)

# 该表按项目汇总，不再按年月拆分
summary_df = (
    df.groupby(['项目编码', '项目名称'], as_index=False)[agg_cols]
    .sum()
)
summary_df['总人工费'] = summary_df[agg_cols].sum(axis=1)
summary_df = summary_df[['项目编码', '项目名称'] + agg_cols + ['总人工费']]

# 新增万泰铢口径汇总（金额列统一除以 10000）
summary_wan_df = summary_df.copy()
summary_amount_cols = agg_cols + ['总人工费']
summary_wan_df[summary_amount_cols] = summary_wan_df[summary_amount_cols] / 10000.0
summary_wan_df = summary_wan_df.rename(columns={c: f"{c}(万泰铢)" for c in summary_amount_cols})

# 输出目录（固定到 B_输出标准化人工费台账）
base_name = os.path.basename(excel_file).rsplit('.', 1)[0]
output_folder = OUTPUT_ROOT
os.makedirs(output_folder, exist_ok=True)
os.makedirs(HTML_OUTPUT_ROOT, exist_ok=True)
run_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

metrics = agg_cols + ['总人工费']
# HTML指标永久排除签证费（Excel汇总仍保留签证费相关数据）
html_metrics = [m for m in metrics if m != '签证费']

# 项目求和汇总（仅保留：项目编码、项目名称、人工费汇总金额）
hidden_export_cols = ['中方正式工元', '中方劳务工元']
sum_components = ['中方正式工铢', '中方劳务工铢', '子公司属地工铢', '劳务公司铢', '签证费']

for col in sum_components:
    if col not in df.columns:
        df[col] = 0.0

proj_summary = (
    df.groupby(['项目编码', '项目名称'], as_index=False)[sum_components]
    .sum()
)
proj_summary['人工费汇总金额'] = proj_summary[sum_components].sum(axis=1)
proj_summary = proj_summary[['项目编码', '项目名称', '人工费汇总金额']]

combined_file = os.path.join(output_folder, f'人工费汇总_{run_timestamp}.xlsx')
raw_output_df = df.drop(columns=['项目状态', '原始项目名称'], errors='ignore')
raw_output_df = raw_output_df.drop(columns=hidden_export_cols, errors='ignore')
with pd.ExcelWriter(combined_file, engine='openpyxl') as writer:
    proj_summary.to_excel(writer, sheet_name='项目求和', index=False)
    summary_df.to_excel(writer, sheet_name='按项目汇总', index=False)
    summary_wan_df.to_excel(writer, sheet_name='按项目汇总(万泰铢)', index=False)
    raw_output_df.to_excel(writer, sheet_name='原始数据', index=False)
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

print("\n📊 生成交互式 HTML 图表（带筛选/折叠）...")

months_sorted = sorted(grouped['年-月'].unique(), key=lambda s: (int(s.split('-')[0]), int(s.split('-')[1])))
project_status_map = df.groupby('项目名称')['项目状态'].first().fillna('未知').to_dict()
project_status_map = {k: (v if v else '未知') for k, v in project_status_map.items()}
project_list = sorted(str(k) for k in project_status_map.keys())

metrics_data = {}
for metric in html_metrics:
    metric_data = {}
    pivot = grouped.pivot_table(index='年-月', columns='项目名称', values=metric, aggfunc='sum', fill_value=0)
    pivot = pivot.reindex(index=months_sorted, fill_value=0)
    for project in project_list:
        series = pivot[project] if project in pivot.columns else pd.Series([0] * len(months_sorted), index=months_sorted)
        metric_data[project] = {'x': list(series.index), 'y': [float(v) if pd.notna(v) else 0.0 for v in series.values]}
    metrics_data[metric] = metric_data

html_file = os.path.join(HTML_OUTPUT_ROOT, f'人工费汇总分析_{base_name}.html')
html_template = render_html_from_template(
        HTML_TEMPLATE_PATH,
        {
                '__METRICS__': json.dumps(html_metrics, ensure_ascii=False),
                '__METRICS_DATA__': json.dumps(metrics_data, ensure_ascii=False),
                '__PROJECT_STATUS__': json.dumps(project_status_map, ensure_ascii=False),
                '__MONTHS__': json.dumps(months_sorted, ensure_ascii=False),
                '__PROJECTS__': json.dumps(project_list, ensure_ascii=False),
        }
)

with open(html_file, 'w', encoding='utf-8') as f:
    f.write(html_template)

print(f"✅ 已生成交互式 HTML 图表：{html_file}")

wb = load_workbook(combined_file)
skip_headers = {'项目名称', '地区', '年', '月', '年-月', '项目编码'}


def find_first_numeric_col(ws):
    headers = [cell.value for cell in ws[1]]
    sample_end = min(ws.max_row, 10)
    for idx, h in enumerate(headers, start=1):
        h_str = str(h) if h is not None else ''
        if h_str in skip_headers:
            continue
        for r in range(2, sample_end + 1):
            val = ws.cell(row=r, column=idx).value
            if isinstance(val, (int, float)):
                return idx
    return ws.max_column + 1


for ws in wb.worksheets:
    numeric_start_col = find_first_numeric_col(ws)
    apply_common_excel_layout(
        ws,
        numeric_start_col=numeric_start_col,
        freeze_panes='B2',
        thousand_number_format='#,##0.00'
    )

wb.save(combined_file)

months = sorted(grouped['年-月'].unique())
print(f"\n✅ 处理完成！处理的年月：{months}")
print(f"共记录项目：{len(summary_df['项目名称'].unique())} 个")
print(f"输出文件夹：{output_folder}")
print("   📊 输出文件：")
print(f"   ├─ {os.path.basename(combined_file)}  (包含多个sheet的Excel汇总表)")
print(f"   └─ {html_file}        (各指标趋势图汇总，支持交互)")
