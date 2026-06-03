import os
import sys
from pathlib import Path
import pandas as pd
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import re
import importlib.util

# ── common 模块加载（单例，避免孤立实例）──────────────────────────────────
_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.project import (
    add_mapping,
    get_project_lookup,
    is_unmatched_alias,
    load_or_init_db,
    mark_unmatched,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
)
from common.excel import apply_common_excel_layout
from settings import MATCH_DB_PATH

# ============= 配置区域 =============
# 文件夹按关键词动态查找，忽略 A_ / B_ / C_ / D_ 等前缀字母
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(Path(__file__).parent / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    # 动态加载失败时回退到脚本目录，避免脚本直接崩溃
    _fallback_root = Path(__file__).parent
    RAW_DATA_ROOT = str(_fallback_root)
    STANDARDIZED_DATA_ROOT = str(_fallback_root)
    SUMMARY_OUTPUT_ROOT = str(_fallback_root)
    SIMPLIFY_DATA_ROOT = str(_fallback_root)
    PROCUREMENT_STD_MODULE_PATH = ""
else:
    _路径_mod = importlib.util.module_from_spec(_路径_spec)
    _路径_spec.loader.exec_module(_路径_mod)
    RAW_DATA_ROOT = str(_路径_mod.RAW_DIR)
    STANDARDIZED_DATA_ROOT = str(_路径_mod.STD_DIR)
    SUMMARY_OUTPUT_ROOT = str(_路径_mod.SUM_DIR)
    SIMPLIFY_DATA_ROOT = str(_路径_mod.SIMPLIFY_DIR)
    PROCUREMENT_STD_MODULE_PATH = str(_路径_mod.find_script("采购部单文件标准化输出"))

PROJECT_INFO_FILE = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
PROJECT_MATCH_DB_FILE = str(MATCH_DB_PATH)
STD_MERGED_SUFFIX = 'std_merged.xlsx'
SIMPLIFY_OUTPUT_FILENAME = "采购台账_simplify_std_merged.xlsx"
DIRECT_STD_INPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/C_采购部采购台账_simplify"
DIRECT_OUTPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/D_采购部采购台账_sum"
DIRECT_OUTPUT_FILENAME = "采购部采购台账_std_merged.xlsx"
# ===================================

# 全局匹配库（来自 0_项目基本信息/project_name_matcher.py）
PROJECT_MATCH_DB = None

FINAL_OUTPUT_COLUMNS = [
    '项目编码', '项目名称', '材料归类', '物品名称', '规格', '单位', '数量',
    '单价', '不含税金额', '7%增值税', '含税金额', '货币种类', '汇率', '年月',
    '标准日期', '供货商名称', '备注', '签订日期'
]


def _load_module_from_path(module_name: str, file_path: str):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def apply_unified_workbook_format(output_wb):
    """统一使用excel_formatting_utils模块格式化输出工作簿。"""
    amount_keywords = ('金额', '单价', '税')

    for sheet_name in ['合并数据', '项目汇总', '采购总金额']:
        if sheet_name not in output_wb.sheetnames:
            continue

        ws = output_wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        amount_headers = [
            h for h in headers
            if isinstance(h, str) and any(k in h for k in amount_keywords)
        ]

        if sheet_name == '采购总金额':
            apply_common_excel_layout(
                ws,
                numeric_start_col=3,
                amount_headers=[h for h in headers[2:] if isinstance(h, str)],
                header_row=1,
                header_height=40,
                freeze_panes='A2',
            )
        else:
            apply_common_excel_layout(
                ws,
                numeric_start_col=ws.max_column + 1,
                amount_headers=amount_headers,
                header_row=1,
                header_height=40,
                freeze_panes='A2',
            )


def _period_key(period_text):
    y, m = str(period_text).split('-')
    return int(y), int(m)


def build_procurement_total_sheet(merged_df):
    """由合并数据构建'采购总金额'宽表（项目 × 年月）。"""
    required = {'项目编码', '项目名称', '年月', '含税金额'}
    if not required.issubset(set(merged_df.columns)):
        return pd.DataFrame(columns=['项目编码', '项目名称'])

    work = merged_df[['项目编码', '项目名称', '年月', '含税金额']].copy()
    work['项目编码'] = work['项目编码'].fillna('').astype(str).str.strip()
    work['项目名称'] = work['项目名称'].fillna('').astype(str).str.strip()
    work['年月'] = work['年月'].fillna('').astype(str).str.strip()
    work['含税金额'] = pd.to_numeric(work['含税金额'], errors='coerce')

    work = work[(work['项目编码'] != '') | (work['项目名称'] != '')]
    work = work[work['年月'].str.fullmatch(r'(?:19|20)\d{2}-(?:0[1-9]|1[0-2])', na=False)]
    work = work[work['含税金额'].notna()]
    if work.empty:
        return pd.DataFrame(columns=['项目编码', '项目名称'])

    grouped = (
        work.groupby(['项目编码', '项目名称', '年月'], as_index=False)
        .agg(含税金额=('含税金额', 'sum'))
    )

    periods = sorted(grouped['年月'].unique().tolist(), key=_period_key)
    label_map = {p: f"{int(p.split('-')[0])}年{int(p.split('-')[1])}月" for p in periods}

    pivot = grouped.pivot_table(
        index=['项目编码', '项目名称'],
        columns='年月',
        values='含税金额',
        aggfunc='sum',
        fill_value=0.0,
    ).reindex(columns=periods, fill_value=0.0)

    amount_wide = pivot.rename(columns=label_map).reset_index()
    sort_cols = [label_map[p] for p in reversed(periods)]
    amount_wide['__编码为空'] = amount_wide['项目编码'].astype(str).str.strip().eq('')
    amount_wide = amount_wide.sort_values(
        ['__编码为空', *sort_cols],
        ascending=[True] + [False] * len(sort_cols),
    )
    amount_wide = amount_wide.drop(columns=['__编码为空']).reset_index(drop=True)
    return amount_wide

def load_project_config():
    """从 Excel 文件加载项目配置数据"""
    config_path = PROJECT_INFO_FILE
    try:
        # 读取Excel文件
        df = pd.read_excel(config_path, sheet_name=0)
        
        # 检查必要列
        required_cols = ['项目编码', '项目名称']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            print(f"警告：Excel文件缺少必要列 {missing_cols}，返回空配置")
            return []
        
        # 转换为字典列表格式（兼容原JSON格式）
        projects = []
        for idx, row in df.iterrows():
            project_dict = {
                'project_code': str(row['项目编码']).strip() if pd.notna(row['项目编码']) else '',
                'project_name': str(row['项目名称']).strip() if pd.notna(row['项目名称']) else '',
            }
            # 添加备注列（如果存在）
            if '备注' in df.columns and pd.notna(row['备注']):
                project_dict['remark'] = str(row['备注']).strip()
            
            # 只添加项目编码或项目名称非空的项目
            if project_dict['project_code'] or project_dict['project_name']:
                projects.append(project_dict)
        
        print(f"已加载 {len(projects)} 个项目配置: {config_path}")
        return projects
    except Exception as e:
        print(f"警告：无法加载项目配置文件 {config_path}: {e}")
        return []

def extract_project_name(filename):
    """从文件名提取项目名称：优先取第一个下划线前的内容。"""
    base_name = os.path.splitext(filename)[0]  # Remove extension
    
    # 需求：项目名称来自文件名中“_”之前的内容
    if '_' in base_name:
        return base_name.split('_', 1)[0].strip()
    
    # 兜底：尝试匹配 "XX项目" 模式
    match = re.match(r'^(.+?项目)', base_name)
    if match:
        return match.group(1).strip()
    
    return base_name.strip()  # 否则使用整个文件名

def find_matching_project(extracted_name, project_config):
    """查找匹配的项目配置（优先使用共享匹配库）。"""
    global PROJECT_MATCH_DB

    if PROJECT_MATCH_DB is not None:
        hit = match_project(extracted_name, PROJECT_MATCH_DB)
        if hit:
            return {
                'project_code': hit.get('project_code', '需手动核对'),
                'project_name': hit.get('project_name', extracted_name),
            }

    # 精确匹配
    for project in project_config:
        if project['project_name'] == extracted_name:
            return project
    
    # 模糊匹配（包含关系）
    for project in project_config:
        if extracted_name in project['project_name'] or project['project_name'] in extracted_name:
            return project
    
    return None


def resolve_project_match(extracted_name, project_config):
    """统一从共享匹配库解析项目名；未匹配且被留空的名称也会被持久化识别。"""
    matched_project = find_matching_project(extracted_name, project_config)
    if matched_project:
        if PROJECT_MATCH_DB is not None:
            add_mapping(PROJECT_MATCH_DB, extracted_name, matched_project['project_name'])
        return {
            'status': 'matched',
            'project_code': matched_project.get('project_code', ''),
            'project_name': matched_project.get('project_name', extracted_name),
        }

    if PROJECT_MATCH_DB is not None and is_unmatched_alias(PROJECT_MATCH_DB, extracted_name):
        return {
            'status': 'unmatched',
            'project_code': '',
            'project_name': extracted_name,
        }

    return None


def extract_year_month_key(folder_name):
    """从文件夹名提取 (year, month) 键值，提取失败返回 None。"""
    m = re.search(r'(\d{4})年(\d{1,2})月', str(folder_name))
    if not m:
        return None
    year = int(m.group(1))
    month = int(m.group(2))
    if 1 <= month <= 12:
        return (year, month)
    return None


def get_month_folders(raw_root):
    """获取原始数据目录下的月份子目录，按年月升序返回。"""
    if not os.path.isdir(raw_root):
        return []

    candidates = []
    for name in os.listdir(raw_root):
        path = os.path.join(raw_root, name)
        if os.path.isdir(path):
            ym_key = extract_year_month_key(name)
            candidates.append((name, path, ym_key))

    # 优先按年月排序；无法识别年月的目录排在后面（按名称）
    with_ym = [x for x in candidates if x[2] is not None]
    without_ym = [x for x in candidates if x[2] is None]
    with_ym.sort(key=lambda x: x[2])
    without_ym.sort(key=lambda x: x[0])
    return with_ym + without_ym


def choose_target_folders(month_folders):
    """直接返回全部月份文件夹，不再弹出选择。"""
    return month_folders


def list_source_excel_files(folder_path):
    """列出待合并的源Excel文件（返回文件名，供原始数据汇总使用）。"""
    return [
        f for f in os.listdir(folder_path)
        if (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsm'))
        and f != 'merged_excel.xlsx'
        and not f.startswith('~$')
        and not f.startswith('merged_excel')
        and not f.endswith('_std_merged.xlsx')
    ]


def list_source_excel_files_recursive(folder_path):
    """递归列出待合并的源Excel文件（返回绝对路径，供标准化汇总使用）。"""
    excel_files = []
    for root, _, files in os.walk(folder_path):
        for f in files:
            lower = f.lower()
            if not (lower.endswith('.xlsx') or lower.endswith('.xls') or lower.endswith('.xlsm')):
                continue
            if f == 'merged_excel.xlsx':
                continue
            if f.startswith('~$') or f.startswith('merged_excel') or f.endswith('_std_merged.xlsx'):
                continue
            excel_files.append(os.path.join(root, f))
    return sorted(excel_files)


def merge_standardized_excels(folder_path):
    """仅汇总已标准化文件，不再执行二次标准化。"""
    excel_paths = list_source_excel_files_recursive(folder_path)

    print("\n" + "=" * 80)
    print(f"【纯汇总模式】开始处理文件夹: {folder_path}")
    print(f"【纯汇总模式】发现 {len(excel_paths)} 个Excel文件")
    print("=" * 80 + "\n")

    if not excel_paths:
        print("⚠️  警告：未找到任何Excel文件！")
        return None, None

    output_wb = openpyxl.Workbook()
    merged_ws = output_wb.active
    if merged_ws is None:
        raise RuntimeError('无法创建工作表: 合并数据')
    merged_ws.title = '合并数据'
    merged_ws.append(FINAL_OUTPUT_COLUMNS)

    current_row = 2
    project_totals = defaultdict(float)

    for idx, file_path in enumerate(excel_paths, start=1):
        file_name = os.path.basename(file_path)
        print(f"[{idx}/{len(excel_paths)}] 汇总文件: {file_name}")

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"  ✗ 读取失败: {e}")
            continue

        if df.empty:
            print("  ⚠️ 空文件，跳过")
            continue

        for col in FINAL_OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = ''
        df = df[FINAL_OUTPUT_COLUMNS]

        for _, row in df.iterrows():
            for col_idx, col_name in enumerate(FINAL_OUTPUT_COLUMNS, start=1):
                value = row[col_name]
                merged_ws.cell(current_row, col_idx).value = value
                if col_name in {'单价', '不含税金额', '7%增值税', '含税金额'} and isinstance(value, (int, float)):
                    merged_ws.cell(current_row, col_idx).number_format = '0.00'

            project_code = '' if pd.isna(row['项目编码']) else str(row['项目编码']).strip()
            project_name = '' if pd.isna(row['项目名称']) else str(row['项目名称']).strip()
            amount_value = pd.to_numeric(pd.Series([row['含税金额']]), errors='coerce').iloc[0]
            if pd.notna(amount_value):
                project_totals[(project_code, project_name)] += float(amount_value)

            current_row += 1

    if current_row == 2:
        print("⚠️  没有找到任何有效数据！")
        return None, None

    merged_ws.row_dimensions[1].height = 40
    for cell in merged_ws[1]:
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    merged_ws.freeze_panes = 'A2'
    last_col = get_column_letter(len(FINAL_OUTPUT_COLUMNS))
    merged_ws.auto_filter.ref = f'A1:{last_col}{current_row - 1}'

    summary_ws = output_wb.create_sheet('项目汇总')
    summary_ws.append(['项目编码', '项目名称', '总采购金额'])
    for row_idx, ((project_code, project_name), total) in enumerate(sorted(project_totals.items()), start=2):
        summary_ws.append([project_code, project_name, total])
        summary_ws.cell(row_idx, 3).number_format = '0.00'

    return output_wb, output_wb.sheetnames


def preload_project_matches_for_targets(targets, project_config):
    """一次性扫描所有目标月份，统一完成项目匹配，避免逐月弹窗。"""
    unmatched_projects = []

    for _, folder_path, _ in targets:
        for file in list_source_excel_files(folder_path):
            extracted_name = extract_project_name(file)

            resolved = resolve_project_match(extracted_name, project_config)
            if resolved is not None:
                continue

            if extracted_name not in unmatched_projects:
                unmatched_projects.append(extracted_name)

    if unmatched_projects and project_config:
        print(f"\n发现 {len(unmatched_projects)} 个未匹配项目，打开一次性批量匹配窗口...")
        print("未匹配项目列表：", unmatched_projects)
        project_lookup = get_project_lookup(PROJECT_MATCH_DB) if PROJECT_MATCH_DB is not None else {}
        batch_results = prompt_unmatched_project_mappings(
            unmatched_projects,
            PROJECT_MATCH_DB,
            window_title="采购-批量项目匹配"
        )

        print("\n=== 批量匹配结果 ===")
        for project_name, std_name in batch_results.items():
            if std_name:
                code, _ = project_lookup.get(std_name, ('', ''))
                print(f"  '{project_name}' → [{code}] {std_name}")
            else:
                if PROJECT_MATCH_DB is not None:
                    mark_unmatched(PROJECT_MATCH_DB, project_name)
                print(f"  '{project_name}' → （未选择）")

        if PROJECT_MATCH_DB is not None:
            save_db(PROJECT_MATCH_DB_FILE, PROJECT_MATCH_DB)
            print(
                f"共享匹配库已保存：已匹配 {len(PROJECT_MATCH_DB.get('alias_to_standard', {}))} 项，"
                f"留空 {len(PROJECT_MATCH_DB.get('unmatched_aliases', []))} 项"
            )


def process_simplify_folder(project_config):
    """读取 B_simplify 下的所有原始台账，标准化提取后合并输出到 B_sum。"""
    if not os.path.isdir(SIMPLIFY_DATA_ROOT):
        print(f"simplify 目录不存在，跳过: {SIMPLIFY_DATA_ROOT}")
        return

    # 动态加载 script 2 的标准化提取函数
    try:
        script2 = _load_module_from_path("_procurement_std", PROCUREMENT_STD_MODULE_PATH)
    except Exception as exc:
        print(f"无法加载标准化模块，跳过 simplify 处理: {exc}")
        return

    # 排除本脚本自身生成的汇总文件和宽表文件
    EXCLUDE_NAMES = {"采购原始台账_项目月份覆盖宽表.xlsx", SIMPLIFY_OUTPUT_FILENAME}
    excel_files = []
    for fname in sorted(os.listdir(SIMPLIFY_DATA_ROOT)):
        if fname.startswith("~$") or fname in EXCLUDE_NAMES:
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext in {".xlsx", ".xls", ".xlsm"}:
            excel_files.append(os.path.join(SIMPLIFY_DATA_ROOT, fname))

    if not excel_files:
        print("simplify 下未找到可处理文件，跳过")
        return

    print(f"\n处理 B_simplify 下 {len(excel_files)} 个原始台账...")

    all_dfs = []
    project_totals = defaultdict(float)

    for file_path in excel_files:
        fname = os.path.basename(file_path)
        stem = os.path.splitext(fname)[0]
        raw_name = stem.split("_", 1)[0].strip() if "_" in stem else stem

        hit = match_project(raw_name, PROJECT_MATCH_DB) if PROJECT_MATCH_DB else None
        if hit:
            project_code = hit.get("project_code", "")
            project_name = hit.get("project_name", raw_name)
        else:
            project_code = ""
            project_name = raw_name

        try:
            rows, reason = script2.extract_rows_from_single_excel(file_path, project_code, project_name)
        except Exception as exc:
            print(f"  ✗ {fname}: {exc}")
            continue

        if not rows:
            print(f"  ⚠ {fname}: {reason or '无有效数据'}")
            continue

        df = pd.DataFrame(rows)

        # 处理本文件积累的未知单位（一次弹窗）
        resolved = script2._unit_map_utils.resolve_pending_unit_mappings()
        if resolved and "单位" in df.columns:
            df["单位"] = df["单位"].apply(lambda v: script2.standardize_unit(v, auto_prompt=False))

        # 过滤金额全零行
        amount_df = df[[c for c in ["单价", "不含税金额", "7%增值税", "含税金额"] if c in df.columns]].apply(
            pd.to_numeric, errors="coerce"
        )
        df = df[amount_df.fillna(0).sum(axis=1).ne(0)]
        if df.empty:
            print(f"  ⚠ {fname}: 金额合计均为0，跳过")
            continue

        df = script2.enrich_dates(df)
        all_dfs.append(df)
        print(f"  ✓ {fname}: {len(df)} 行")

    if not all_dfs:
        print("simplify 无有效数据，不输出汇总文件")
        return

    merged_df = pd.concat(all_dfs, ignore_index=True)
    for col in FINAL_OUTPUT_COLUMNS:
        if col not in merged_df.columns:
            merged_df[col] = ""
    merged_df = merged_df[FINAL_OUTPUT_COLUMNS]

    # 构建工作簿
    output_wb = openpyxl.Workbook()
    merged_ws = output_wb.active
    if merged_ws is None:
        raise RuntimeError('无法创建工作表: 合并数据')
    merged_ws.title = "合并数据"
    merged_ws.append(FINAL_OUTPUT_COLUMNS)

    for _, row in merged_df.iterrows():
        row_data = [row[col] for col in FINAL_OUTPUT_COLUMNS]
        merged_ws.append(row_data)
        code = str(row.get("项目编码", "")).strip()
        name = str(row.get("项目名称", "")).strip()
        amt = pd.to_numeric(row.get("含税金额", 0), errors="coerce")
        if pd.notna(amt):
            project_totals[(code, name)] += float(amt)

    summary_ws = output_wb.create_sheet("项目汇总")
    summary_ws.append(["项目编码", "项目名称", "总采购金额"])
    for (code, name), total in sorted(project_totals.items()):
        summary_ws.append([code, name, total])

    procurement_total_df = build_procurement_total_sheet(merged_df)
    total_ws = output_wb.create_sheet("采购总金额")
    total_ws.append(procurement_total_df.columns.tolist())
    for row in procurement_total_df.itertuples(index=False, name=None):
        total_ws.append(list(row))

    apply_unified_workbook_format(output_wb)

    os.makedirs(SUMMARY_OUTPUT_ROOT, exist_ok=True)
    output_path = os.path.join(SUMMARY_OUTPUT_ROOT, SIMPLIFY_OUTPUT_FILENAME)
    output_wb.save(output_path)
    print(f"\n已输出 B_simplify 汇总: {output_path}")
    print(f"共 {len(merged_df)} 行，{len(project_totals)} 个项目")


def main():
    global PROJECT_MATCH_DB

    raw_root = RAW_DATA_ROOT
    standardized_root = DIRECT_STD_INPUT_ROOT
    out_root = DIRECT_OUTPUT_ROOT

    # 直连模式：直接汇总指定标准化目录，并输出到指定简化目录
    if os.path.isdir(standardized_root):
        os.makedirs(out_root, exist_ok=True)
        print("\n直连模式：")
        print(f"- 输入目录: {standardized_root}")
        print(f"- 输出目录: {out_root}")

        output_wb_std, _ = merge_standardized_excels(standardized_root)
        if output_wb_std is None:
            print("直连模式未产生有效汇总数据")
            return

        std_output_path = os.path.join(out_root, DIRECT_OUTPUT_FILENAME)
        apply_unified_workbook_format(output_wb_std)
        output_wb_std.save(std_output_path)
        print(f"已输出（直连汇总）: {std_output_path}")
        return

    # 兜底：若直连目录不存在，回退原有按月份目录处理模式
    standardized_root = STANDARDIZED_DATA_ROOT
    out_root = SUMMARY_OUTPUT_ROOT

    if not os.path.isdir(raw_root):
        print(f"未找到原始数据目录: {raw_root}")
        return

    os.makedirs(out_root, exist_ok=True)

    month_folders = get_month_folders(raw_root)
    if not month_folders:
        print(f"在目录 {raw_root} 下未找到任何子文件夹")
        return

    print("\n检测到以下月份文件夹：")
    for idx, (name, _, ym_key) in enumerate(month_folders, start=1):
        tag = f"{ym_key[0]}-{ym_key[1]:02d}" if ym_key else "未识别年月"
        print(f"  {idx}. {name}  [{tag}]")

    targets = choose_target_folders(month_folders)

    print("\n将处理以下文件夹：")
    for name, folder_path, _ in targets:
        print(f"  - {name}: {folder_path}")

    project_config = load_project_config()
    PROJECT_MATCH_DB = load_or_init_db(PROJECT_MATCH_DB_FILE, project_config)
    print(
        f"已加载共享匹配库：别名 {len(PROJECT_MATCH_DB.get('alias_to_standard', {}))} 个，"
        f"留空 {len(PROJECT_MATCH_DB.get('unmatched_aliases', []))} 个"
    )
    preload_project_matches_for_targets(targets, project_config)

    print("\n单输出模式：")
    print(f"- 仅输出 B_标准化汇总，使用标准化数据: {standardized_root}")

    for name, folder_path, ym_key in targets:
        print("\n" + "#" * 100)
        print(f"开始处理月份目录: {name}")
        print("#" * 100)

        month_tag = f"{ym_key[0]}年{ym_key[1]:02d}月" if ym_key else name
        std_output_filename = f"{month_tag}_采购台账_{STD_MERGED_SUFFIX}"

        # --- B: 标准化汇总（来自B_采购部标准化）---
        std_folder_path = os.path.join(standardized_root, name)
        if not os.path.isdir(std_folder_path):
            print(f"B_标准化汇总跳过：未找到对应目录 {std_folder_path}")
            continue

        output_wb_std, _ = merge_standardized_excels(std_folder_path)
        if output_wb_std is None:
            print(f"B_标准化汇总跳过：{name} 未产生有效汇总数据")
            continue

        std_output_path = os.path.join(out_root, std_output_filename)
        apply_unified_workbook_format(output_wb_std)
        output_wb_std.save(std_output_path)
        print(f"已输出（B_标准化汇总）: {std_output_path}")

    if PROJECT_MATCH_DB is not None:
        save_db(PROJECT_MATCH_DB_FILE, PROJECT_MATCH_DB)
        print(f"共享项目匹配库已保存: {PROJECT_MATCH_DB_FILE}")

    # 处理 B_simplify 下的最新快照文件
    print("\n" + "=" * 80)
    print("处理 B_simplify 最新快照文件...")
    print("=" * 80)
    process_simplify_folder(project_config)

if __name__ == "__main__":
    main()