import pandas as pd
import os
import re
import sys
import json
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# 引入格式化工具
_UTILS_DIR = Path("/Users/wyattsun/Desktop/code/0_项目基本信息")
if str(_UTILS_DIR) not in sys.path and _UTILS_DIR.exists():
    sys.path.insert(0, str(_UTILS_DIR))
from excel_formatting_utils import apply_common_excel_layout  # noqa: E402


# =========================
# 配置区域
# =========================
DATA_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/2_分包合同台账_std"
OUTPUT_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/3_分包合同_sum"
AMOUNT_TOLERANCE = 1.0  # 金额差异容差（泰铢）

OUTPUT_FILENAME = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report/分包合同汇总逐月对比汇总.xlsx"   # ← 固定文件名，无时间戳后缀


def extract_year_month(filename):
    """
    从文件名中提取年月信息
    支持格式：202511, 2025-11, 2025_11 等
    
    Returns:
        str: YYYYMM格式的年月字符串，如 "202511"
    """
    # 匹配 YYYYMM 格式（如 202511）
    match1 = re.search(r'(\d{4})[-_]?(\d{2})', filename)
    if match1:
        return match1.group(1) + match1.group(2)
    
    # 匹配 YYYY-MM 格式
    match2 = re.search(r'(\d{4})-(\d{1,2})', filename)
    if match2:
        year = match2.group(1)
        month = match2.group(2).zfill(2)
        return year + month
    
    return None


def read_contract_file(filepath):
    """
    读取合同文件（支持Excel和JSON）
    
    Returns:
        pd.DataFrame: 合同数据
    """
    ext = os.path.splitext(filepath)[1].lower()
    
    try:
        if ext in ['.xlsx', '.xls']:
            # 读取Excel文件，尝试不同的sheet名称
            df = None
            with pd.ExcelFile(filepath) as xls:
                # 优先读取常见的sheet名称
                for sheet_name in ['Sheet1', '分包合同', '总包合同', xls.sheet_names[0]]:
                    if sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        break
            
            if df is None:
                df = pd.read_excel(filepath)
            
            # 清理空列
            df = df.dropna(axis=1, how='all')
            return df
            
        elif ext == '.json':
            # 读取JSON文件
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 如果是字典格式，尝试提取数据
            if isinstance(data, dict):
                # 尝试常见的键名
                for key in ['分包合同', '总包合同', 'data', 'contracts']:
                    if key in data:
                        data = data[key]
                        break
            
            df = pd.DataFrame(data) if isinstance(data, list) else pd.DataFrame([data])
            return df
        
    except Exception as e:
        print(f"  警告：读取文件失败 {filepath}: {e}")
        return None
    
    return None


def scan_contract_files(data_dir):
    """
    扫描目录下的所有合同文件，按时间排序
    
    Returns:
        list: [(year_month, filepath, dataframe), ...]
    """
    files_data = []
    
    print(f"正在扫描目录: {data_dir}")
    
    if not os.path.exists(data_dir):
        print(f"错误：目录不存在 {data_dir}")
        return files_data
    
    for filename in os.listdir(data_dir):
        filepath = os.path.join(data_dir, filename)
        
        # 跳过目录和隐藏文件
        if not os.path.isfile(filepath) or filename.startswith('.'):
            continue
        
        ext = os.path.splitext(filename)[1].lower()
        
        # 只处理xlsx和json文件（优先xlsx）
        if ext not in ['.xlsx', '.xls', '.json']:
            continue
        
        # 提取年月
        year_month = extract_year_month(filename)
        if not year_month:
            print(f"  跳过（无法提取年月）: {filename}")
            continue
        
        print(f"  发现文件: {filename} -> {year_month}")
        
        # 读取文件
        df = read_contract_file(filepath)
        if df is not None and len(df) > 0:
            files_data.append((year_month, filename, filepath, df))
        else:
            print(f"  警告：文件为空或读取失败: {filename}")
    
    # 按年月排序
    files_data.sort(key=lambda x: x[0])
    
    # 如果同一月份有多个文件，优先选择xlsx
    unique_data = {}
    for year_month, filename, filepath, df in files_data:
        if year_month not in unique_data:
            unique_data[year_month] = (year_month, filename, filepath, df)
        else:
            # 如果当前是xlsx而已存在的不是，则替换
            current_ext = os.path.splitext(filepath)[1].lower()
            existing_ext = os.path.splitext(unique_data[year_month][2])[1].lower()
            if current_ext in ['.xlsx', '.xls'] and existing_ext not in ['.xlsx', '.xls']:
                unique_data[year_month] = (year_month, filename, filepath, df)
    
    result = sorted(unique_data.values(), key=lambda x: x[0])
    
    print(f"\n共找到 {len(result)} 个有效月份的数据")
    for ym, fn, _, _ in result:
        print(f"  {ym}: {fn}")
    
    return result


def compare_two_months(df_prev, df_curr, prev_month, curr_month, amount_col='合同额（泰铢）', contract_id_col='合同编号'):
    """
    对比两个月份的合同数据
    
    Returns:
        dict: {
            'new': DataFrame,      # 新增合同
            'deleted': DataFrame,  # 删除合同
            'changed': DataFrame   # 金额变化的合同
        }
    """
    # 确保合同编号列存在
    if contract_id_col not in df_prev.columns or contract_id_col not in df_curr.columns:
        print(f"  警告：找不到合同编号列 '{contract_id_col}'")
        return {'new': pd.DataFrame(), 'deleted': pd.DataFrame(), 'changed': pd.DataFrame()}
    
    # 获取合同编号集合
    prev_contracts = set(df_prev[contract_id_col].astype(str).dropna().unique())
    curr_contracts = set(df_curr[contract_id_col].astype(str).dropna().unique())
    
    # 新增和删除的合同
    new_contract_ids = curr_contracts - prev_contracts
    deleted_contract_ids = prev_contracts - curr_contracts
    
    new_df = df_curr[df_curr[contract_id_col].astype(str).isin(new_contract_ids)].copy()
    deleted_df = df_prev[df_prev[contract_id_col].astype(str).isin(deleted_contract_ids)].copy()
    
    # 金额变化的合同
    common_contracts = prev_contracts & curr_contracts
    changed_list = []
    
    # 确定金额列名称（尝试多种可能）
    possible_amount_cols = [amount_col, '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）']
    
    amount_col_prev = None
    amount_col_curr = None
    
    for col in possible_amount_cols:
        if col in df_prev.columns:
            amount_col_prev = col
            break
    
    for col in possible_amount_cols:
        if col in df_curr.columns:
            amount_col_curr = col
            break
    
    if amount_col_prev and amount_col_curr:
        for contract_id in common_contracts:
            prev_row = df_prev[df_prev[contract_id_col].astype(str) == contract_id]
            curr_row = df_curr[df_curr[contract_id_col].astype(str) == contract_id]
            
            if len(prev_row) == 0 or len(curr_row) == 0:
                continue
            
            prev_amount = pd.to_numeric(prev_row[amount_col_prev], errors='coerce').sum()
            curr_amount = pd.to_numeric(curr_row[amount_col_curr], errors='coerce').sum()
            
            # 检查金额差异是否超过容差
            if abs(curr_amount - prev_amount) >= AMOUNT_TOLERANCE:
                # 获取项目名称（尝试多种可能的列名）
                project_name = ""
                for col in ['标准项目名称(JSON)', '项目名称', '标准项目名称']:
                    if col in curr_row.columns:
                        project_name = curr_row.iloc[0][col]
                        break
                
                changed_list.append({
                    '合同编号': contract_id,
                    '项目名称': project_name,
                    f'{prev_month}合同额': prev_amount,
                    f'{curr_month}合同额': curr_amount,
                    '变化额': curr_amount - prev_amount,
                    '变化率(%)': ((curr_amount - prev_amount) / prev_amount * 100) if prev_amount != 0 else 0
                })
    
    changed_df = pd.DataFrame(changed_list)
    
    return {
        'new': new_df,
        'deleted': deleted_df,
        'changed': changed_df
    }


def format_month_display(year_month):
    """将YYYYMM格式转换为显示格式 YYYY年MM月"""
    if len(year_month) == 6:
        return f"{year_month[:4]}年{year_month[4:]}月"
    return year_month


def _first_nonempty(row, candidates):
    for col in candidates:
        if col in row.index:
            val = row.get(col)
            if pd.notna(val):
                text = str(val).strip()
                if text and text.lower() != 'nan':
                    return val
    return ''


def _extract_amount_from_row(row):
    for col in ['合同额（泰铢）', '分包合同额（泰铢）', '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）']:
        if col in row.index:
            val = pd.to_numeric(pd.Series([row.get(col)]), errors='coerce').iloc[0]
            if pd.notna(val):
                return float(val)
    return 0.0


def build_simple_change_rows(comparison, prev_month, curr_month):
    """生成精简版合同变化清单（仅新增合同额与金额变化）。"""
    rows = []
    display_period = f"{format_month_display(prev_month)}→{format_month_display(curr_month)}"

    new_df = comparison.get('new', pd.DataFrame())
    if new_df is not None and len(new_df) > 0:
        for _, row in new_df.iterrows():
            contract_id = _first_nonempty(row, ['合同编号'])
            project_name = _first_nonempty(row, ['项目名称', '标准项目名称(JSON)', '标准项目名称'])
            curr_amount = _extract_amount_from_row(row)
            rows.append({
                '对比月份': display_period,
                '变化类型': '新增合同额',
                '合同编号': contract_id,
                '项目名称': project_name,
                '上月合同额': 0.0,
                '本月合同额': curr_amount,
                '变化额': curr_amount,
            })

    changed_df = comparison.get('changed', pd.DataFrame())
    if changed_df is not None and len(changed_df) > 0:
        prev_col = f'{prev_month}合同额'
        curr_col = f'{curr_month}合同额'
        for _, row in changed_df.iterrows():
            prev_amount = pd.to_numeric(pd.Series([row.get(prev_col, 0)]), errors='coerce').fillna(0).iloc[0]
            curr_amount = pd.to_numeric(pd.Series([row.get(curr_col, 0)]), errors='coerce').fillna(0).iloc[0]
            diff_amount = pd.to_numeric(pd.Series([row.get('变化额', curr_amount - prev_amount)]), errors='coerce').fillna(curr_amount - prev_amount).iloc[0]
            rows.append({
                '对比月份': display_period,
                '变化类型': '合同额变化',
                '合同编号': row.get('合同编号', ''),
                '项目名称': row.get('项目名称', ''),
                '上月合同额': float(prev_amount),
                '本月合同额': float(curr_amount),
                '变化额': float(diff_amount),
            })

    if not rows:
        return pd.DataFrame(columns=['对比月份', '变化类型', '合同编号', '项目名称', '上月合同额', '本月合同额', '变化额'])

    return pd.DataFrame(rows)


def build_project_subcontractor_summary(files_data):
    """
    生成项目分包单位统计表（仅项目名称 + 金额不为0的协作体数量）
    
    Returns:
        pd.DataFrame: 包含项目名称、分包数量
    """
    records = []
    
    for year_month, _, _, df in files_data:
        if df is None or len(df) == 0:
            continue
        
        # 识别项目名称列
        project_col = None
        for c in ['项目名称', '标准项目名称(JSON)', '标准项目名称']:
            if c in df.columns:
                project_col = c
                break
        
        # 识别分包单位列
        unit_col = None
        for c in ['协作体名称', '对方公司名称', '对方单位公司名称']:
            if c in df.columns:
                unit_col = c
                break
        
        # 识别合同金额列
        amount_col = None
        for c in ['合同额（泰铢）', '分包合同额（泰铢）', '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）']:
            if c in df.columns:
                amount_col = c
                break
        
        if not project_col or not unit_col or not amount_col:
            continue
        
        # 仅保留：协作体有值 + 金额不为0
        tmp = df[[project_col, unit_col, amount_col]].copy()
        tmp = tmp.dropna(subset=[project_col, unit_col])
        tmp[project_col] = tmp[project_col].astype(str).str.strip()
        tmp[unit_col] = tmp[unit_col].astype(str).str.strip()
        tmp = tmp[(tmp[project_col] != '') & (tmp[project_col].str.lower() != 'nan')]
        tmp = tmp[(tmp[unit_col] != '') & (tmp[unit_col].str.lower() != 'nan')]
        tmp[amount_col] = pd.to_numeric(tmp[amount_col], errors='coerce').fillna(0)
        tmp = tmp[tmp[amount_col] != 0]

        for _, row in tmp.iterrows():
            records.append({
                '项目名称': row[project_col],
                '分包单位': row[unit_col],
            })
    
    if not records:
        return pd.DataFrame(columns=['项目名称', '分包数量'])
    
    result_df = pd.DataFrame(records)

    # 同一项目内对协作体去重计数
    summary = result_df.drop_duplicates(subset=['项目名称', '分包单位'])
    summary = summary.groupby('项目名称')['分包单位'].nunique().reset_index(name='分包数量')
    summary = summary.sort_values(by=['分包数量', '项目名称'], ascending=[False, True])

    return summary


def build_project_month_amount_summary(files_data):
    """
    生成“项目名称 x 年月”的分包签订合同额汇总透视表。

    Returns:
        pd.DataFrame: 行为项目名称，列为年月，值为合同额汇总
    """
    records = []

    for year_month, _, _, df in files_data:
        if df is None or len(df) == 0:
            continue

        # 兼容不同脚本导出的项目名称列
        project_col = None
        for c in ['项目名称', '标准项目名称(JSON)', '标准项目名称']:
            if c in df.columns:
                project_col = c
                break

        if project_col is None:
            continue

        # 兼容不同金额列名
        amount_col = None
        for c in ['合同额（泰铢）', '分包合同额（泰铢）', '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）']:
            if c in df.columns:
                amount_col = c
                break

        if amount_col is None:
            continue

        # 兼容不同项目编码列名
        code_col = None
        for c in ['项目编码', '项目编码(匹配)']:
            if c in df.columns:
                code_col = c
                break

        tmp = df[[project_col, amount_col]].copy()
        if code_col:
            tmp[code_col] = df[code_col]
        
        tmp[project_col] = tmp[project_col].astype(str).str.strip()
        tmp = tmp[(tmp[project_col] != '') & (tmp[project_col].str.lower() != 'nan')]
        tmp[amount_col] = pd.to_numeric(tmp[amount_col], errors='coerce').fillna(0)
        tmp['年月'] = year_month
        tmp = tmp.rename(columns={project_col: '项目名称', amount_col: '分包签订合同额汇总'})
        if code_col:
            tmp = tmp.rename(columns={code_col: '项目编码'})
            records.append(tmp[['项目编码', '项目名称', '年月', '分包签订合同额汇总']])
        else:
            records.append(tmp[['项目名称', '年月', '分包签订合同额汇总']])

    if not records:
        return pd.DataFrame(columns=['项目编码', '项目名称'])

    all_df = pd.concat(records, ignore_index=True)
    
    # 按项目编码和名称分组（如果有编码列）
    if '项目编码' in all_df.columns:
        grouped = all_df.groupby(['项目编码', '项目名称', '年月'], as_index=False)['分包签订合同额汇总'].sum()
        pivot = grouped.pivot_table(index=['项目编码', '项目名称'], columns='年月', values='分包签订合同额汇总', fill_value=0)
    else:
        grouped = all_df.groupby(['项目名称', '年月'], as_index=False)['分包签订合同额汇总'].sum()
        pivot = grouped.pivot(index='项目名称', columns='年月', values='分包签订合同额汇总').fillna(0)

    # 保持年月列按时间顺序，并转换为显示格式
    sorted_cols = sorted(list(pivot.columns))
    display_cols = [format_month_display(col) if len(col) == 6 and col.isdigit() else col for col in sorted_cols]
    pivot.columns = display_cols
    pivot = pivot.reset_index()
    return pivot


def build_project_month_increment_summary(files_data):
    """
    生成按月签订金额增量对比表。
    首月列显示为空（无参照月），后续月列显示与上月的差值（正=新增，负=减少）。
    """
    records = []
    for year_month, _, _, df in files_data:
        if df is None or len(df) == 0:
            continue
        project_col = next((c for c in ['项目名称', '标准项目名称(JSON)', '标准项目名称'] if c in df.columns), None)
        if project_col is None:
            continue
        amount_col = next(
            (c for c in ['合同额（泰铢）', '分包合同额（泰铢）', '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）']
             if c in df.columns), None)
        if amount_col is None:
            continue
        code_col = next((c for c in ['项目编码', '项目编码(匹配)'] if c in df.columns), None)

        tmp = df[[project_col, amount_col]].copy()
        if code_col:
            tmp[code_col] = df[code_col]
        tmp[project_col] = tmp[project_col].astype(str).str.strip()
        tmp = tmp[(tmp[project_col] != '') & (tmp[project_col].str.lower() != 'nan')]
        tmp[amount_col] = pd.to_numeric(tmp[amount_col], errors='coerce').fillna(0)
        tmp['年月'] = year_month
        tmp = tmp.rename(columns={project_col: '项目名称', amount_col: '分包签订合同额汇总'})
        if code_col:
            tmp = tmp.rename(columns={code_col: '项目编码'})
            records.append(tmp[['项目编码', '项目名称', '年月', '分包签订合同额汇总']])
        else:
            records.append(tmp[['项目名称', '年月', '分包签订合同额汇总']])

    if not records:
        return pd.DataFrame(columns=['项目编码', '项目名称'])

    all_df = pd.concat(records, ignore_index=True)

    # 建立金额透视表（保留原始年月列名便于排序计算增量）
    if '项目编码' in all_df.columns:
        grouped = all_df.groupby(['项目编码', '项目名称', '年月'], as_index=False)['分包签订合同额汇总'].sum()
        pivot = grouped.pivot_table(index=['项目编码', '项目名称'], columns='年月',
                                    values='分包签订合同额汇总', fill_value=0)
    else:
        grouped = all_df.groupby(['项目名称', '年月'], as_index=False)['分包签订合同额汇总'].sum()
        pivot = grouped.pivot(index='项目名称', columns='年月',
                              values='分包签订合同额汇总').fillna(0)

    sorted_cols = sorted(list(pivot.columns))

    # 计算增量：首列置空（无参照月），后续列 = 当月 - 上月
    increment = pd.DataFrame(index=pivot.index)
    for i, col in enumerate(sorted_cols):
        display_col = format_month_display(col) if len(col) == 6 and col.isdigit() else col
        if i == 0:
            increment[display_col] = None  # 首月无参照，留空
        else:
            increment[display_col] = pivot[col].values - pivot[sorted_cols[i - 1]].values

    increment = increment.reset_index()
    return increment


def build_project_contract_total_summary(files_data):
    """
    生成项目分包合同汇总表（基于最新月份快照）。

    Returns:
        pd.DataFrame: 列为 项目编码、项目名称、分包合同汇总金额
    """
    if not files_data:
        return pd.DataFrame(columns=['项目编码', '项目名称', '分包合同汇总金额'])

    # 使用最新月份数据，避免跨月重复累计
    latest_month, _, _, df = files_data[-1]
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=['项目编码', '项目名称', '分包合同汇总金额'])

    project_col = next((c for c in ['项目名称', '标准项目名称(JSON)', '标准项目名称'] if c in df.columns), None)
    amount_col = next(
        (c for c in ['分包合同额（泰铢）', '合同额（泰铢）', '合同金额', '合同额', '金额（万泰铢）', '合同额（万泰铢）'] if c in df.columns),
        None
    )
    code_col = next((c for c in ['项目编码', '项目编码(匹配)'] if c in df.columns), None)

    if project_col is None or amount_col is None:
        return pd.DataFrame(columns=['项目编码', '项目名称', '分包合同汇总金额'])

    use_cols = [project_col, amount_col] + ([code_col] if code_col else [])
    tmp = df[use_cols].copy()
    tmp[project_col] = tmp[project_col].astype(str).str.strip()
    tmp = tmp[(tmp[project_col] != '') & (tmp[project_col].str.lower() != 'nan')]
    tmp[amount_col] = pd.to_numeric(tmp[amount_col], errors='coerce').fillna(0)

    if code_col:
        grouped = tmp.groupby([code_col, project_col], as_index=False)[amount_col].sum()
        grouped = grouped.rename(columns={
            code_col: '项目编码',
            project_col: '项目名称',
            amount_col: '分包合同汇总金额'
        })
    else:
        grouped = tmp.groupby([project_col], as_index=False)[amount_col].sum()
        grouped = grouped.rename(columns={
            project_col: '项目名称',
            amount_col: '分包合同汇总金额'
        })
        grouped.insert(0, '项目编码', '')

    grouped = grouped.sort_values(by='分包合同汇总金额', ascending=False).reset_index(drop=True)
    return grouped


def main():
    """主函数 - 支持增量更新 Excel"""
    print("="*80)
    print("合同汇总自动对比系统（支持增量更新版本）")
    print("="*80)

    # 扫描文件
    files_data = scan_contract_files(DATA_DIR)

    if len(files_data) < 2:
        print("\n错误：至少需要2个月份的数据才能进行对比")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_file = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)

    file_exists = os.path.exists(output_file)
    print(f"输出文件: {output_file}  (已存在: {file_exists})")

    # 存储本次新产生的变化记录（仅用于追加）
    all_new_change_parts = []

    print("\n开始逐月对比...")
    for i in range(len(files_data) - 1):
        prev_month, prev_file, _, df_prev = files_data[i]
        curr_month, curr_file, _, df_curr = files_data[i + 1]

        print(f"  对比: {format_month_display(prev_month)} vs {format_month_display(curr_month)}")

        comparison = compare_two_months(df_prev, df_curr, prev_month, curr_month)
        simple_change_df = build_simple_change_rows(comparison, prev_month, curr_month)

        if len(simple_change_df) > 0:
            all_new_change_parts.append(simple_change_df)
            print(f"    ✓ 本次新增变化记录: {len(simple_change_df)} 条")

    # 生成各汇总表（每次都基于最新数据重新生成）
    print("\n生成汇总表...")
    project_contract_total_df = build_project_contract_total_summary(files_data)
    project_month_summary_df = build_project_month_amount_summary(files_data)
    project_month_increment_df = build_project_month_increment_summary(files_data)
    project_subcontractor_df = build_project_subcontractor_summary(files_data)

    # 合并本次变化记录
    new_changes_df = pd.concat(all_new_change_parts, ignore_index=True) if all_new_change_parts else pd.DataFrame(
        columns=['对比月份', '变化类型', '合同编号', '项目名称', '上月合同额', '本月合同额', '变化额']
    )

    # 写入Excel（支持增量更新）
    print("正在写入/更新 Excel...")
    
    writer_kwargs = dict(engine='openpyxl', mode='a' if file_exists else 'w')
    if file_exists:
        writer_kwargs['if_sheet_exists'] = 'replace'
    with pd.ExcelWriter(output_file, **writer_kwargs) as writer:
        # 1. 完整覆盖更新的Sheet（快照类）
        project_contract_total_df.to_excel(writer, sheet_name='项目分包合同汇总', index=False)
        project_month_summary_df.to_excel(writer, sheet_name='项目年月合同额汇总', index=False)
        project_month_increment_df.to_excel(writer, sheet_name='签订金额增量对比', index=False)
        project_subcontractor_df.to_excel(writer, sheet_name='项目分包单位统计', index=False)

        # 2. 合同变化汇总：仅追加新记录（核心需求）
        sheet_name = '合同变化汇总'
        
        if file_exists and sheet_name in writer.book.sheetnames:
            # 已存在Sheet → 追加（不写表头）
            new_changes_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=writer.book[sheet_name].max_row)
            print(f"  → 已追加 {len(new_changes_df)} 条新变化记录到 '合同变化汇总'")
        else:
            # 不存在或首次创建 → 正常写入（带表头）
            new_changes_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  → 创建新Sheet '合同变化汇总'，写入 {len(new_changes_df)} 条记录")

    # 对所有Sheet应用统一格式化
    wb = load_workbook(output_file)
    for ws in wb.worksheets:
        apply_common_excel_layout(ws, numeric_start_col=3)
    wb.save(output_file)

    print(f"\n✓ 更新完成！文件路径: {output_file}")
    print("包含以下Sheet：")
    print("  • 项目分包合同汇总     （最新快照）")
    print("  • 项目年月合同额汇总    （透视表）")
    print("  • 签订金额增量对比      （月度增量）")
    print("  • 项目分包单位统计      （协作体数量）")
    print("  • 合同变化汇总          （历史累计 + 本次追加）")

    print("\n" + "="*80)
    print("合同逐月对比更新完毕！")
    print("="*80)


if __name__ == '__main__':
    main()