# -*- coding: utf-8 -*-
"""
泰国子公司机具租赁费用分析报告生成器 - 完全修复版
已修复：
1. Excel 多表头/标题不在第一行的问题（自动跳过空白行找到真实列名）
2. 月份列格式不统一导致读取为空
3. 空图表 title.text 报错
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

# ==================== 1. 参数设置 ====================
# 自动查找最新的汇总文件
import glob
EXCEL_DIR = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum"
# 兼容多种命名：优先“已汇总”，其次普通“汇总”
excel_files = glob.glob(os.path.join(EXCEL_DIR, "*已汇总*.xlsx"))
if not excel_files:
    excel_files = [
        f for f in glob.glob(os.path.join(EXCEL_DIR, "*汇总*.xlsx"))
        if not os.path.basename(f).startswith("~$")
    ]
if not excel_files:
    raise FileNotFoundError(f"未找到汇总文件，请检查目录: {EXCEL_DIR}")
EXCEL_PATH = max(excel_files, key=os.path.getmtime)  # 获取最新文件
print(f"使用文件: {os.path.basename(EXCEL_PATH)}")

OUTPUT_DIR = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/C_机具租赁分析报告"
os.makedirs(OUTPUT_DIR, exist_ok=True)

HTML_REPORT = os.path.join(OUTPUT_DIR, f"机具租赁费用分析报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
EXCEL_SUMMARY = os.path.join(OUTPUT_DIR, f"机具租赁统计汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# ==================== 2. 智能读取 Excel（读取所有sheet） ====================
print("正在读取 Excel 文件的所有工作表...")

# 读取所有sheet名称
xl_file = pd.ExcelFile(EXCEL_PATH)
print(f"检测到 {len(xl_file.sheet_names)} 个工作表: {xl_file.sheet_names}")

# 读取主数据表（汇总表）
df_raw = pd.read_excel(EXCEL_PATH, sheet_name='汇总', header=None)  # 先不设表头，全读进来

# 找到真正包含列名的那一行（关键词匹配）
header_row: int = 0
keywords = ['区域', '标准机具种类', '型号规格', '月份', '台班费_泰铢']
for idx, row in df_raw.iterrows():
    row_str = ''.join(row.astype(str)).lower()
    if all(k.lower() in row_str for k in ['区域', '标准机具种类']):
        header_row = int(idx) if isinstance(idx, (int, float)) else 0
        print(f"检测到真实列名在第 {header_row+1} 行（从0开始计数）")
        break

# 重新读取，指定正确的 header 行
df = pd.read_excel(EXCEL_PATH, sheet_name='汇总', header=header_row)

# 如果第一列是“来源Sheet”这种合并过来的，可以直接删掉
if '来源Sheet' in df.columns and df['来源Sheet'].isna().all():
    df = df.drop(columns=['来源Sheet'])

print(f"成功读取数据：共 {len(df)} 行，{len(df.columns)} 列")
print("列名预览：", list(df.columns)[:10])

# ==================== 3. 统一列名 + 数据清洗 ====================
df.columns = [str(col).strip() for col in df.columns]

# 智能列名映射（容错处理）
def smart_find_column(df, possible_names, target_name):
    """从可能的列名中找到第一个存在的列，避免冲突"""
    # 如果目标列名已存在，直接返回
    if target_name in df.columns:
        return target_name
    
    # 否则查找相似的列名
    for name in possible_names:
        for col in df.columns:
            if name.lower().replace('_', '').replace(' ', '') in col.lower().replace('_', '').replace(' ', ''):
                # 确保不会映射到已经是标准名的列
                if col not in ['区域', '标准机具种类', '型号规格', '月份', '台班费_泰铢', '本次结算台班', '结算费用_泰铢', '出租单位', '项目名称']:
                    return col
    return None

# 映射关键列（按优先级顺序）
column_groups = {
    '区域': ['区域', '项目区域', '所在区域'],
    '标准机具种类': ['标准机具种类', '机具种类', '种类'],
    '月份': ['月份', '开始月份', '使用月份', '结算月份'],
    '台班费_泰铢': ['台班费_泰铢', '台班费', '单价'],
    '本次结算台班': ['本次结算台班', '结算台班', '台班数量', '数量'],
    '结算费用_泰铢': ['结算费用_泰铢', '本次费用合计', '费用合计'],
    '出租单位': ['出租单位', '供应商', '租赁单位'],
    '项目名称': ['项目名称', '项目']
}

# 特殊处理：如果有'型号规格'和'标准型号规格'，优先使用'标准型号规格'
if '标准型号规格' in df.columns and '型号规格' in df.columns:
    # 删除普通的'型号规格'，保留'标准型号规格'
    df = df.drop(columns=['型号规格'])
    df.rename(columns={'标准型号规格': '型号规格'}, inplace=True)
    print(f"  列名处理: 删除'型号规格'，将'标准型号规格'重命名为'型号规格'")
elif '标准型号规格' in df.columns:
    df.rename(columns={'标准型号规格': '型号规格'}, inplace=True)
    print(f"  列名映射: '标准型号规格' → '型号规格'")
elif '型号规格' not in df.columns:
    # 如果都不存在，查找其他可能的列
    for col in df.columns:
        if '规格' in col or '型号' in col:
            df.rename(columns={col: '型号规格'}, inplace=True)
            print(f"  列名映射: '{col}' → '型号规格'")
            break

rename_dict = {}
for std_name, possible_names in column_groups.items():
    found_col = smart_find_column(df, possible_names, std_name)
    if found_col and found_col != std_name:
        # 检查是否会造成冲突
        if std_name not in df.columns:
            rename_dict[found_col] = std_name
            print(f"  列名映射: '{found_col}' → '{std_name}'")

if rename_dict:
    df.rename(columns=rename_dict, inplace=True)


def _clean_text(val):
    if pd.isna(val):
        return ''
    text = str(val).strip()
    return '' if text.lower() in {'', 'nan', 'none'} else text


def _is_valid_region(val):
    text = _clean_text(val)
    return text not in {'', '未知区域', '未知'}


# 若缺少“区域”列，尝试从“处理报告”回填
if '区域' not in df.columns:
    print("  未找到'区域'列，尝试从'处理报告'回填...")
    df['区域'] = '未知区域'
    try:
        df_report_for_region = pd.read_excel(EXCEL_PATH, sheet_name='处理报告')
        report_cols = set(df_report_for_region.columns)

        if {'区域', '项目编码'}.issubset(report_cols):
            code_region_map = {}
            for _, r in df_report_for_region.iterrows():
                code = _clean_text(r.get('项目编码'))
                region = _clean_text(r.get('区域'))
                if code and _is_valid_region(region):
                    code_region_map.setdefault(code, region)
            if '项目编码' in df.columns and code_region_map:
                df['区域'] = df['项目编码'].astype(str).str.strip().map(code_region_map).fillna(df['区域'])

        if {'区域', '项目名称'}.issubset(report_cols):
            name_region_map = {}
            for _, r in df_report_for_region.iterrows():
                name = _clean_text(r.get('项目名称'))
                region = _clean_text(r.get('区域'))
                if name and _is_valid_region(region):
                    name_region_map.setdefault(name, region)
            if '项目名称' in df.columns and name_region_map:
                missing_mask = (df['区域'].astype(str).str.strip() == '') | (df['区域'] == '未知区域')
                df.loc[missing_mask, '区域'] = (
                    df.loc[missing_mask, '项目名称'].astype(str).str.strip().map(name_region_map).fillna('未知区域')
                )

        if {'区域', 'sheet名称'}.issubset(report_cols):
            sheet_region_map = {}
            for _, r in df_report_for_region.iterrows():
                sheet_name = _clean_text(r.get('sheet名称'))
                region = _clean_text(r.get('区域'))
                if sheet_name and _is_valid_region(region):
                    sheet_region_map.setdefault(sheet_name, region)
            if 'sheet名称' in df.columns and sheet_region_map:
                missing_mask = (df['区域'].astype(str).str.strip() == '') | (df['区域'] == '未知区域')
                df.loc[missing_mask, '区域'] = (
                    df.loc[missing_mask, 'sheet名称'].astype(str).str.strip().map(sheet_region_map).fillna('未知区域')
                )

        known_region_count = int((df['区域'].astype(str).str.strip() != '未知区域').sum())
        print(f"  区域回填完成: 已识别 {known_region_count} 行，未知 {len(df) - known_region_count} 行")
    except Exception as e:
        print(f"  警告: 区域回填失败，将使用默认值'未知区域'。原因: {e}")

# 必填列检查
required = ['区域', '标准机具种类', '型号规格', '月份', '台班费_泰铢', '本次结算台班']
missing = [c for c in required if c not in df.columns]
if missing:
    print("缺少以下必要列：", missing)
    print("当前所有列：", list(df.columns))
    raise KeyError(f"缺少必要列：{missing}")

# 数据类型转换（重点：月份列各种奇葩格式都能识别）
print("\n数据清洗中...")
original_count = len(df)

# 特殊处理月份列（可能是字符串格式如'202411'或'6月份'）
print(f"月份列前5个样本: {df['月份'].head().tolist()}")
print(f"月份列数据类型: {df['月份'].dtype}")

# 智能转换月份（处理多种格式）
def parse_month(month_val):
    """智能解析各种月份格式"""
    if pd.isna(month_val):
        return None
    
    month_str = str(month_val).strip()
    
    # 情况1: "6月份" 或 "10月份" 格式
    if '月份' in month_str or '月' in month_str:
        # 提取数字
        import re
        match = re.search(r'(\d+)', month_str)
        if match:
            month_num = int(match.group(1))
            # 假设是当前年份或上一年
            current_year = datetime.now().year
            try:
                return pd.Timestamp(year=current_year, month=month_num, day=1)
            except:
                return None
    
    # 情况2: "202411" 格式
    if len(month_str) == 6 and month_str.isdigit():
        try:
            return pd.to_datetime(month_str, format='%Y%m')
        except:
            pass
    
    # 情况3: 标准日期格式
    try:
        return pd.to_datetime(month_str)
    except:
        return None

print("  正在智能转换月份格式...")
df['月份'] = df['月份'].apply(parse_month)

invalid_dates = df['月份'].isna().sum()
print(f"  月份转换结果: 成功 {len(df)-invalid_dates} 行, 失败 {invalid_dates} 行")

if invalid_dates == len(df):
    print(f"\n❌ 错误: 所有月份转换失败！")
    print(f"月份列样本: {df['月份'].head(10).tolist()}")
    raise ValueError("无法解析月份格式，请检查数据")

df = df.dropna(subset=['月份'])  # 删掉月份为空的

# 转换为字符串格式（YYYY-MM）以兼容Plotly
df['月份_原始'] = df['月份']  # 保留原始datetime用于排序
df['月份'] = df['月份'].dt.strftime('%Y-%m')

df['本次结算台班'] = pd.to_numeric(df['本次结算台班'], errors='coerce').fillna(0)
df['台班费_泰铢'] = pd.to_numeric(df['台班费_泰铢'], errors='coerce').fillna(0)

# 特殊处理：挖掘机台班费超过20000时视为月租金额，需除以30
print("  正在处理挖掘机月租金额...")
excavator_mask = (df['标准机具种类'].str.contains('挖掘机', na=False)) & (df['台班费_泰铢'] > 20000)
excavator_count = excavator_mask.sum()
if excavator_count > 0:
    print(f"    发现 {excavator_count} 条挖掘机记录台班费>20000，视为月租金额，已自动除以30")
    # 保存原始月租金额到新列（如果需要的话）
    df.loc[excavator_mask, '原始月租金额'] = df.loc[excavator_mask, '台班费_泰铢']
    # 将台班费除以30转换为日租金
    df.loc[excavator_mask, '台班费_泰铢'] = df.loc[excavator_mask, '台班费_泰铢'] / 30
    print(f"    示例：台班费 {df.loc[excavator_mask, '原始月租金额'].iloc[0]:.0f} → {df.loc[excavator_mask, '台班费_泰铢'].iloc[0]:.2f} 泰铢/天")

# 如果结算费用列不存在，尝试计算
if '结算费用_泰铢' in df.columns:
    df['结算费用_泰铢'] = pd.to_numeric(df['结算费用_泰铢'], errors='coerce').fillna(0)
else:
    print("  警告: 未找到'结算费用_泰铢'列，将用'台班费×台班数'计算")
    df['结算费用_泰铢'] = df['台班费_泰铢'] * df['本次结算台班']

# 确保可选列存在
for col in ['出租单位', '项目名称']:
    if col not in df.columns:
        df[col] = '未知'
        print(f"  警告: 未找到'{col}'列，已填充默认值")

df = df.dropna(subset=['区域', '标准机具种类'])

print(f"数据清洗完成: {original_count:,} 行 → {len(df):,} 行 (过滤掉 {original_count-len(df)} 行)")
if invalid_dates > 0:
    print(f"  ⚠️  月份格式无效: {invalid_dates} 行")

# 读取其他工作表
print("\n读取其他分析表...")
try:
    df_report = pd.read_excel(EXCEL_PATH, sheet_name='处理报告')
    print(f"  ✅ 处理报告: {len(df_report)} 行")
except:
    df_report = None
    print("  ⚠️  未找到'处理报告'sheet")

try:
    df_project_kind = pd.read_excel(EXCEL_PATH, sheet_name='项目_机具种类统计')
    print(f"  ✅ 项目_机具种类统计: {len(df_project_kind)} 行")
except:
    df_project_kind = None
    print("  ⚠️  未找到'项目_机具种类统计'sheet")

try:
    df_project_month = pd.read_excel(EXCEL_PATH, sheet_name='项目_月份统计')
    print(f"  ✅ 项目_月份统计: {len(df_project_month)} 行")
except:
    df_project_month = None
    print("  ⚠️  未找到'项目_月份统计'sheet")

try:
    df_region_price = pd.read_excel(EXCEL_PATH, sheet_name='区域_机具_型号_台班费统计')
    print(f"  ✅ 区域_机具_型号_台班费统计: {len(df_region_price)} 行")
except:
    df_region_price = None
    print("  ⚠️  未找到'区域_机具_型号_台班费统计'sheet")

try:
    df_supplier = pd.read_excel(EXCEL_PATH, sheet_name='出租单位统计')
    print(f"  ✅ 出租单位统计: {len(df_supplier)} 行")
except:
    df_supplier = None
    print("  ⚠️  未找到'出租单位统计'sheet")

try:
    df_monthly_report = pd.read_excel(EXCEL_PATH, sheet_name='月度报告')
    print(f"  ✅ 月度报告: {len(df_monthly_report)} 行")
except:
    df_monthly_report = None
    print("  ⚠️  未找到'月度报告'sheet")

# ==================== 4. 处理报告分析（新增） ====================
print("\n" + "="*60)
print("处理报告分析")
print("="*60)

# 初始化变量（避免未定义错误）
total_overtime_errors = 0
total_settlement_errors = 0
total_amount_errors = 0
high_error_projects = pd.DataFrame()
df_report_detail = pd.DataFrame()
df_report_total = None

# 从处理报告提取关键信息
if df_report is not None:
    # 排除合计行
    df_report_detail = df_report[df_report['sheet名称'] != '【合计】'].copy()
    df_report_total = df_report[df_report['sheet名称'] == '【合计】'].iloc[0] if len(df_report[df_report['sheet名称'] == '【合计】']) > 0 else None
    
    # 统计验证错误情况
    if '加班费错误行数' in df_report_detail.columns:
        total_overtime_errors = int(df_report_detail['加班费错误行数'].sum())
        total_settlement_errors = int(df_report_detail['结算费用错误行数'].sum())
        total_amount_errors = int(df_report_detail['本次费用合计错误行数'].sum())
        
        print(f"数据验证结果:")
        print(f"  加班费计算错误: {total_overtime_errors} 行")
        print(f"  结算费用计算错误: {total_settlement_errors} 行")
        print(f"  本次费用合计计算错误: {total_amount_errors} 行")
        print(f"  总错误行数: {total_overtime_errors + total_settlement_errors + total_amount_errors}")
        
        # 找出错误率最高的项目
        if '加班费错误率(%)' in df_report_detail.columns:
            high_error_projects = df_report_detail[
                (df_report_detail['加班费错误率(%)'] > 10) | 
                (df_report_detail['结算费用错误率(%)'] > 10) | 
                (df_report_detail['本次费用合计错误率(%)'] > 10)
            ]
            if len(high_error_projects) > 0:
                print(f"\n⚠️  高错误率项目 (>10%): {len(high_error_projects)} 个")
                for idx, row in high_error_projects.head(5).iterrows():
                    print(f"  {row['项目名称']}: 加班费{row['加班费错误率(%)']}%, 结算费用{row['结算费用错误率(%)']}%, 费用合计{row['本次费用合计错误率(%)']}%")

# ==================== 5. 数据质量报告 ====================
print("\n" + "="*60)
print("数据质量检查报告")
print("="*60)

# 检查异常数据
zero_price = df[df['台班费_泰铢'] == 0]
zero_quantity = df[df['本次结算台班'] == 0]
high_price = df[df['台班费_泰铢'] > 50000]  # 假设超过5万泰铢为异常

print(f"总记录数: {len(df):,}")
print(f"  台班费为0: {len(zero_price):,} 行")
print(f"  台班数量为0: {len(zero_quantity):,} 行")
print(f"  台班费>50000: {len(high_price):,} 行（可能异常）")
print(f"涉及区域: {df['区域'].nunique()} 个 - {', '.join(df['区域'].unique())}")
print(f"涉及机具种类: {df['标准机具种类'].nunique()} 种")
print(f"时间跨度: {df['月份'].min()} 至 {df['月份'].max()}")
print(f"供应商数量: {df['出租单位'].nunique()} 家")

# ==================== 6. 统计分析（增强版） ====================
print("\n正在生成统计分析...")

WAN_DIVISOR = 10000.0

def to_wan(value):
    """将金额从泰铢转换为万泰铢"""
    try:
        return float(value) / WAN_DIVISOR
    except:
        return 0.0

# 使用年月列进行分析（如果存在）
if '年月' in df.columns:
    print("  使用'年月'列进行时间分析")
    time_col = '年月'
    # 清理空值
    df_time = df[df[time_col].notna() & (df[time_col] != '')].copy()
else:
    print("  使用'月份'列进行时间分析")
    time_col = '月份'
    df_time = df.copy()

monthly_summary = df_time.groupby(['区域', '标准机具种类', time_col], as_index=False).agg({
    '本次结算台班': 'sum',
    '结算费用_泰铢': 'sum',
    '台班费_泰铢': 'mean',
    '型号规格': 'nunique',
    '出租单位': 'nunique'
}).rename(columns={
    '本次结算台班': '台班数量',
    '结算费用_泰铢': '结算费用_泰铢',
    '台班费_泰铢': '平均台班费_泰铢',
    '型号规格': '规格种类数',
    '出租单位': '供应商数'
})
monthly_summary['结算费用_万泰铢'] = monthly_summary['结算费用_泰铢'].apply(to_wan)

price_analysis = df.groupby(['区域', '标准机具种类', '型号规格', '台班费_泰铢'], as_index=False).agg({
    '本次结算台班': 'sum',
    '结算费用_泰铢': 'sum',
    '出租单位': lambda x: ', '.join(x.dropna().unique()[:3]),  # 显示前3个供应商
    '项目名称': 'nunique'
}).rename(columns={
    '本次结算台班': '台班数量',
    '结算费用_泰铢': '总费用_泰铢',
    '出租单位': '供应商列表',
    '项目名称': '项目数'
})

# 多价格分析（增强版）
multi_price_groups = df.groupby(['区域', '标准机具种类', '型号规格']).agg({
    '台班费_泰铢': ['min', 'max', 'mean', 'nunique'],
    '本次结算台班': 'sum',
    '结算费用_泰铢': 'sum'
}).reset_index()

multi_price_groups.columns = ['区域', '标准机具种类', '型号规格', '最低台班费', '最高台班费', '平均台班费', '价格种类数', '总台班量', '总费用']
multi_price_groups['价格差异率(%)'] = ((multi_price_groups['最高台班费'] - multi_price_groups['最低台班费']) / multi_price_groups['最低台班费'] * 100).round(2)
multi_price_groups['潜在节省'] = (multi_price_groups['最高台班费'] - multi_price_groups['最低台班费']) * multi_price_groups['总台班量']
multi_price_groups['潜在节省_万泰铢'] = multi_price_groups['潜在节省'].apply(to_wan)

multi_price_models = multi_price_groups[multi_price_groups['价格种类数'] > 1].sort_values('潜在节省', ascending=False)
multi_price_count = len(multi_price_models)

# 供应商分析（使用读取的数据）
if df_supplier is not None:
    print("\n供应商分析:")
    print(f"  总供应商数: {len(df_supplier)}")
    if len(df_supplier) > 0:
        top_suppliers = df_supplier.nlargest(5, '费用合计(泰铢)')
        print("  Top 5 供应商:")
        for idx, row in top_suppliers.iterrows():
            print(f"    {row['出租单位']}: {row['费用合计(泰铢)']:,.0f} 泰铢 ({row['使用次数']} 次)")
else:
    top_suppliers = df.groupby('出租单位')['结算费用_泰铢'].sum().nlargest(5)

if multi_price_count > 0:
    print(f"\n⚠️  价格异常警告: 发现 {multi_price_count} 组机具存在多种台班费")
    print("\n前5个最大差异:")
    for idx, row in multi_price_models.head(5).iterrows():
        print(f"  {row['标准机具种类']} - {row['型号规格']}: ")
        print(f"    价格范围: {row['最低台班费']:.0f} ~ {row['最高台班费']:.0f} 泰铢 (差异 {row['价格差异率(%)']}%)")
        print(f"    如统一为最低价，可节省: {row['潜在节省']:,.0f} 泰铢")

# ==================== 6. 绘图（修复 title.text 问题） ====================
print("\n正在生成可视化图表...")

def add_title(fig, title):
    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor='center', font=dict(size=18, family='Microsoft YaHei')),
        font=dict(family='Microsoft YaHei')
    )
    return fig

figs = []
fig_titles = []

# 图1：旭日图 - 整体结构
try:
    fig1 = add_title(px.sunburst(monthly_summary, path=['区域', '标准机具种类', '月份'],
                                values='台班数量', color='结算费用_万泰铢',
                                color_continuous_scale='OrRd'),
                     '图1: 各区域 → 标准机具种类 → 月份 台班使用量（颜色越深费用越高，单位：万泰铢）')
    fig1.update_coloraxes(colorbar_title_text='万泰铢')
    figs.append(fig1)
    fig_titles.append('图1: 整体使用结构')
except Exception as e:
    print(f"  警告: 图1生成失败 - {e}")

# 图2：平均台班费对比
try:
    fig2 = add_title(px.bar(monthly_summary.groupby(['区域', '标准机具种类'], as_index=False)['平均台班费_泰铢'].mean(),
                            x='区域', y='平均台班费_泰铢', color='标准机具种类', barmode='group'),
                     '图2: 各区域各类机具平均台班费对比')
    figs.append(fig2)
    fig_titles.append('图2: 平均台班费对比')
except Exception as e:
    print(f"  警告: 图2生成失败 - {e}")

# 图3：价格异常散点图
try:
    if multi_price_count > 0:
        bubble_data = multi_price_models.head(50)  # 只显示前50个
        bubble_data = bubble_data.copy()
        fig3 = px.scatter(bubble_data, x='型号规格', y='平均台班费', size='总台班量',
                          color='价格差异率(%)', color_continuous_scale='RdYlGn_r',
                          hover_data=['最低台班费', '最高台班费', '潜在节省_万泰铢'],
                          height=600)
        fig3 = add_title(fig3, f'图3: 【重点警示】价格异常机具 (共{multi_price_count}组，显示前50)')
    else:
        fig3 = go.Figure().add_annotation(text="✓ 所有相同型号规格价格统一", x=0.5, y=0.5,
                                          font=dict(size=24, color='green'), showarrow=False)
        fig3 = add_title(fig3, '图3: 价格一致性检查')
    figs.append(fig3)
    fig_titles.append('图3: 价格异常分析')
except Exception as e:
    print(f"  警告: 图3生成失败 - {e}")

# 图4：月度使用趋势
try:
    trend_data = monthly_summary.copy()
    # 按月份排序
    trend_data['月份_排序'] = pd.to_datetime(trend_data['月份'])
    trend_data = trend_data.sort_values('月份_排序')
    
    fig4 = px.line(trend_data, x='月份', y='台班数量',
                   color='标准机具种类', facet_col='区域', facet_col_wrap=2, height=700)
    fig4 = add_title(fig4, '图4: 各区域机具种类月度使用趋势')
    figs.append(fig4)
    fig_titles.append('图4: 月度使用趋势')
except Exception as e:
    print(f"  警告: 图4生成失败 - {e}")

# ==================== 6.8 月度报告图表生成 ====================
print("\n正在处理月度报告数据...")

monthly_report_cards_html = ''
monthly_fig1_html = ''
monthly_fig2_html = ''

if df_monthly_report is not None and len(df_monthly_report) > 0:
    try:
        project_id_cols_mr = [c for c in ['项目编码', '项目名称'] if c in df_monthly_report.columns]
        month_cols_mr = [c for c in df_monthly_report.columns if c not in project_id_cols_mr]

        if month_cols_mr:
            import re as _re

            def _month_sort_key_mr(m):
                s = str(m).strip()
                mt = _re.match(r'^(\d{4})-(\d{1,2})$', s)
                if mt:
                    return (0, int(mt.group(1)) * 100 + int(mt.group(2)))
                return (1, month_cols_mr.index(m))

            month_cols_order = sorted(month_cols_mr, key=_month_sort_key_mr)

            # 严格按「月度报告」sheet页原始矩阵计算，不做筛零和重排
            matrix_source = df_monthly_report.copy()
            for mcol in month_cols_order:
                matrix_source[mcol] = pd.to_numeric(matrix_source[mcol], errors='coerce').fillna(0.0)

            if '项目编码' in project_id_cols_mr and '项目名称' in project_id_cols_mr:
                matrix_source['项目编码_mr'] = matrix_source['项目编码'].fillna('').astype(str).str.strip()
                matrix_source['项目名称_mr'] = matrix_source['项目名称'].fillna('').astype(str).str.strip()
                matrix_source['项目标识_mr'] = matrix_source.apply(
                    lambda r: (
                        f"{r['项目编码_mr']} | {r['项目名称_mr']}"
                        if r['项目编码_mr']
                        else (r['项目名称_mr'] if r['项目名称_mr'] else '未匹配项目')
                    ),
                    axis=1
                )
            elif '项目名称' in project_id_cols_mr:
                matrix_source['项目标识_mr'] = matrix_source['项目名称'].fillna('').astype(str).str.strip()
            elif '项目编码' in project_id_cols_mr:
                matrix_source['项目标识_mr'] = matrix_source['项目编码'].fillna('').astype(str).str.strip()
            else:
                matrix_source['项目标识_mr'] = [f"项目{i+1}" for i in range(len(matrix_source))]

            matrix_source['项目标识_mr'] = matrix_source['项目标识_mr'].replace('', '未匹配项目')
            matrix_df = matrix_source[['项目标识_mr'] + month_cols_order].copy()

            monthly_totals_mr = pd.DataFrame({
                '月份_mr': month_cols_order,
                '费用合计': [matrix_df[m].sum() for m in month_cols_order]
            })
            monthly_totals_mr['费用合计_万泰铢'] = monthly_totals_mr['费用合计'].apply(to_wan)

            # 生成月度卡片 HTML
            cards_html_parts = []
            for _, mrow in monthly_totals_mr.iterrows():
                active_n = int((matrix_df[mrow['月份_mr']] > 0).sum())
                cards_html_parts.append(
                    '<div class="metric-card" style="border-top:3px solid #667eea;">'
                    + f'<div class="metric-value" style="font-size:20px;color:#2c3e50;">{mrow["费用合计_万泰铢"]:,.2f}</div>'
                    + f'<div class="metric-label" style="font-weight:bold;">{mrow["月份_mr"]}</div>'
                    + f'<div class="metric-small">{active_n} 个活跃项目 | 万泰铢</div>'
                    + '</div>'
                )
            monthly_report_cards_html = '\n'.join(cards_html_parts)

            # 图表1：月度总费用柱状图 + 趋势点线
            fig_mr1 = go.Figure()
            fig_mr1.add_trace(go.Bar(
                x=list(monthly_totals_mr['月份_mr']),
                y=list(monthly_totals_mr['费用合计_万泰铢']),
                name='月度总费用(万泰铢)',
                marker_color='#667eea',
                opacity=0.85,
                text=[f'{v:,.2f}' for v in monthly_totals_mr['费用合计_万泰铢']],
                textposition='outside',
                textfont=dict(size=11),
                hovertemplate='月份: %{x}<br>费用: %{y:,.2f} 万泰铢<extra></extra>'
            ))
            fig_mr1.add_trace(go.Scatter(
                x=list(monthly_totals_mr['月份_mr']),
                y=list(monthly_totals_mr['费用合计_万泰铢']),
                name='趋势',
                mode='lines+markers',
                line=dict(color='#764ba2', width=2, dash='dot'),
                hovertemplate='月份: %{x}<br>费用: %{y:,.2f} 万泰铢<extra></extra>',
                marker=dict(size=8, color='#764ba2'),
            ))
            fig_mr1.update_layout(
                title=dict(text='月度机具费用总览', x=0.5, xanchor='center',
                           font=dict(size=16, family='Microsoft YaHei')),
                font=dict(family='Microsoft YaHei'),
                xaxis_title='月份', yaxis_title='费用合计 (万泰铢)',
                legend=dict(orientation='h', y=1.1),
                yaxis=dict(tickformat=','),
                height=440,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=70, b=60),
            )
            fig_mr1.update_yaxes(gridcolor='#e8e8e8')
            monthly_fig1_html = fig_mr1.to_html(full_html=False, include_plotlyjs=False)

            # 图表2：各项目月度机具费变化折线图（同图展示，按最新支付时间排序）
            def _to_float_mr(v):
                try:
                    return float(v)
                except:
                    return 0.0

            def _latest_paid_idx(row):
                for idx in range(len(month_cols_order) - 1, -1, -1):
                    if _to_float_mr(row[month_cols_order[idx]]) > 0:
                        return idx
                return -1

            project_trend_df = matrix_df.copy()
            project_trend_df['_latest_paid_idx'] = project_trend_df.apply(_latest_paid_idx, axis=1)
            project_trend_df['_latest_paid_amount'] = project_trend_df.apply(
                lambda r: _to_float_mr(r[month_cols_order[r['_latest_paid_idx']]]) if r['_latest_paid_idx'] >= 0 else 0.0,
                axis=1
            )
            project_trend_df['_total_paid'] = project_trend_df[month_cols_order].apply(
                lambda r: r.apply(_to_float_mr).sum(),
                axis=1
            )
            project_trend_df = project_trend_df.sort_values(
                by=['_latest_paid_idx', '_latest_paid_amount', '_total_paid', '项目标识_mr'],
                ascending=[False, False, False, True]
            )

            fig_mr2 = go.Figure()
            for _, prow in project_trend_df.iterrows():
                project_name = str(prow['项目标识_mr'])
                y_values = [to_wan(_to_float_mr(prow[m])) for m in month_cols_order]
                fig_mr2.add_trace(go.Scatter(
                    x=month_cols_order,
                    y=y_values,
                    mode='lines+markers',
                    name=project_name,
                    line=dict(width=1.8),
                    marker=dict(size=5),
                    hovertemplate='项目: %{fullData.name}<br>月份: %{x}<br>费用: %{y:,.2f} 万泰铢<extra></extra>'
                ))

            fig_mr2.update_layout(
                title=dict(text='各项目月度机具费变化（按最新支付时间排序）', x=0.5, xanchor='center',
                           font=dict(size=16, family='Microsoft YaHei')),
                font=dict(family='Microsoft YaHei'),
                xaxis_title='月份',
                yaxis_title='费用合计 (万泰铢)',
                yaxis=dict(tickformat=','),
                height=max(620, 360 + 6 * len(project_trend_df)),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=80, b=60, l=80, r=320),
                legend=dict(
                    title='项目（按最新支付时间排序）',
                    orientation='v',
                    x=1.02,
                    y=1,
                    traceorder='normal'
                )
            )
            fig_mr2.update_yaxes(gridcolor='#e8e8e8')

            monthly_fig2_html = fig_mr2.to_html(full_html=False, include_plotlyjs=False)

            print(f"  ✅ 月度报告图表生成成功：{len(month_cols_order)} 个月份，{len(matrix_df)} 个项目")
    except Exception as e:
        print(f"  ⚠️  月度报告图表生成失败: {e}")
        import traceback; traceback.print_exc()

# ==================== 7. 生成 HTML 报告 ====================
print("\n正在生成HTML报告...")

total_cost = df['结算费用_泰铢'].sum()
total_cost_wan = to_wan(total_cost)
potential_savings = multi_price_models['潜在节省'].sum() if multi_price_count > 0 else 0
potential_savings_wan = to_wan(potential_savings)

# 准备HTML表格内容（避免f-string中的反斜杠问题）
newline = '\n'
high_error_table_rows = newline.join([f"<tr><td>{row['项目名称']}</td><td>{row['数据行数']}</td><td><span class='error-badge'>{row['加班费错误率(%)']}%</span></td><td><span class='error-badge'>{row['结算费用错误率(%)']}%</span></td><td><span class='error-badge'>{row['本次费用合计错误率(%)']}%</span></td></tr>" for idx, row in high_error_projects.head(10).iterrows()]) if len(high_error_projects) > 0 else ''

supplier_table_rows = ''
if df_supplier is not None and len(df_supplier) > 0:
    supplier_table_rows = newline.join([f"<tr><td>{i+1}</td><td>{row['出租单位']}</td><td>{to_wan(row['费用合计(泰铢)']):,.2f}</td><td>{row['使用次数']}</td><td>{row['机具种类数']}</td><td>{row['费用合计(泰铢)']/total_cost*100:.2f}%</td></tr>" for i, (idx, row) in enumerate(df_supplier.nlargest(10, '费用合计(泰铢)').iterrows())])

project_table_rows = ''
if df_report is not None and len(df_report_detail) > 0:
    project_table_rows = newline.join([
        f"<tr><td>{row.get('项目编码', '')}</td><td>{row['项目名称']}</td><td>{row['区域']}</td><td>{row['数据行数']}</td><td>{row['机具种类数']}</td><td>{to_wan(row['费用合计(泰铢)']):,.2f}</td><td>{row['费用合计(泰铢)']/total_cost*100:.2f}%</td></tr>"
        for idx, row in df_report_detail[df_report_detail['状态'] == '成功']
        .sort_values(['费用合计(泰铢)', '项目名称'], ascending=[False, True])
        .iterrows()
    ])

html_content = f"""
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>泰国机具租赁分析报告</title>
<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
<style>
    body {{ font-family: "Microsoft YaHei",Arial,sans-serif; margin:20px 40px; background:#f5f5f5; }}
    h1 {{ text-align:center; color:#2c3e50; padding:20px; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; border-radius:10px; }}
    h2 {{ color:#2c3e50; border-left:4px solid #667eea; padding-left:15px; margin-top:40px; }}
    h3 {{ color:#34495e; margin-top:25px; }}
    .info {{ background:white; padding:25px; border-radius:10px; box-shadow:0 2px 15px rgba(0,0,0,0.08); margin:20px 0; }}
    .highlight {{ background:#fff3cd; padding:15px; border-left:4px solid #ffc107; margin:15px 0; border-radius:5px; }}
    .success {{ background:#d4edda; border-left:4px solid #28a745; }}
    .warning {{ background:#f8d7da; border-left:4px solid #dc3545; }}
    .danger {{ background:#f8d7da; border-left:4px solid #dc3545; }}
    .metrics {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:15px; margin:20px 0; }}
    .metric-card {{ background:white; padding:20px; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.05); text-align:center; }}
    .metric-value {{ font-size:32px; font-weight:bold; color:#667eea; }}
    .metric-label {{ color:#666; margin-top:5px; }}
    .metric-small {{ font-size:12px; color:#999; }}
    table {{ width:100%; border-collapse:collapse; margin:20px 0; background:white; box-shadow:0 2px 8px rgba(0,0,0,0.05); }}
    th {{ background:#667eea; color:white; padding:12px; text-align:left; }}
    td {{ padding:10px; border-bottom:1px solid #ddd; }}
    tr:hover {{ background:#f5f5f5; }}
    .section {{ background:white; padding:25px; border-radius:10px; box-shadow:0 2px 15px rgba(0,0,0,0.08); margin:30px 0; }}
    .error-badge {{ background:#dc3545; color:white; padding:2px 8px; border-radius:3px; font-size:12px; }}
    .ok-badge {{ background:#28a745; color:white; padding:2px 8px; border-radius:3px; font-size:12px; }}
</style>
</head><body>
<h1>🏗️ 泰国子公司机具租赁费用分析报告</h1>

<div class="info">
<h3>📊 基础信息</h3>
<p><strong>生成时间：</strong>{datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}</p>
<p><strong>分析周期：</strong>{df['月份'].min()} 至 {df['月份'].max()}</p>
</div>

<div class="metrics">
    <div class="metric-card">
        <div class="metric-value">{len(df):,}</div>
        <div class="metric-label">总记录数</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{df['区域'].nunique()}</div>
        <div class="metric-label">涉及区域</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{df['标准机具种类'].nunique()}</div>
        <div class="metric-label">机具种类</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{total_cost_wan:,.2f}</div>
        <div class="metric-label">总费用(万泰铢)</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{df['出租单位'].nunique()}</div>
        <div class="metric-label">供应商数量</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{df['项目名称'].nunique()}</div>
        <div class="metric-label">涉及项目</div>
    </div>
</div>

<div class="section">
<h2>📋 数据验证结果</h2>
{'<div class="highlight danger"><h3>⚠️ 发现数据计算错误</h3>' if total_overtime_errors + total_settlement_errors + total_amount_errors > 0 else '<div class="highlight success"><h3>✅ 数据验证通过</h3>'}
<div class="metrics">
    <div class="metric-card">
        <div class="metric-value" style="color:{'#dc3545' if total_overtime_errors > 0 else '#28a745'}">{total_overtime_errors}</div>
        <div class="metric-label">加班费计算错误</div>
        {f'<div class="metric-small">差异金额: {to_wan(df_report_total["加班费差异金额(泰铢)"]):,.2f} 万泰铢</div>' if df_report_total is not None and '加班费差异金额(泰铢)' in df_report_total.index else ''}
    </div>
    <div class="metric-card">
        <div class="metric-value" style="color:{'#dc3545' if total_settlement_errors > 0 else '#28a745'}">{total_settlement_errors}</div>
        <div class="metric-label">结算费用计算错误</div>
        {f'<div class="metric-small">差异金额: {to_wan(df_report_total["结算费用差异金额(泰铢)"]):,.2f} 万泰铢</div>' if df_report_total is not None and '结算费用差异金额(泰铢)' in df_report_total.index else ''}
    </div>
    <div class="metric-card">
        <div class="metric-value" style="color:{'#dc3545' if total_amount_errors > 0 else '#28a745'}">{total_amount_errors}</div>
        <div class="metric-label">费用合计计算错误</div>
        {f'<div class="metric-small">差异金额: {to_wan(df_report_total["本次费用合计差异金额(泰铢)"]):,.2f} 万泰铢</div>' if df_report_total is not None and '本次费用合计差异金额(泰铢)' in df_report_total.index else ''}
    </div>
</div>
{f'<p><strong>建议：</strong>请核查标注为红色/粉色/金色/浅绿色的行，修正计算错误。</p>' if total_overtime_errors + total_settlement_errors + total_amount_errors > 0 else '<p>所有数据验证通过，计算准确无误。</p>'}
</div>
</div>

{f'''<div class="section">
<h2>🏢 供应商分析</h2>
<p><strong>总供应商数量：</strong>{df['出租单位'].nunique()} 家</p>
<h3>Top 10 供应商（按费用排序）</h3>
<table>
<tr><th>排名</th><th>供应商名称</th><th>费用合计(万泰铢)</th><th>使用次数</th><th>机具种类数</th><th>占比</th></tr>
{supplier_table_rows}
</table>
</div>''' if df_supplier is not None and len(df_supplier) > 0 else ''}

{f'''<div class="section">
<h2>📊 项目统计分析</h2>
<h3>各项目费用汇总</h3>
<table>
<tr><th>项目编码</th><th>项目名称</th><th>区域</th><th>数据行数</th><th>机具种类数</th><th>费用合计(万泰铢)</th><th>占比</th></tr>
{project_table_rows}
</table>
</div>''' if df_report is not None and len(df_report_detail) > 0 else ''}
"""

# 注入月度报告 HTML 段落（卡片 + 图表）
if monthly_fig1_html:
    _chart_wrap = ('<div style="background:white;border-radius:10px;padding:20px;'
                   'box-shadow:0 2px 8px rgba(0,0,0,0.05);margin-top:20px;">{}</div>')
    html_content += (
        '<div class="section">'
        '<h2>📅 月度机具费用报告</h2>'
        + _chart_wrap.format(monthly_fig1_html)
        + (_chart_wrap.format(monthly_fig2_html) if monthly_fig2_html else '')
        + '</div>'
    )

html_content += """
<div class="info" style="margin-top:50px; text-align:center; color:#666;">
<p>📧 如有疑问请联系数据分析部门</p>
<p style="font-size:12px;">本报告由Python自动生成 | 数据仅供内部参考</p>
</div>
</body></html>
"""

with open(HTML_REPORT, "w", encoding="utf-8") as f:
    f.write(html_content)

# ==================== 8. 导出 Excel ====================
print("\n正在导出Excel汇总...")

with pd.ExcelWriter(EXCEL_SUMMARY, engine='openpyxl') as writer:
    # 汇总统计
    monthly_summary.to_excel(writer, '月度统计', index=False)
    
    # 价格明细
    price_analysis.to_excel(writer, '价格明细', index=False)
    
    # 价格异常清单（如果有）
    if multi_price_count > 0:
        multi_price_models.to_excel(writer, '价格异常_重点关注', index=False)
    
    # 处理报告（如果有）
    if df_report is not None:
        df_report.to_excel(writer, '处理报告', index=False)
    
    # 项目机具统计（如果有）
    if df_project_kind is not None:
        df_project_kind.to_excel(writer, '项目_机具种类统计', index=False)
    
    if df_project_month is not None:
        df_project_month.to_excel(writer, '项目_月份统计', index=False)
    
    # 区域机具价格统计（如果有）
    if df_region_price is not None:
        df_region_price.to_excel(writer, '区域_机具_型号_台班费统计', index=False)
    
    # 供应商统计（如果有）
    if df_supplier is not None:
        df_supplier.to_excel(writer, '供应商统计', index=False)

    # 月度报告（如果有）
    if df_monthly_report is not None:
        df_monthly_report.to_excel(writer, '月度报告', index=False)

    # 高错误率项目（如果有）
    if len(high_error_projects) > 0:
        high_error_projects.to_excel(writer, '高错误率项目', index=False)
    
    # 数据质量报告
    quality_report = pd.DataFrame({
        '检查项': ['总记录数', '台班费为0', '台班数量为0', '台班费>50000', '区域数', '机具种类数', '供应商数', 
                  '加班费错误', '结算费用错误', '费用合计错误'],
        '数量': [len(df), len(zero_price), len(zero_quantity), len(high_price), 
                df['区域'].nunique(), df['标准机具种类'].nunique(), df['出租单位'].nunique(),
                total_overtime_errors, total_settlement_errors, total_amount_errors]
    })
    quality_report.to_excel(writer, '数据质量报告', index=False)
    
    # 原始数据
    df_export = df.drop(columns=['月份_原始'], errors='ignore')
    df_export.to_excel(writer, '原始数据', index=False)

print("\n" + "="*60)
print("全部成功！")
print(f"HTML 报告：{HTML_REPORT}")
print(f"Excel 汇总：{EXCEL_SUMMARY}")

# ==================== 9. 清理旧文件 ====================
print("\n正在清理旧文件...")

import glob

# 删除旧的 HTML 报告（保留最新的）
html_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "机具租赁费用分析报告_*.html")))
if len(html_files) > 1:
    for old_file in html_files[:-1]:  # 保留最后一个（最新的）
        try:
            os.remove(old_file)
            print(f"  ✓ 已删除旧报告：{os.path.basename(old_file)}")
        except Exception as e:
            print(f"  ✗ 无法删除 {os.path.basename(old_file)}: {e}")

# 删除旧的 Excel 汇总（保留最新的）
excel_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "机具租赁统计汇总_*.xlsx")))
if len(excel_files) > 1:
    for old_file in excel_files[:-1]:  # 保留最后一个（最新的）
        try:
            os.remove(old_file)
            print(f"  ✓ 已删除旧汇总：{os.path.basename(old_file)}")
        except Exception as e:
            print(f"  ✗ 无法删除 {os.path.basename(old_file)}: {e}")

print("="*60)