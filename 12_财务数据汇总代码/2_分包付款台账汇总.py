import os
import pandas as pd
import subprocess
import sys
import warnings
import openpyxl
from openpyxl.utils import get_column_letter
import json
import re
import importlib.util

# 忽略UserWarning，通常来自openpyxl或pandas
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# ==============================================================================
# 1. 定义和配置
# ==============================================================================

# 固定输入输出根路径（批量模式：输入根目录下每个子文件夹视为一个月份）
INPUT_FOLDER = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_分包对账与付款/A_分包对账与付款_原始数据"
OUTPUT_FOLDER = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_分包对账与付款/B_分包对账与付款_汇总数据"

# 后置分析脚本：在分包台账汇总完成后自动执行
POST_ANALYSIS_SCRIPT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "3_分包付款汇总分析_按日期与月增量.py"
)

# 项目名称匹配模块与数据库（固定路径）
PROJECT_MATCHER_MODULE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    '0_项目基本信息',
    'project_name_matcher.py'
)
PROJECT_MATCHER_DB_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/project_name_match_db.json"
EXCEL_FORMATTING_UTILS_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/excel_formatting_utils.py"

# 定义要提取的列名（基于用户要求更新，包含人民币换算列和汇率）
columns_to_extract = [
    "付款单编号",
    "已完成产值（泰铢）",
    "合同应付款",
    "含税实付金额（泰铢）",
    "含税实付金额（人民币）", # 新增
    "含税实付金额（换算人民币）", # 新增
    "国内代付金额（人民币）", # 新增：从泰铢列提取的人民币金额
    "换算泰铢金额", # 新增：人民币换算回泰铢的公式列
    "付款汇总（泰铢）", # 新增：含税实付金额（泰铢）+ 换算泰铢金额
    "扣除借款（泰铢）",
    "扣除借款（人民币）", # 新增
    "扣除预扣税（泰铢）",
    "扣除预扣税（人民币）", # 新增
    "公司实际付款（财务数据）（泰铢）",
    "公司实际付款（财务数据）（人民币）", # 新增
    "财务付款日期",
    "汇率" # 假定存在汇率列用于计算
]

# 汇率列可能的名称（模糊匹配用）
exchange_rate_variations = ["汇率", "汇率（THB/CNY）", "THB/CNY", "泰铢汇率", "exchange rate"]

# 默认汇率（当找不到汇率列时使用）
DEFAULT_EXCHANGE_RATE = 0.2  # 1泰铢 ≈ 0.2人民币

# 定义数值列
numeric_cols = [
    "已完成产值（泰铢）", "合同应付款",
    "含税实付金额（泰铢）", "含税实付金额（人民币）", "含税实付金额（换算人民币）",
    "国内代付金额（人民币）", "换算泰铢金额", "付款汇总（泰铢）",
    "扣除借款（泰铢）", "扣除借款（人民币）",
    "扣除预扣税（泰铢）", "扣除预扣税（人民币）",
    "公司实际付款（财务数据）（泰铢）", "公司实际付款（财务数据）（人民币）"
]

# 定义泰铢到人民币的转换关系（泰铢列名 -> 人民币列名）
thb_to_rmb_mapping = {
    "含税实付金额（泰铢）": "含税实付金额（换算人民币）",
    "扣除借款（泰铢）": "扣除借款（人民币）",
    "扣除预扣税（泰铢）": "扣除预扣税（人民币）",
    "公司实际付款（财务数据）（泰铢）": "公司实际付款（财务数据）（人民币）"
}

# ==============================================================================
# 2. 辅助函数
# ==============================================================================

def load_project_name_matcher_context():
    """加载项目名称匹配模块与数据库上下文。"""
    if not os.path.exists(PROJECT_MATCHER_MODULE_PATH):
        print(f"⚠️  未找到项目匹配模块: {PROJECT_MATCHER_MODULE_PATH}")
        return None

    try:
        spec = importlib.util.spec_from_file_location("project_name_matcher", PROJECT_MATCHER_MODULE_PATH)
        if spec is None or spec.loader is None:
            print("⚠️  项目匹配模块加载失败：无法创建模块规范")
            return None

        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        db = module.load_or_init_db(PROJECT_MATCHER_DB_PATH, [])
        print(f"✓ 已加载项目名称匹配数据库: {PROJECT_MATCHER_DB_PATH}")
        return {
            "module": module,
            "db": db,
            "db_path": PROJECT_MATCHER_DB_PATH,
        }
    except Exception as e:
        print(f"⚠️  加载项目名称匹配模块失败: {e}")
        return None


def match_with_project_name_db(extracted_name, matcher_ctx):
    """先按数据库精确映射匹配，未命中时再调用 project_name_matcher 模块匹配。"""
    if not extracted_name or not matcher_ctx:
        return None

    try:
        module = matcher_ctx["module"]
        db = matcher_ctx["db"]

        # 1) 仅用数据库做精确匹配（alias_to_standard / projects）
        raw_name = str(extracted_name).strip()
        alias_map = db.get("alias_to_standard", {}) or {}
        projects = db.get("projects", {}) or {}
        std_name = alias_map.get(raw_name)
        if std_name and std_name in projects:
            meta = projects.get(std_name, {})
            return {
                "project_code": meta.get("project_code", ""),
                "project_name": std_name,
                "source": "db_exact",
            }

        # 2) 数据库未命中，再走模块匹配（归一化/模糊等）
        result = module.match_project(extracted_name, db)
        if result:
            module.save_db(matcher_ctx["db_path"], db)
            return {
                "project_code": result.get("project_code", ""),
                "project_name": result.get("project_name", extracted_name),
                "source": result.get("source", "project_name_db"),
            }
    except Exception as e:
        print(f"  ⚠️ 项目名称数据库匹配失败（{extracted_name}）: {e}")

    return None


def load_excel_formatting_utils():
    """加载统一的Excel格式化工具模块。"""
    if not os.path.exists(EXCEL_FORMATTING_UTILS_PATH):
        print(f"⚠️  未找到格式化模块: {EXCEL_FORMATTING_UTILS_PATH}")
        return None

    try:
        spec = importlib.util.spec_from_file_location("excel_formatting_utils", EXCEL_FORMATTING_UTILS_PATH)
        if spec is None or spec.loader is None:
            print("⚠️  格式化模块加载失败：无法创建模块规范")
            return None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        print(f"⚠️  加载格式化模块失败: {e}")
        return None

# 函数：从项目名称提取标准名称
def extract_project_name(project_name_str):
    """提取项目名：优先取“_分包管控表.xlsx”前缀，否则兼容旧规则。"""
    if not project_name_str:
        return None
    
    name = str(project_name_str).strip()

    # 若传入完整文件名，先去扩展名
    if name.lower().endswith('.xlsx'):
        name = os.path.splitext(name)[0]

    # 优先规则：识别“_分包管控表”之前的内容
    # 例如："邯郸某项目_分包管控表" -> "邯郸某项目"
    m = re.match(r'^(.*?)_分包管控表(?:\s*\(\d+\))?$', name)
    if m and m.group(1).strip():
        return m.group(1).strip()

    # 兼容无下划线写法
    m2 = re.match(r'^(.*?)\s*分包管控表(?:\s*\(\d+\))?$', name)
    if m2 and m2.group(1).strip():
        return m2.group(1).strip()
    
    # 移除常见后缀
    suffixes = [' 总计', '总计', ' 小计', '小计', ' 分包管控表', '分包管控表']
    for suffix in suffixes:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
    
    # 尝试匹配 "XX项目" 模式
    match = re.match(r'^(.+?项目)', name)
    if match:
        return match.group(1)
    
    return name

# 函数：从文本中提取人民币金额
def extract_rmb_amount(text):
    """
    从包含人民币的文本中提取数字金额。
    支持格式：
    - 国内代付200万人民币 -> 2000000
    - 国内支付668604.21 -> 668604.21
    - 国内代付1762846.99元 -> 1762846.99
    - 国内支付100万元整 -> 1000000
    """
    import re
    
    if pd.isna(text) or not isinstance(text, str):
        return None
    
    text = str(text).strip()
    
    # 检查是否包含人民币关键词
    rmb_keywords = ['人民币', '元', 'RMB', 'CNY']
    domestic_keywords = ['国内代付', '国内支付', '国内付款']
    
    has_rmb = any(kw in text for kw in rmb_keywords)
    has_domestic = any(kw in text for kw in domestic_keywords)
    
    if not (has_rmb or has_domestic):
        return None
    
    # 提取数字（支持小数和逗号分隔）
    # 先处理"万"单位
    wan_pattern = r'([0-9]+(?:\.[0-9]+)?)\s*万'
    wan_match = re.search(wan_pattern, text)
    if wan_match:
        amount = float(wan_match.group(1)) * 10000
        return amount
    
    # 提取普通数字（支持空格分隔，如 800 000）
    num_pattern = r'([0-9]+(?:[\s,][0-9]+)*(?:\.[0-9]+)?)'
    num_match = re.search(num_pattern, text)
    if num_match:
        num_str = num_match.group(1).replace(' ', '').replace(',', '')
        try:
            return float(num_str)
        except:
            return None
    
    return None

# ==============================================================================
# 3. 核心处理函数
# ==============================================================================

# 函数：自动检测header行
def find_header_row(df):
    """尝试找到包含足够多目标列名的行作为表头。"""
    for row_idx in range(min(20, len(df))):  # 只检查前20行提高效率
        # 将行数据转换为字符串并去除首尾空格
        row = df.iloc[row_idx].astype(str).str.strip()
        
        # 计算匹配目标列名的数量（使用模糊匹配）
        matched_count = 0
        for col in columns_to_extract:
            # 去除括号内容进行核心词匹配，提高匹配成功率
            core_keyword = col.split("（")[0].split("(")[0].strip()
            if any(core_keyword in str(cell) for cell in row if pd.notna(cell)):
                matched_count += 1
        
        # 降低阈值：至少匹配4个列或四分之一（更宽松）
        if matched_count >= 4 or matched_count >= len(columns_to_extract) // 4:
            return row_idx
    return None

# 函数：处理单个Sheet，返回清理后的数据DF和合计字典
def process_sheet(file_path, sheet_name, project_name, file_order, sheet_order, project_matcher_ctx=None, verification_log=None):
    """
    处理单个Excel Sheet，提取数据，计算人民币列，并排除公式行。
    verification_log: 可选，用于记录发现的公式单元格等异常情况
    """
    try:
        # 用两个工作簿：一个检测公式，一个获取计算值
        wb_formula = openpyxl.load_workbook(file_path, data_only=False)  # 保留公式用于检测
        wb_values = openpyxl.load_workbook(file_path, data_only=True)    # 获取计算值
        ws_formula = wb_formula[sheet_name]
        ws_values = wb_values[sheet_name]
        sheet_df_all = pd.DataFrame(ws_formula.values)
    except Exception as e:
        print(f"Error loading sheet '{sheet_name}' in project '{project_name}' via openpyxl: {str(e)}")
        return None, None
        
    header_row_idx = find_header_row(sheet_df_all)
    if header_row_idx is None:
        print(f"Warning: Could not find header in sheet '{sheet_name}' of project '{project_name}'")
        return None, None
    
    # 使用 openpyxl 提取数据（保留公式用于检测，但保存计算值）
    header_values = [str(cell.value).strip() if cell.value is not None else "" for cell in ws_formula[header_row_idx + 1]]
    data_rows = []
    formula_markers = []  # 标记每行是否包含列求和公式
    
    for row_idx in range(header_row_idx + 2, ws_formula.max_row + 1):
        row_data = []
        row_formulas = []  # 记录该行的公式
        
        for col_idx, (cell_formula, cell_value) in enumerate(zip(ws_formula[row_idx], ws_values[row_idx])):
            # 检查是否为公式
            if cell_formula.data_type == 'f':
                formula_text = f"={cell_formula.value}"
                row_formulas.append(formula_text)
                # 保存计算后的值（而非公式文本）
                row_data.append(cell_value.value)
            else:
                row_data.append(cell_formula.value)
        
        data_rows.append(row_data)
        formula_markers.append(row_formulas)
    
    df = pd.DataFrame(data_rows, columns=header_values)
    df['_formula_markers'] = formula_markers  # 临时列，用于后续过滤

    # 安全数字转换函数，避免类型比较错误
    def _safe_float(x):
        try:
            if pd.isna(x):
                return None
            s = str(x).strip()
            if s == "":
                return None
            s = s.replace(",", "").replace(" ", "")
            return float(s)
        except Exception:
            return None
    
    # 提取相关列（改进：使用核心关键词模糊匹配）
    # 映射别名，尽量匹配精确列
    alias_map = {
        "付款单编号": ["付款单编号", "付款编号", "付款单号"],
        "已完成产值（泰铢）": ["已完成产值（泰铢）", "已完成产值", "完成产值(THB)", "完成产值（泰铢）"],
        "合同应付款": ["合同应付款", "应付款", "合同应付"],
        "含税实付金额（泰铢）": ["含税实付金额（泰铢）", "含税实付金额", "实付金额(含税)(THB)"],
        "含税实付金额（人民币）": ["含税实付金额（人民币）", "含税实付金额(人民币)", "实付金额(RMB)"],
        "含税实付金额（换算人民币）": ["含税实付金额（换算人民币）", "含税实付金额(换算RMB)", "含税实付(换算人民币)"],
        "扣除借款（泰铢）": ["扣除借款（泰铢）", "扣除借款", "借款扣除(THB)"],
        "扣除借款（人民币）": ["扣除借款（人民币）", "扣除借款(RMB)"],
        "扣除预扣税（泰铢）": ["扣除预扣税（泰铢）", "预扣税(THB)", "扣除预扣税"],
        "扣除预扣税（人民币）": ["扣除预扣税（人民币）", "预扣税(RMB)"],
        "公司实际付款（财务数据）（泰铢）": ["公司实际付款（财务数据）（泰铢）", "公司实际付款(THB)", "实际付款（泰铢）"],
        "公司实际付款（财务数据）（人民币）": ["公司实际付款（财务数据）（人民币）", "公司实际付款(RMB)", "实际付款（人民币）"],
        "财务付款日期": ["财务付款日期", "付款日期", "支付日期"],
        "汇率": exchange_rate_variations,
    }

    def match_col(candidates):
        for cand in candidates:
            for col in df.columns:
                if str(col).strip() == cand:
                    return col
        # 退化到包含匹配
        for cand in candidates:
            for col in df.columns:
                if cand in str(col):
                    return col
        return None

    available_cols = {}
    for target_col in columns_to_extract:
        candidates = alias_map.get(target_col, [target_col])
        colname = match_col(candidates)
        if colname is not None:
            available_cols[target_col] = df[colname].copy()
    
    # 特殊处理汇率列：尝试多种可能的列名
    exchange_rate_col = None
    for rate_variation in exchange_rate_variations:
        matching = [c for c in df.columns if pd.notna(c) and rate_variation in str(c)]
        if matching:
            available_cols["汇率"] = df[matching[0]].copy()
            exchange_rate_col = matching[0]
            break
    
    if len(available_cols) < 2:
        print(f"Warning: Insufficient columns found in sheet '{sheet_name}' of project '{project_name}' (found {len(available_cols)} columns)")
        return None, None
    
    # 保留 _formula_markers 列用于后续过滤
    available_cols['_formula_markers'] = df['_formula_markers'].copy()
    
    sub_df = pd.DataFrame(available_cols)
    
    # ========== 新增：处理人民币数据 ==========
    # 检查"含税实付金额（泰铢）"列是否存在
    if "含税实付金额（泰铢）" in sub_df.columns:
        # 创建新列存放提取的人民币金额（保留原列文本不变，方便核对）
        sub_df["国内代付金额（人民币）"] = sub_df["含税实付金额（泰铢）"].apply(extract_rmb_amount)
        
        # 当检测到人民币金额且汇率为空时，设置默认汇率为5（1人民币=5泰铢）
        if "汇率" in sub_df.columns:
            for idx in sub_df.index:
                rmb_val = _safe_float(sub_df.loc[idx, "国内代付金额（人民币）"])
                if rmb_val is not None and rmb_val > 0:
                    # 检查汇率是否为空或无效
                    rate_val = _safe_float(sub_df.loc[idx, "汇率"]) or 0
                    if rate_val == 0:
                        sub_df.at[idx, "汇率"] = 5  # 默认汇率：1人民币=5泰铢
        else:
            # 如果没有汇率列，创建一个
            sub_df["汇率"] = None
            for idx in sub_df.index:
                rmb_val = _safe_float(sub_df.loc[idx, "国内代付金额（人民币）"])
                if rmb_val is not None and rmb_val > 0:
                    sub_df.at[idx, "汇率"] = 5
    else:
        sub_df["国内代付金额（人民币）"] = None
    
    # 添加占位列"换算泰铢金额"（将在Excel中用公式填充）
    sub_df["换算泰铢金额"] = None
    
    # 添加占位列"付款汇总（泰铢）"（将在Excel中用公式填充）
    sub_df["付款汇总（泰铢）"] = None
    # ========================================

    # 删除：目标列全部为空的行（用户要求）
    target_cols = [c for c in columns_to_extract if c in sub_df.columns]
    def is_empty(v):
        if pd.isna(v):
            return True
        if isinstance(v, str) and v.strip() == "":
            return True
        return False
    if target_cols:
        sub_df = sub_df[~sub_df[target_cols].apply(lambda r: all(is_empty(x) for x in r), axis=1)]

    # 删除：包含列求和公式的行，例如 "=SUM(V6:V17)"（不提取原始Excel中的汇总行）
    # 但保留行求和公式，例如 "=A2+B2" 或 "=SUM(A2:C2)"（正常数据处理）
    def is_column_sum_formula(v):
        """
        检测是否为列求和公式（垂直求和）
        - =SUM(V6:V17) -> True (同一列的多行求和)
        - =SUM(V:V) -> True (整列求和)
        - =SUM(A2:C2) -> False (同一行的多列求和)
        - =A2+B2 -> False (行内计算)
        """
        if not isinstance(v, str):
            return False
        formula = str(v).strip().upper()
        if not formula.startswith("="):
            return False
        
        # 检查是否包含 SUM 函数
        if "SUM(" not in formula:
            return False
        
        # 提取 SUM 函数中的范围
        import re
        # 匹配 SUM(范围) 模式
        sum_pattern = r'SUM\(([A-Z]+)(\d+)?:([A-Z]+)(\d+)?\)'
        matches = re.findall(sum_pattern, formula)
        
        for match in matches:
            col1, row1, col2, row2 = match
            # 如果列字母相同，说明是列求和（垂直）
            if col1 == col2:
                # 进一步检查：如果有行号且行号不同，确认是列求和
                if (row1 and row2 and row1 != row2) or (not row1 and not row2):
                    return True
        
        return False
    
    # 使用公式标记来过滤行
    if '_formula_markers' in sub_df.columns:
        # 检查每行的公式标记
        def row_has_column_sum(formulas):
            if not formulas:
                return False
            return any(is_column_sum_formula(f) for f in formulas)
        
        formula_rows_mask = sub_df['_formula_markers'].apply(row_has_column_sum)
        formula_rows = sub_df[formula_rows_mask]
        
        if not formula_rows.empty and verification_log is not None:
            # 提取公式示例
            formula_examples = []
            for _, row in formula_rows.head(3).iterrows():
                formulas = row['_formula_markers']
                for f in formulas:
                    if is_column_sum_formula(f):
                        formula_examples.append(f)
                        break
            
            verification_log.append({
                "文件名": os.path.basename(file_path),
                "项目名": project_name,
                "问题类型": "包含列求和公式",
                "详细信息": f"sheet: {sheet_name}, 汇总行数: {len(formula_rows)}, 示例: {'; '.join(formula_examples[:3])}"
            })
        
        # 删除包含列求和公式的行
        sub_df = sub_df[~formula_rows_mask]
        
        # 删除临时的公式标记列
        sub_df = sub_df.drop(columns=['_formula_markers'])

    # 删除：数值列全部为0的行（用户新增要求）
    check_numeric_cols = [
        "含税实付金额（泰铢）", "含税实付金额（人民币）", "含税实付金额（换算人民币）",
        "国内代付金额（人民币）", "换算泰铢金额", "付款汇总（泰铢）",
        "扣除借款（泰铢）", "扣除借款（人民币）",
        "扣除预扣税（泰铢）", "扣除预扣税（人民币）",
        "公司实际付款（财务数据）（泰铢）", "公司实际付款（财务数据）（人民币）"
    ]
    present_check_cols = [c for c in check_numeric_cols if c in sub_df.columns]
    if present_check_cols:
        def is_zero(v):
            if pd.isna(v):
                return True
            try:
                return float(v) == 0
            except:
                return False
        sub_df = sub_df[~sub_df[present_check_cols].apply(lambda r: all(is_zero(x) for x in r), axis=1)]

    # 如果无数据行则跳过
    if sub_df.empty:
        return None, None

    # 不计算子项合计
    sub_sums = None
    
    # 增加排序辅助列
    sub_df.insert(0, "sheet_order", sheet_order)
    sub_df.insert(0, "file_order", file_order)
    
    # ========== 新增：项目名称匹配与标准化 ==========
    # 提取标准项目名称
    extracted_name = extract_project_name(project_name)
    
    # 直接使用项目名称数据库作为唯一持久化来源
    matched_info = None
    
    if matched_info is None and extracted_name:
        db_match = match_with_project_name_db(extracted_name, project_matcher_ctx)
        if db_match:
            matched_info = {
                'project_code': db_match.get('project_code', ''),
                'project_name': db_match.get('project_name', extracted_name)
            }
            print(f"      ✓ 名称匹配成功: '{extracted_name}' → [{matched_info['project_code']}] {matched_info['project_name']} ({db_match.get('source', 'project_name_db')})")
        else:
            print(f"      ⚠️  未找到匹配项目: '{extracted_name}'")
    
    # 添加项目编码和项目名称列
    if matched_info and isinstance(matched_info, dict):
        sub_df.insert(2, "项目编码", matched_info.get('project_code', ''))
        sub_df.insert(3, "项目名称", matched_info.get('project_name', extracted_name or project_name))
    else:
        sub_df.insert(2, "项目编码", '')
        sub_df.insert(3, "项目名称", extracted_name or project_name)
    # ========================================
    
    return sub_df, sub_sums

# ==============================================================================
# 4. 主函数
# ==============================================================================

def process_month_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"❌ 目录不存在: {folder_path}")
        return False

    month_tag = os.path.basename(os.path.normpath(folder_path))

    print(f"\n{'='*60}")
    print(f"开始处理分包台账数据（月度目录）: {month_tag}")
    print(f"输入目录: {folder_path}")
    print(f"输出目录: {OUTPUT_FOLDER}")
    print(f"{'='*60}\n")
    
    # 仅加载项目名称匹配数据库与模块
    project_matcher_ctx = load_project_name_matcher_context()
    formatting_utils = load_excel_formatting_utils()
    print(f"✓ 项目名称匹配来源: {PROJECT_MATCHER_DB_PATH}")
    print("✓ 已启用：仅使用 project_name_match_db + project_name_matcher 进行匹配\n")
    
    # 新增：记录数据核对信息
    verification_log = []  # 格式：[{"文件名": "", "项目名": "", "问题类型": "", "详细信息": ""}]
    
    # ========== 新增：预扫描文件名中的项目并批量/手动匹配 ==========
    try:
        # 预获取文件列表
        pre_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and f not in ["详细汇总结果.xlsx", "汇总结果.xlsx"] and not f.startswith("~$")]
        # 提取并标准化项目名（从文件名）
        extracted_names = []
        for fn in pre_files:
            name = extract_project_name(fn)
            if name:
                extracted_names.append(name)
        unique_names = sorted(set(extracted_names))
        print(f"🔎 预扫描到 {len(unique_names)} 个唯一项目名")

        # 先自动匹配项目（统一写入 project_name_match_db）
        unmatched_names = []
        for name in unique_names:
            db_match = match_with_project_name_db(name, project_matcher_ctx)
            if db_match:
                print(f"   ✓ 名称匹配: '{name}' → [{db_match.get('project_code', '')}] {db_match.get('project_name', name)} ({db_match.get('source', 'project_name_db')})")
            else:
                unmatched_names.append(name)
                if project_matcher_ctx:
                    module = project_matcher_ctx["module"]
                    db = project_matcher_ctx["db"]
                    module.mark_unmatched(db, name)
                    module.save_db(project_matcher_ctx["db_path"], db)
                print(f"   ⚠️ 未匹配: '{name}'")

        if unmatched_names:
            print(f"⚠️ 仍有 {len(unmatched_names)} 个项目未匹配，已记录到数据库 unmatched_aliases")

        # 预扫描阶段写入 project_name_match_db 作为唯一持久化
        if project_matcher_ctx:
            project_matcher_ctx["module"].save_db(project_matcher_ctx["db_path"], project_matcher_ctx["db"])
    except Exception as e:
        print(f"⚠️ 预扫描项目匹配步骤异常：{e}")
    # ===============================================================
    
    # 收集所有数据
    all_data = []
    
    # 获取文件列表并排序，以确保项目顺序
    file_list = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and f not in ["详细汇总结果.xlsx", "汇总结果.xlsx"] and not f.startswith("~$")]
    
    print(f"📁 找到 {len(file_list)} 个Excel文件\n")
    
    # 遍历文件
    for file_order, file_name in enumerate(file_list, 1): # file_order 用于项目排序
        print(f"[{file_order}/{len(file_list)}] 正在处理: {file_name}")
        project_name = extract_project_name(file_name) or os.path.splitext(file_name)[0]
        file_path = os.path.join(folder_path, file_name)
        
        try:
            # 读取Excel所有sheets名称，确保按原始顺序遍历
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            
            # 检查是否有汇总sheet
            summary_sheets = [str(s) for s in xls.sheet_names if '汇总' in str(s)]
            if summary_sheets:
                verification_log.append({
                    "文件名": file_name,
                    "项目名": project_name,
                    "问题类型": "包含汇总sheet",
                    "详细信息": f"汇总sheet: {', '.join(summary_sheets)}"
                })
                print(f"   ⚠️  发现汇总sheet: {', '.join(summary_sheets)}（已跳过）")
            
            print(f"   └─ 共 {len(xls.sheet_names)} 个工作表")
            for sheet_order, sheet_name in enumerate(xls.sheet_names): # sheet_order 用于子项排序
                sheet_name = str(sheet_name)
                
                # 跳过汇总sheet
                if '汇总' in sheet_name:
                    print(f"      ⊘ {sheet_name}: 跳过（汇总sheet）")
                    continue
                
                # 处理单个Sheet（传入verification_log参数）
                sub_df, sub_sums = process_sheet(
                    file_path,
                    sheet_name,
                    project_name,
                    file_order,
                    sheet_order,
                    project_matcher_ctx,
                    verification_log
                )
                
                if sub_df is not None:
                    print(f"      ✓ {sheet_name}: 提取 {len(sub_df)} 行数据")
                else:
                    print(f"      ⊘ {sheet_name}: 无有效数据")
                
                if sub_df is not None:
                    # 添加项目和子项列
                    sub_df.insert(2, "子项", sheet_name)
                    sub_df.insert(2, "项目", project_name)
                    all_data.append(sub_df)
            
        except Exception as e:
            print(f"   ❌ 处理文件出错: {str(e)}")
            continue
    
    if not all_data:
        print(f"\n⚠️ 月份目录无有效数据，已跳过: {month_tag}")
        return False
    
    # 保存更新的项目名称匹配数据库
    if project_matcher_ctx:
        project_matcher_ctx["module"].save_db(project_matcher_ctx["db_path"], project_matcher_ctx["db"])
    
    print(f"\n{'='*60}")
    print(f"✓ 数据提取完成，共 {len(all_data)} 个工作表")
    print(f"{'='*60}\n")
    
    # 合并所有数据
    print("📊 正在合并数据...")
    full_df = pd.concat(all_data, ignore_index=True)
    
    # 排序：先按 file_order (项目)，再按 sheet_order (子项顺序)
    print("📊 正在排序数据...")
    full_df = full_df.sort_values(by=["file_order", "sheet_order"]) # 满足需求2
    print(f"✓ 数据合并完成，共 {len(full_df)} 行数据\n")
    
# ... (省略合并和排序部分)
    
    # 删除排序辅助列
    full_df = full_df.drop(columns=["file_order", "sheet_order"])
    
    # 预计算唯一项目名称列表（用于项目汇总页）
    unique_project_names = []
    try:
        unique_project_names = [name for name in pd.unique(full_df["项目名称"]) if pd.notna(name) and str(name).strip()]
    except Exception:
        unique_project_names = []

    # 预计算项目名称+子项唯一组合（用于子项汇总页）
    unique_project_subitems = []
    try:
        project_sub_df = full_df[["项目名称", "子项", "项目编码"]].copy()
        project_sub_df = project_sub_df[project_sub_df["项目名称"].notna() & project_sub_df["子项"].notna()]
        project_sub_df["项目名称"] = project_sub_df["项目名称"].astype(str).str.strip()
        project_sub_df["子项"] = project_sub_df["子项"].astype(str).str.strip()
        project_sub_df = project_sub_df[(project_sub_df["项目名称"] != "") & (project_sub_df["子项"] != "")]
        unique_project_subitems = project_sub_df.drop_duplicates(subset=["项目名称", "子项"]).to_dict("records")
    except Exception:
        unique_project_subitems = []
    
    # 输出到新Excel（文件名带时间戳）
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 确保输出目录存在
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"✓ 创建输出目录: {OUTPUT_FOLDER}\n")
    
    output_path = os.path.join(OUTPUT_FOLDER, f"{month_tag}_分包台账汇总_{ts}.xlsx")
    
    print(f"{'='*60}")
    print(f"正在生成Excel文件...")
    print(f"{'='*60}\n")
    
    # 定义最终的列顺序基准
    export_columns_to_extract = [c for c in columns_to_extract if c != "付款单编号"]
    final_cols_base = ["项目编码", "项目名称", "子项"] + export_columns_to_extract
    
    # 🌟 关键修复：动态过滤，只保留 full_df 中实际存在的列 🌟
    final_cols = [col for col in final_cols_base if col in full_df.columns]

    def to_numeric_series(series):
        if series is None:
            return pd.Series(dtype=float)
        cleaned = series.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
        cleaned = cleaned.replace({"nan": None, "None": None, "": None})
        return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

    def first_non_empty_value(series):
        for value in series:
            if pd.notna(value) and str(value).strip() != "":
                return str(value).strip()
        return ""

    summary_source_df = full_df.copy()
    summary_source_df["项目名称"] = summary_source_df.get("项目名称", "").fillna("").astype(str).str.strip()
    summary_source_df["项目编码"] = summary_source_df.get("项目编码", "").fillna("").astype(str).str.strip()
    summary_source_df["子项"] = summary_source_df.get("子项", "").fillna("").astype(str).str.strip()

    if "含税实付金额（泰铢）" in summary_source_df.columns:
        summary_source_df["含税实付金额（泰铢）"] = to_numeric_series(summary_source_df["含税实付金额（泰铢）"])
    if "国内代付金额（人民币）" in summary_source_df.columns:
        summary_source_df["国内代付金额（人民币）"] = to_numeric_series(summary_source_df["国内代付金额（人民币）"])
    if "汇率" in summary_source_df.columns:
        summary_source_df["汇率"] = to_numeric_series(summary_source_df["汇率"])
    else:
        summary_source_df["汇率"] = 0.0

    summary_source_df["换算泰铢金额_数值"] = summary_source_df.get("国内代付金额（人民币）", 0) * summary_source_df.get("汇率", 0)
    summary_source_df["付款汇总（泰铢）_数值"] = summary_source_df.get("含税实付金额（泰铢）", 0) + summary_source_df["换算泰铢金额_数值"]

    subitem_summary_df = pd.DataFrame()
    try:
        subitem_summary_df = (
            summary_source_df[
                (summary_source_df["项目名称"] != "") &
                (summary_source_df["子项"] != "")
            ]
            .groupby(["项目名称", "子项"], dropna=False)
            .agg({
                "项目编码": first_non_empty_value,
                "含税实付金额（泰铢）": "sum",
                "国内代付金额（人民币）": "sum",
                "换算泰铢金额_数值": "sum",
                "付款汇总（泰铢）_数值": "sum",
            })
            .reset_index()
        )
    except Exception:
        subitem_summary_df = pd.DataFrame()
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # 只创建分类汇总Sheet（不生成详细数据sheet）
        # ======================================================================
        # 分类汇总 Sheet
        # ======================================================================
        print("📝 正在生成分类汇总工作表...")
        wb = writer.book
        summary_ws = wb.create_sheet("分类汇总")
        # 删除默认的空白Sheet1
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        
        # 确保 headers 仅包含实际存在于 full_df 中的列
        headers = ["项目编码", "项目名称", "子项"] + [c for c in export_columns_to_extract if c in full_df.columns]
        for col_idx, header in enumerate(headers, start=1):
            summary_ws.cell(row=1, column=col_idx, value=header)
            
        # 找出数值列在 headers 中的索引
        numeric_col_indices = {col: headers.index(col) + 1 for col in numeric_cols if col in headers}
        subtotal_fill_color = "FFF2CC"  # 浅黄
        total_fill_color = "C6EFCE"     # 浅绿
        
        row_counter = 2
        current_project_name = None
        current_sub = None
        sub_start_row = None
        sub_data_start_row = None  # 子项数据开始行（用于三级分组）
        project_start_row = None
        project_subtotal_rows = []  # 记录当前项目的所有子项小计行号
        
        total_rows = len(full_df)
        processed_rows = 0
        last_progress = -1
        
        print("📊 正在写入分类汇总数据...")
        
        for _, row in full_df.iterrows():
            processed_rows += 1
            progress = int(processed_rows * 100 / total_rows)
            if progress != last_progress and progress % 10 == 0:
                print(f"   进度: {progress}% ({processed_rows}/{total_rows} 行)")
                last_progress = progress
            
            # --- 遇到新项目 ---
            row_project_name = row.get("项目名称", "")

            if row_project_name != current_project_name:
                if current_project_name is not None:
                    # 结束前一个子项的小计/分组
                    if current_sub is not None:
                        # 插入子项小计行
                        data_end = row_counter - 1
                        if sub_data_start_row is not None and sub_data_start_row <= data_end:
                            summary_ws.cell(row=row_counter, column=3, value=f"{current_sub} 小计")
                            # 为数值列添加求和公式
                            for col, col_idx in numeric_col_indices.items():
                                cell = summary_ws.cell(row=row_counter, column=col_idx)
                                cell.value = f"=SUM({get_column_letter(col_idx)}{sub_data_start_row}:{get_column_letter(col_idx)}{data_end})"
                            # 着色加粗整行
                            if formatting_utils:
                                formatting_utils.style_full_row(
                                    summary_ws,
                                    row_idx=row_counter,
                                    max_col=len(headers),
                                    fill_color=subtotal_fill_color,
                                    bold=True,
                                )
                            
                            # 记录小计行号用于项目总计求和
                            project_subtotal_rows.append(row_counter)
                            
                            # 先对子项数据行进行三级分组
                            for r in range(sub_data_start_row, data_end + 1):
                                summary_ws.row_dimensions[r].outlineLevel = 3
                                summary_ws.row_dimensions[r].hidden = False
                            
                            # 对整个子项（数据+小计）进行二级分组
                            if sub_start_row is not None:
                                for r in range(sub_start_row, row_counter + 1):
                                    if summary_ws.row_dimensions[r].outlineLevel != 3:
                                        summary_ws.row_dimensions[r].outlineLevel = 2
                                    summary_ws.row_dimensions[r].hidden = False
                            
                            row_counter += 1
                    
                    # 插入项目总计并分组（一级）
                    project_end = row_counter - 1
                    if project_start_row is not None and project_start_row <= project_end:
                        summary_ws.cell(row=row_counter, column=2, value=f"{current_project_name} 总计")
                        # 为数值列添加求和公式：只对子项小计行求和
                        for col, col_idx in numeric_col_indices.items():
                            cell = summary_ws.cell(row=row_counter, column=col_idx)
                            if project_subtotal_rows:
                                # 构建求和公式：只对小计行求和
                                sum_formula_parts = [f"{get_column_letter(col_idx)}{r}" for r in project_subtotal_rows]
                                cell.value = f"=SUM({','.join(sum_formula_parts)})"
                            else:
                                cell.value = 0
                        # 着色加粗整行
                        if formatting_utils:
                            formatting_utils.style_full_row(
                                summary_ws,
                                row_idx=row_counter,
                                max_col=len(headers),
                                fill_color=total_fill_color,
                                bold=True,
                            )
                        
                        # 对整个项目（所有子项+总计）进行一级分组
                        for r in range(project_start_row, row_counter + 1):
                            if summary_ws.row_dimensions[r].outlineLevel == 0 or summary_ws.row_dimensions[r].outlineLevel is None:
                                summary_ws.row_dimensions[r].outlineLevel = 1
                            summary_ws.row_dimensions[r].hidden = False
                        
                        row_counter += 1
                
                # 开始新项目
                current_project_name = row_project_name
                project_start_row = row_counter
                project_subtotal_rows = []  # 重置小计行记录
                current_sub = None # 重置子项
            
            # --- 遇到新子项 ---
            if row["子项"] != current_sub:
                if current_sub is not None:
                    # 插入前一个子项小计
                    data_end = row_counter - 1
                    if sub_data_start_row is not None and sub_data_start_row <= data_end:
                        summary_ws.cell(row=row_counter, column=3, value=f"{current_sub} 小计")
                        # 为数值列添加求和公式
                        for col, col_idx in numeric_col_indices.items():
                            cell = summary_ws.cell(row=row_counter, column=col_idx)
                            cell.value = f"=SUM({get_column_letter(col_idx)}{sub_data_start_row}:{get_column_letter(col_idx)}{data_end})"
                        # 着色加粗整行
                        if formatting_utils:
                            formatting_utils.style_full_row(
                                summary_ws,
                                row_idx=row_counter,
                                max_col=len(headers),
                                fill_color=subtotal_fill_color,
                                bold=True,
                            )
                        
                        # 记录小计行号用于项目总计求和
                        project_subtotal_rows.append(row_counter)
                        
                        # 先对子项数据行进行三级分组
                        for r in range(sub_data_start_row, data_end + 1):
                            summary_ws.row_dimensions[r].outlineLevel = 3
                            summary_ws.row_dimensions[r].hidden = False
                        
                        # 对整个子项（数据+小计）进行二级分组
                        if sub_start_row is not None:
                            for r in range(sub_start_row, row_counter + 1):
                                if summary_ws.row_dimensions[r].outlineLevel != 3:
                                    summary_ws.row_dimensions[r].outlineLevel = 2
                                summary_ws.row_dimensions[r].hidden = False
                        
                        row_counter += 1
                
                # 开始新子项
                current_sub = row["子项"]
                sub_start_row = row_counter  # 子项开始行（包括后续的数据和小计）
                sub_data_start_row = row_counter  # 数据开始行
            
            # --- 写入数据行 ---
            for col_idx, col in enumerate(headers, start=1):
                value = row.get(col)
                cell = summary_ws.cell(row=row_counter, column=col_idx, value=value)
                
                # 为"换算泰铢金额"列添加公式（人民币*汇率）
                if col == "换算泰铢金额":
                    # 找到"国内代付金额（人民币）"和"汇率"列的索引
                    try:
                        rmb_col_idx = headers.index("国内代付金额（人民币）") + 1
                        rate_col_idx = headers.index("汇率") + 1
                        rmb_col_letter = get_column_letter(rmb_col_idx)
                        rate_col_letter = get_column_letter(rate_col_idx)
                        # 公式：=IF(人民币列有值, 人民币*汇率, "")
                        cell.value = f"=IF(ISNUMBER({rmb_col_letter}{row_counter}), {rmb_col_letter}{row_counter}*{rate_col_letter}{row_counter}, \"\")"
                    except ValueError:
                        pass  # 如果找不到列，跳过
                
                # 为"付款汇总（泰铢）"列添加公式（含税实付金额+换算泰铢金额）
                if col == "付款汇总（泰铢）":
                    try:
                        thb_col_idx = headers.index("含税实付金额（泰铢）") + 1
                        convert_col_idx = headers.index("换算泰铢金额") + 1
                        thb_col_letter = get_column_letter(thb_col_idx)
                        convert_col_letter = get_column_letter(convert_col_idx)
                        # 公式：=IF(ISNUMBER(含税实付金额列), 含税实付金额, 0) + IF(ISNUMBER(换算泰铢列), 换算泰铢, 0)
                        cell.value = f"=IF(ISNUMBER({thb_col_letter}{row_counter}), {thb_col_letter}{row_counter}, 0) + IF(ISNUMBER({convert_col_letter}{row_counter}), {convert_col_letter}{row_counter}, 0)"
                    except ValueError:
                        pass  # 如果找不到列，跳过
                
            row_counter += 1
        
        # --- 循环结束后的收尾工作 ---
        
        # 1. 最后一个子项
        if current_sub is not None:
            # 插入子项小计
            data_end = row_counter - 1
            if sub_data_start_row is not None and sub_data_start_row <= data_end:
                summary_ws.cell(row=row_counter, column=3, value=f"{current_sub} 小计")
                # 为数值列添加求和公式
                for col, col_idx in numeric_col_indices.items():
                    cell = summary_ws.cell(row=row_counter, column=col_idx)
                    cell.value = f"=SUM({get_column_letter(col_idx)}{sub_data_start_row}:{get_column_letter(col_idx)}{data_end})"
                # 着色加粗整行
                if formatting_utils:
                    formatting_utils.style_full_row(
                        summary_ws,
                        row_idx=row_counter,
                        max_col=len(headers),
                        fill_color=subtotal_fill_color,
                        bold=True,
                    )
                
                # 记录小计行号用于项目总计求和
                project_subtotal_rows.append(row_counter)
                
                # 先对子项数据行进行三级分组
                for r in range(sub_data_start_row, data_end + 1):
                    summary_ws.row_dimensions[r].outlineLevel = 3
                    summary_ws.row_dimensions[r].hidden = False
                
                # 对整个子项（数据+小计）进行二级分组
                if sub_start_row is not None:
                    for r in range(sub_start_row, row_counter + 1):
                        if summary_ws.row_dimensions[r].outlineLevel != 3:
                            summary_ws.row_dimensions[r].outlineLevel = 2
                        summary_ws.row_dimensions[r].hidden = False
                
                row_counter += 1
        
        # 2. 最后一个项目
        if current_project_name is not None:
            # 插入项目总计并分组（一级）
            project_end = row_counter - 1
            if project_start_row is not None and project_start_row <= project_end:
                summary_ws.cell(row=row_counter, column=2, value=f"{current_project_name} 总计")
                # 为数值列添加求和公式：只对子项小计行求和
                for col, col_idx in numeric_col_indices.items():
                    cell = summary_ws.cell(row=row_counter, column=col_idx)
                    if project_subtotal_rows:
                        # 构建求和公式：只对小计行求和
                        sum_formula_parts = [f"{get_column_letter(col_idx)}{r}" for r in project_subtotal_rows]
                        cell.value = f"=SUM({','.join(sum_formula_parts)})"
                    else:
                        cell.value = 0
                # 着色加粗整行
                if formatting_utils:
                    formatting_utils.style_full_row(
                        summary_ws,
                        row_idx=row_counter,
                        max_col=len(headers),
                        fill_color=total_fill_color,
                        bold=True,
                    )
                
                # 对整个项目（所有子项+总计）进行一级分组
                for r in range(project_start_row, row_counter + 1):
                    if summary_ws.row_dimensions[r].outlineLevel == 0 or summary_ws.row_dimensions[r].outlineLevel is None:
                        summary_ws.row_dimensions[r].outlineLevel = 1
                    summary_ws.row_dimensions[r].hidden = False
            row_counter += 1
        
        # 为"分类汇总"sheet应用格式
        # 首行冻结
        if formatting_utils:
            formatting_utils.set_header_row_style(summary_ws, header_row=1, header_height=40)
            summary_ws.freeze_panes = "A2"
            formatting_utils.set_columns_number_format(
                summary_ws,
                col_indices=numeric_col_indices.values(),
                number_format="#,##0.00",
                start_row=2,
                end_row=row_counter - 1,
            )
            if row_counter > 2:
                formatting_utils.set_auto_filter_range(
                    summary_ws,
                    start_col=1,
                    end_col=len(headers),
                    start_row=1,
                    end_row=1,
                )
        
        print(f"\n✓ 分类汇总完成，共 {row_counter-1} 行（含小计和总计）\n")

        # ======================================================================
        # 新增：项目汇总（万泰铢） Sheet
        # ======================================================================
        print("🧮 正在生成项目汇总（万泰铢）工作表...")
        summary2_ws = wb.create_sheet("项目汇总（万泰铢）")
        # 表头
        summary2_headers = [
            "项目名称", "项目编码",
            "含税实付金额（万泰铢）", "国内代付金额（人民币）", "国内代付人民币金额（万泰铢）", "付款汇总（万泰铢）"
        ]
        for cidx, h in enumerate(summary2_headers, start=1):
            summary2_ws.cell(row=1, column=cidx, value=h)

        # 计算在分类汇总中的关键列位置
        def _col_letter_in_summary(col_name):
            try:
                ci = headers.index(col_name) + 1
                return get_column_letter(ci)
            except ValueError:
                return None

        project_name_col_letter = _col_letter_in_summary("项目名称") or "C"
        code_col_letter = _col_letter_in_summary("项目编码")
        thb_col_letter = _col_letter_in_summary("含税实付金额（泰铢）")
        converted_thb_col_letter = _col_letter_in_summary("换算泰铢金额")
        pay_sum_col_letter = _col_letter_in_summary("付款汇总（泰铢）")
        rmb_domestic_col_letter = _col_letter_in_summary("国内代付金额（人民币）")

        # 数据行
        r = 2
        for project_name in unique_project_names:
            summary2_ws.cell(row=r, column=1, value=project_name)

            # 项目编码：INDEX/MATCH
            if code_col_letter:
                summary2_ws.cell(row=r, column=2, value=f"=INDEX(分类汇总!{code_col_letter}:{code_col_letter}, MATCH(A{r}, 分类汇总!{project_name_col_letter}:{project_name_col_letter}, 0))")
            else:
                summary2_ws.cell(row=r, column=2, value="")

            # 含税实付金额（万泰铢） = SUMIF(..., 含税实付金额（泰铢）)/10000
            if thb_col_letter:
                summary2_ws.cell(row=r, column=3, value=f"=SUMIF(分类汇总!{project_name_col_letter}:{project_name_col_letter}, A{r}, 分类汇总!{thb_col_letter}:{thb_col_letter})/10000")
            else:
                summary2_ws.cell(row=r, column=3, value="")

            # 国内代付金额（万人民币） = SUMIF(..., 国内代付金额（人民币）)/10000
            if rmb_domestic_col_letter:
                summary2_ws.cell(row=r, column=4, value=f"=SUMIF(分类汇总!{project_name_col_letter}:{project_name_col_letter}, A{r}, 分类汇总!{rmb_domestic_col_letter}:{rmb_domestic_col_letter})/10000")
            else:
                summary2_ws.cell(row=r, column=4, value="")

            # 国内代付人民币金额（万泰铢） = SUMIF(..., 换算泰铢金额)/10000
            if converted_thb_col_letter:
                summary2_ws.cell(row=r, column=5, value=f"=SUMIF(分类汇总!{project_name_col_letter}:{project_name_col_letter}, A{r}, 分类汇总!{converted_thb_col_letter}:{converted_thb_col_letter})/10000")
            else:
                summary2_ws.cell(row=r, column=5, value="")

            # 付款汇总（万泰铢） = SUMIF(..., 付款汇总（泰铢）)/10000
            if pay_sum_col_letter:
                summary2_ws.cell(row=r, column=6, value=f"=SUMIF(分类汇总!{project_name_col_letter}:{project_name_col_letter}, A{r}, 分类汇总!{pay_sum_col_letter}:{pay_sum_col_letter})/10000")
            else:
                summary2_ws.cell(row=r, column=6, value="")

            r += 1

        # 格式：冻结首行、自动筛选、调整首行高度和居中换行
        if formatting_utils:
            formatting_utils.set_header_row_style(summary2_ws, header_row=1, header_height=40)
        summary2_ws.freeze_panes = "A2"
        if r > 2 and formatting_utils:
            formatting_utils.set_auto_filter_range(
                summary2_ws,
                start_col=1,
                end_col=len(summary2_headers),
                start_row=1,
                end_row=1,
            )

        # 设置金额列为整数千分位显示（无货币符号）
        amount_cols_by_name = [
            "国内代付金额（人民币）",
            "含税实付金额（万泰铢）",
            "国内代付人民币金额（万泰铢）",
            "付款汇总（万泰铢）",
        ]
        amount_col_indices = [summary2_headers.index(n) + 1 for n in amount_cols_by_name if n in summary2_headers]
        if formatting_utils:
            formatting_utils.set_columns_number_format(
                summary2_ws,
                col_indices=amount_col_indices,
                number_format="#,##0",
                start_row=2,
                end_row=r - 1,
            )

        # 自适应列宽（对文本列按内容长度，金额列设合理宽度）
        if formatting_utils:
            formatting_utils.auto_adjust_column_widths(
                summary2_ws,
                start_col=1,
                end_col=len(summary2_headers),
                start_row=1,
                end_row=max(r - 1, 1),
                min_col_width=10,
                max_col_width=60,
                extra_padding=2,
                ignore_formula=True,
            )
            for col_idx in (3, 4, 5, 6):
                col_letter = get_column_letter(col_idx)
                if summary2_ws.column_dimensions[col_letter].width < 14:
                    summary2_ws.column_dimensions[col_letter].width = 14
        
        print(f"\n✓ 项目汇总（万泰铢）完成\n")

        # ======================================================================
        # 新增：项目名称子项汇总（万泰铢） Sheet
        # ======================================================================
        print("🧮 正在生成项目名称子项汇总（万泰铢）工作表...")
        summary3_ws = wb.create_sheet("项目名称子项汇总（万泰铢）")
        summary3_headers = [
            "项目名称", "项目编码", "子项",
            "含税实付金额（万泰铢）", "国内代付金额（人民币）", "国内代付人民币金额（万泰铢）", "付款汇总（万泰铢）"
        ]
        for cidx, h in enumerate(summary3_headers, start=1):
            summary3_ws.cell(row=1, column=cidx, value=h)

        r3 = 2
        for _, item in subitem_summary_df.iterrows():
            project_name = item.get("项目名称", "")
            subitem_name = item.get("子项", "")
            project_code = item.get("项目编码", "")

            summary3_ws.cell(row=r3, column=1, value=project_name)
            summary3_ws.cell(row=r3, column=2, value=project_code)
            summary3_ws.cell(row=r3, column=3, value=subitem_name)

            summary3_ws.cell(row=r3, column=4, value=float(item.get("含税实付金额（泰铢）", 0) or 0) / 10000)
            summary3_ws.cell(row=r3, column=5, value=float(item.get("国内代付金额（人民币）", 0) or 0) / 10000)
            summary3_ws.cell(row=r3, column=6, value=float(item.get("换算泰铢金额_数值", 0) or 0) / 10000)
            summary3_ws.cell(row=r3, column=7, value=float(item.get("付款汇总（泰铢）_数值", 0) or 0) / 10000)

            r3 += 1

        if formatting_utils:
            formatting_utils.set_header_row_style(summary3_ws, header_row=1, header_height=40)
        summary3_ws.freeze_panes = "A2"
        if r3 > 2 and formatting_utils:
            formatting_utils.set_auto_filter_range(
                summary3_ws,
                start_col=1,
                end_col=len(summary3_headers),
                start_row=1,
                end_row=1,
            )

        summary3_amount_cols = [4, 5, 6, 7]
        if formatting_utils:
            formatting_utils.set_columns_number_format(
                summary3_ws,
                col_indices=summary3_amount_cols,
                number_format="#,##0.00",
                start_row=2,
                end_row=r3 - 1,
            )

        if formatting_utils:
            formatting_utils.auto_adjust_column_widths(
                summary3_ws,
                start_col=1,
                end_col=len(summary3_headers),
                start_row=1,
                end_row=max(r3 - 1, 1),
                min_col_width=10,
                max_col_width=60,
                extra_padding=2,
                ignore_formula=True,
            )
            for col_idx in summary3_amount_cols:
                col_letter = get_column_letter(col_idx)
                if summary3_ws.column_dimensions[col_letter].width < 14:
                    summary3_ws.column_dimensions[col_letter].width = 14

        print(f"\n✓ 项目名称子项汇总（万泰铢）完成\n")
        
        # ======================================================================
        # 新增：数据核对信息 Sheet
        # ======================================================================
        if verification_log:
            print(f"📋 正在生成数据核对信息工作表...")
            verify_ws = wb.create_sheet("数据核对信息")
            
            # 表头
            verify_headers = ["文件名", "项目名", "问题类型", "详细信息"]
            for cidx, h in enumerate(verify_headers, start=1):
                verify_ws.cell(row=1, column=cidx, value=h)

            if formatting_utils:
                formatting_utils.set_header_row_style(
                    verify_ws,
                    header_row=1,
                    header_height=30,
                    wrap_text=True,
                    vertical='center',
                    horizontal='center',
                    bold=True,
                )
                formatting_utils.style_full_row(
                    verify_ws,
                    row_idx=1,
                    max_col=len(verify_headers),
                    fill_color="D9E1F2",
                    bold=True,
                )
            
            # 数据行
            for row_idx, log_entry in enumerate(verification_log, start=2):
                verify_ws.cell(row=row_idx, column=1, value=log_entry.get("文件名", ""))
                verify_ws.cell(row=row_idx, column=2, value=log_entry.get("项目名", ""))
                verify_ws.cell(row=row_idx, column=3, value=log_entry.get("问题类型", ""))
                verify_ws.cell(row=row_idx, column=4, value=log_entry.get("详细信息", ""))
            
            # 格式设置
            verify_ws.freeze_panes = "A2"
            verify_ws.row_dimensions[1].height = 30
            if len(verification_log) > 0 and formatting_utils:
                formatting_utils.set_auto_filter_range(
                    verify_ws,
                    start_col=1,
                    end_col=4,
                    start_row=1,
                    end_row=len(verification_log) + 1,
                )
            
            # 自适应列宽
            if formatting_utils:
                formatting_utils.auto_adjust_column_widths(
                    verify_ws,
                    start_col=1,
                    end_col=4,
                    start_row=1,
                    end_row=len(verification_log) + 1,
                    min_col_width=10,
                    max_col_width=50,
                    extra_padding=2,
                    ignore_formula=True,
                )
            
            print(f"✓ 数据核对信息完成，记录 {len(verification_log)} 条异常情况\n")
        else:
            print(f"✓ 未发现需要核对的异常情况\n")
    
    print(f"{'='*60}")
    print(f"✅ 处理完成！")
    print(f"{'='*60}")
    print(f"📄 输出文件: {os.path.basename(output_path)}")
    print(f"📁 保存位置: {OUTPUT_FOLDER}")
    print(f"\n说明:")
    print(f"  • 工作表包含三级分组结构")
    print(f"    - 级别1: 项目总计（绿色）")
    print(f"    - 级别2: 子项小计（黄色）")
    print(f"    - 级别3: 明细数据")
    print(f"  • 点击Excel左侧的1/2/3按钮可折叠/展开不同级别")
    print(f"  • 项目总计仅对各子项小计求和，避免重复计算")
    if verification_log:
        print(f"  • 查看【数据核对信息】工作表了解跳过的汇总sheet和公式行")
    print(f"{'='*60}\n")
    return True


def run_post_analysis_script():
    """汇总完成后执行后置分析脚本。"""
    if not os.path.exists(POST_ANALYSIS_SCRIPT):
        print(f"⚠️ 未找到后置分析脚本，已跳过: {POST_ANALYSIS_SCRIPT}")
        return False

    print(f"\n{'='*60}")
    print("🚀 开始执行后置分析脚本（按日期汇总与月增量）...")
    print(f"脚本路径: {POST_ANALYSIS_SCRIPT}")
    print(f"{'='*60}\n")

    try:
        subprocess.run([sys.executable, POST_ANALYSIS_SCRIPT], check=True)
        print("\n✅ 后置分析脚本执行完成")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 后置分析脚本执行失败，退出码: {e.returncode}")
        return False
    except Exception as e:
        print(f"\n❌ 后置分析脚本执行异常: {e}")
        return False


def main():
    input_root = INPUT_FOLDER
    output_root = OUTPUT_FOLDER

    if not os.path.isdir(input_root):
        print(f"❌ 输入根目录不存在: {input_root}")
        return

    os.makedirs(output_root, exist_ok=True)

    month_folders = [
        os.path.join(input_root, d)
        for d in sorted(os.listdir(input_root))
        if os.path.isdir(os.path.join(input_root, d))
    ]

    if not month_folders:
        print("⚠️ 输入根目录下未发现子文件夹，将按单目录模式处理")
        process_month_folder(input_root)
        return

    print(f"\n检测到 {len(month_folders)} 个月份目录，开始批量处理...")

    success_count = 0
    for idx, month_folder in enumerate(month_folders, 1):
        print(f"\n[{idx}/{len(month_folders)}] 处理中: {os.path.basename(month_folder)}")
        ok = process_month_folder(month_folder)
        if ok:
            success_count += 1

    print(f"\n批量处理完成：成功 {success_count} 个，跳过/失败 {len(month_folders) - success_count} 个")

    if success_count > 0:
        run_post_analysis_script()
    else:
        print("⚠️ 未生成任何有效月度汇总文件，跳过后置分析脚本执行")

if __name__ == "__main__":
    main()