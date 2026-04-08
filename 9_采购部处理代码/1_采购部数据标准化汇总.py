import os
from copy import copy
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import re
from datetime import datetime
import sys
import importlib.util

PROJECT_MATCHER_MODULE_PATH = '/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/project_name_matcher.py'


def _load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_project_name_matcher = _load_module_from_path('project_name_matcher', PROJECT_MATCHER_MODULE_PATH)
add_mapping = _project_name_matcher.add_mapping
get_project_lookup = _project_name_matcher.get_project_lookup
is_unmatched_alias = _project_name_matcher.is_unmatched_alias
load_or_init_db = _project_name_matcher.load_or_init_db
mark_unmatched = _project_name_matcher.mark_unmatched
match_project = _project_name_matcher.match_project
prompt_unmatched_project_mappings = _project_name_matcher.prompt_unmatched_project_mappings
save_db = _project_name_matcher.save_db

# ============= 配置区域 =============
# 原始数据根目录（用于A_合并汇总）
RAW_DATA_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购部采购台账/A_原始数据"
# 已标准化数据根目录（用于B_标准化汇总）
STANDARDIZED_DATA_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购部采购台账/B_采购部标准化"
# 汇总输出根目录
SUMMARY_OUTPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购部采购台账/B_汇总数据"

PROJECT_INFO_FILE = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
PROJECT_MATCH_DB_FILE = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/project_name_match_db.json"
EXCEL_FORMATTING_UTILS_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/excel_formatting_utils.py"
OUTPUT_MERGED_FILENAME = '材料采购_merged.xlsx'
# ===================================

# 全局匹配库（来自 19_代码/project_name_matcher.py）
PROJECT_MATCH_DB = None

# Define the standard columns
standard_columns = [
    '材料归类', '物品名称', '规格', '标准单位', '数量', 
    '单价', '不含税金额', '7%增值税', '含税金额', '货币种类', '汇率', 
    '签订日期', '供货商名称', '备注'
]

# Column name mappings for variations
column_mappings = {
    '物品名称': ['物品名称', '采购物资名称', '物资名称'],  # Add more variations as needed
    # You can add mappings for other columns if variations exist, e.g.,
    # '规格': ['规格', '规格型号'],
}

# 定义核心必需列（缺少这些列的文件才会被跳过）
CORE_REQUIRED_COLUMNS = [
    '物品名称', '单位', '数量', '单价', '不含税金额', '含税金额'
]

# 定义可选列（缺少时可以跳过）
OPTIONAL_COLUMNS = [
    '材料归类', '规格', '货币种类', '汇率', '签订日期', '供货商名称', '备注', '7%增值税'
]

FINAL_OUTPUT_COLUMNS = [
    '项目编码', '项目名称', '材料归类', '物品名称', '规格', '标准单位', '数量',
    '单价', '不含税金额', '7%增值税', '含税金额', '货币种类', '汇率', '年月',
    '标准日期', '供货商名称', '备注', '签订日期'
]

CENTER_ALIGN_COLUMNS = {'标准单位', '数量', '货币种类', '汇率'}
LEFT_ALIGN_COLUMNS = {'签订日期'}

_excel_formatting_utils = _load_module_from_path('excel_formatting_utils', EXCEL_FORMATTING_UTILS_PATH)
apply_common_excel_layout = _excel_formatting_utils.apply_common_excel_layout


def apply_unified_workbook_format(output_wb):
    """统一使用excel_formatting_utils模块格式化输出工作簿。"""
    amount_keywords = ('金额', '单价', '税')

    for sheet_name in ['合并数据', '项目汇总']:
        if sheet_name not in output_wb.sheetnames:
            continue

        ws = output_wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        amount_headers = [
            h for h in headers
            if isinstance(h, str) and any(k in h for k in amount_keywords)
        ]

        apply_common_excel_layout(
            ws,
            numeric_start_col=ws.max_column + 1,
            amount_headers=amount_headers,
            header_row=1,
            header_height=40,
            freeze_panes='A2',
        )

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

def show_batch_project_selection_dialog(unmatched_projects, project_config):
    """显示批量项目匹配对话框"""
    class BatchProjectSelectionDialog:
        def __init__(self, parent, unmatched_projects, all_projects):
            self.results = {}
            self.top = tk.Toplevel(parent)
            self.top.title(f"批量项目匹配 - 共{len(unmatched_projects)}个项目需要匹配")
            self.top.geometry("800x600")
            
            # 确保窗口显示在最前面
            self.top.lift()
            self.top.attributes('-topmost', True)
            self.top.after_idle(self.top.attributes, '-topmost', False)
            
            # 说明文字
            label = tk.Label(self.top, 
                           text=f"以下 {len(unmatched_projects)} 个项目未找到匹配，请为每个项目选择对应的标准项目（或留空）",
                           font=('Arial', 12, 'bold'),
                           pady=15)
            label.pack()
            
            # 滚动框架
            canvas = tk.Canvas(self.top)
            scrollbar = tk.Scrollbar(self.top, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # 准备下拉选项
            self.all_projects = all_projects
            self.project_options = ['（留空）'] + [f"{p['project_code']} - {p['project_name']}" for p in all_projects]
            self.combo_vars = {}
            
            # 为每个未匹配项目创建选择框
            for idx, project_name in enumerate(unmatched_projects):
                row_frame = tk.Frame(scrollable_frame, relief=tk.RIDGE, borderwidth=1)
                row_frame.pack(fill='x', padx=10, pady=5)
                
                # 项目名称标签
                name_label = tk.Label(row_frame, text=project_name, 
                                     font=('Arial', 11, 'bold'),
                                     width=25, anchor='w')
                name_label.pack(side='left', padx=10, pady=10)
                
                # 箭头
                tk.Label(row_frame, text="→", font=('Arial', 12)).pack(side='left', padx=5)
                
                # 下拉选择框
                combo_var = tk.StringVar()
                combo = ttk.Combobox(row_frame, 
                                   textvariable=combo_var,
                                   values=self.project_options,
                                   width=45,
                                   font=('Arial', 10),
                                   state='readonly')
                combo.pack(side='left', padx=10, pady=10, fill='x', expand=True)
                combo.current(0)  # 默认选择空
                
                self.combo_vars[project_name] = combo_var
            
            canvas.pack(side="left", fill="both", expand=True, padx=10)
            scrollbar.pack(side="right", fill="y")
            
            # 按钮
            btn_frame = tk.Frame(self.top)
            btn_frame.pack(pady=15)
            
            tk.Button(btn_frame, text="确定并保存", command=self.ok, 
                     width=15, height=1, font=('Arial', 11, 'bold'),
                     bg='#4CAF50', fg='white').pack(side='left', padx=10)
            tk.Button(btn_frame, text="全部跳过", command=self.skip_all, 
                     width=15, height=1, font=('Arial', 11)).pack(side='left', padx=10)
            
            self.top.protocol("WM_DELETE_WINDOW", self.skip_all)
            
            # 居中显示
            self.top.update_idletasks()
            x = (self.top.winfo_screenwidth() // 2) - (self.top.winfo_width() // 2)
            y = (self.top.winfo_screenheight() // 2) - (self.top.winfo_height() // 2)
            self.top.geometry(f"+{x}+{y}")
        
        def ok(self):
            # 收集所有选择
            for project_name, combo_var in self.combo_vars.items():
                selected = combo_var.get()
                if selected and selected != '（留空）':
                    # 找到对应的项目
                    for i, option in enumerate(self.project_options[1:], start=0):
                        if option == selected:
                            self.results[project_name] = self.all_projects[i]
                            break
                else:
                    self.results[project_name] = None
            self.top.destroy()
        
        def skip_all(self):
            # 全部设为None
            for project_name in self.combo_vars.keys():
                self.results[project_name] = None
            self.top.destroy()
    
    # 创建主窗口
    root = tk.Tk()
    root.withdraw()
    root.update()
    
    # 创建对话框
    dialog = BatchProjectSelectionDialog(root, unmatched_projects, project_config)
    
    # 等待对话框关闭
    root.wait_window(dialog.top)
    
    results = dialog.results
    root.destroy()
    
    return results

def show_project_selection_dialog(extracted_name, project_config):
    """显示项目选择对话框（单个）"""
    class ProjectSelectionDialog:
        def __init__(self, parent, extracted_name, projects):
            self.result = None
            self.top = tk.Toplevel(parent)
            self.top.title(f"项目匹50x200")
            
            # 确保窗口显示在最前面
            self.top.lift()
            self.top.attributes('-topmost', True)
            self.top.after_idle(self.top.attributes, '-topmost', False)
            
            # 说明文字
            label = tk.Label(self.top, 
                           text=f"未找到项目 '{extracted_name}' 的匹配项\n请从下拉框中选择对应的标准项目（或留空）",
                           font=('Arial', 12),
                           pady=15)
            label.pack()
            
            # 下拉框
            combo_frame = tk.Frame(self.top)
            combo_frame.pack(fill='x', padx=20, pady=10)
            
            tk.Label(combo_frame, text="选择项目：", font=('Arial', 11)).pack(side='left', padx=5)
            
            # 准备下拉选项
            self.projects = projects
            self.project_options = [''] + [f"{p['project_code']} - {p['project_name']}" for p in projects]
            
            self.combo_var = tk.StringVar()
            self.combo = ttk.Combobox(combo_frame, 
                                     textvariable=self.combo_var,
                                     values=self.project_options,
                                     width=50,
                                     font=('Arial', 11),
                                     state='readonly')
            self.combo.pack(side='left', padx=5, fill='x', expand=True)
            self.combo.current(0)  # 默认选择空
            
            # 按钮
            btn_frame = tk.Frame(self.top)
            btn_frame.pack(pady=20)
            
            tk.Button(btn_frame, text="确定", command=self.ok, 
                     width=12, height=1, font=('Arial', 11)).pack(side='left', padx=10)
            tk.Button(btn_frame, text="跳过（留空）", command=self.skip, 
                     width=12, height=1, font=('Arial', 11)).pack(side='left', padx=10)
            
            self.top.protocol("WM_DELETE_WINDOW", self.skip)
            
            # 居中显示
            self.top.update_idletasks()
            x = (self.top.winfo_screenwidth() // 2) - (self.top.winfo_width() // 2)
            y = (self.top.winfo_screenheight() // 2) - (self.top.winfo_height() // 2)
            self.top.geometry(f"+{x}+{y}")
            
            # 设置焦点
            self.combo.focus_set()
        
        def ok(self):
            selected = self.combo_var.get()
            if selected and selected != '':
                # 找到对应的项目
                for i, option in enumerate(self.project_options[1:], start=0):
                    if option == selected:
                        self.result = self.projects[i]
                        break
            self.top.destroy()
        
        def skip(self):
            self.result = None
            self.top.destroy()
    
    # 创建主窗口
    root = tk.Tk()
    root.withdraw()
    
    # 确保主窗口初始化完成
    root.update()
    
    # 创建对话框
    dialog = ProjectSelectionDialog(root, extracted_name, project_config)
    
    # 等待对话框关闭
    root.wait_window(dialog.top)
    
    result = dialog.result
    root.destroy()
    
    return result

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="Select Folder with Excel Files")
    return folder_path


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
    """终端选择处理范围：全部或最新（默认最新）。"""
    print("\n请选择处理范围：")
    print("  1) 处理 A_原始数据 下所有月份文件夹")
    print("  2) 仅处理最新月份文件夹（默认）")
    choice = input("请输入选项 [回车=2]: ").strip()

    if choice == "1":
        return month_folders

    # 默认最新：优先取最后一个可识别年月的目录
    with_ym = [x for x in month_folders if x[2] is not None]
    if with_ym:
        return [with_ym[-1]]
    return [month_folders[-1]] if month_folders else []


def list_source_excel_files(folder_path):
    """列出待合并的源Excel文件（返回文件名，供原始数据汇总使用）。"""
    return [
        f for f in os.listdir(folder_path)
        if (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsm'))
        and f != 'merged_excel.xlsx'
        and f != OUTPUT_MERGED_FILENAME
        and not f.startswith('~$')
        and not f.startswith('merged_excel')
        and not f.startswith('材料采购_merged')
    ]


def list_source_excel_files_recursive(folder_path):
    """递归列出待合并的源Excel文件（返回绝对路径，供标准化汇总使用）。"""
    excel_files = []
    for root, _, files in os.walk(folder_path):
        for f in files:
            lower = f.lower()
            if not (lower.endswith('.xlsx') or lower.endswith('.xls') or lower.endswith('.xlsm')):
                continue
            if f == 'merged_excel.xlsx' or f == OUTPUT_MERGED_FILENAME:
                continue
            if f.startswith('~$') or f.startswith('merged_excel') or f.startswith('材料采购_merged'):
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

def find_header_row(sheet):
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        if row[0] and str(row[0]).strip() in ['序号', 'No.', '编号']:
            row_values = [str(val).strip() for val in row if val is not None]
            matching_cols = [col for col in standard_columns if any(col_var in row_values for col_var in column_mappings.get(col, [col]))]
            if len(matching_cols) >= len(standard_columns) // 2:
                return row_idx
    return None

def normalize_headers(headers):
    normalized = []
    for col in headers:
        stripped_col = str(col).strip() if col else ''
        for std_col, variations in column_mappings.items():
            if stripped_col in variations:
                normalized.append(std_col)
                break
        else:
            normalized.append(stripped_col)
    return normalized

def is_summary_row(formula_sheet, value_sheet, row_idx, headers, amount_col):
    """
    判断是否为汇总行的逻辑：
    检查含税金额列是否为纵向SUM公式（对该列的多行求和，如 SUM(L:L) 或 SUM(L2:L100)）
    排除横向求和（如 SUM(I85:J85) 这种同一行不同列的求和）
    """
    # 检查含税金额列的公式
    amount_cell = formula_sheet.cell(row_idx, amount_col)
    if amount_cell.data_type == 'f':
        formula_str = str(amount_cell.value).upper()
        
        # 获取当前列的列字母
        amount_col_letter = get_column_letter(amount_col)
        
        # 检查是否是对该列求和的公式
        if 'SUM' in formula_str:
            # 调试输出：打印公式内容
            print(f"  [调试] 第{row_idx}行含税金额列公式: {amount_cell.value}")
            
            import re
            
            # 检查整列求和：L:L
            if f'{amount_col_letter}:{amount_col_letter}' in formula_str:
                print(f"  → 识别为整列求和 ({amount_col_letter}:{amount_col_letter})")
                return True
            
            # 检查纵向范围求和：L2:L100（起止列相同，表示纵向）
            # 匹配模式：$?L$?数字:$?L$?数字
            range_pattern = f'\\$?{amount_col_letter}\\$?\\d+:\\$?{amount_col_letter}\\$?\\d+'
            match = re.search(range_pattern, formula_str)
            if match:
                print(f"  → 识别为纵向范围求和 (匹配: {match.group()})")
                return True
            
            # 如果是横向求和（如 SUM(I85:J85)），不同列的求和，不算汇总行
            # 这种情况公式中不会有 amount_col_letter:amount_col_letter 的模式
            print(f"  → 非纵向求和，可能是横向计算")
    
    return False

# ============= 单位标准化函数 =============
# 单位映射规则：将各种变体统一为标准单位
UNIT_MAPPING = {
    # 面积单位
    'm2': '㎡',
    'M2': '㎡',
    '平方米': '㎡',
    '平米': '㎡',
    '平方': '㎡',
    '平': '㎡',
    
    # 重量单位（千克）
    'KG': 'kg',
    '公斤': 'kg',
    'KG-袋': 'kg',
    
    # 重量单位（吨）
    't': '吨',
    'dun': '吨',

    
    # 长度单位
    'm': '米',
    'M': '米',
    
    # 体积单位
    '立方': 'm³',
    
    # 数量单位
    'PC': '个',
    'PCS': '个',
    '(个)': '个',
    '个)': '个',
    
    # 其他括号不规范的单位
    '(箱)': '箱',
    '箱)': '箱',
    '(根)': '根',
    '根)': '根',
    '片)': '片',
    
    # 明显的错误
    'j卷': '卷',
    '快': '块',
    
    # 复合单位简化
    '桶/7KG': '桶',
    '桶/20KG': '桶',
    '捆/10KG': '捆',
}

def standardize_unit(unit_str):
    """
    标准化单位字符串
    返回标准化后的单位，如果无法识别则返回原值
    """
    if not unit_str or str(unit_str).strip() == '':
        return ''
    
    unit_str = str(unit_str).strip()
    
    # 检查是否是明显的错误数据（纯数字、单个异常字符等）
    try:
        float(unit_str)
        return ''  # 如果是数字，返回空（表示无效单位）
    except ValueError:
        pass
    
    # 过滤明显错误的单位（单个非常见字符）
    if unit_str in ['上', '想', '和', '干', '寸']:
        return ''
    
    # 检查映射表
    if unit_str in UNIT_MAPPING:
        return UNIT_MAPPING[unit_str]
    
    # 返回原值（已经是标准单位或未知单位）
    return unit_str

# ============= 日期标准化函数 =============
def parse_date(date_str):
    """
    解析各种格式的日期字符串，返回 (year, month, day) 元组
    支持格式：
    - datetime 对象（openpyxl data_only=True 读取日期格式单元格时返回）
    - 2024.1, 2024.01 → (2024, 1, None)
    - 2024.1.15, 2024.01.15 → (2024, 1, 15)
    - 2025/11/13 00:00:00 → (2025, 11, 13)
    - 2025/11/13 → (2025, 11, 13)
    - 2025-11-13 → (2025, 11, 13)
    """
    if date_str is None:
        return (None, None, None)

    # 直接处理 datetime/date 对象（openpyxl data_only=True 读取日期单元格时常见）
    if isinstance(date_str, datetime):
        return (date_str.year, date_str.month, date_str.day)
    from datetime import date as date_type
    if isinstance(date_str, date_type):
        return (date_str.year, date_str.month, date_str.day)

    if str(date_str).strip() == '':
        return (None, None, None)

    date_str = str(date_str).strip()

    # 模式1：YYYY/M/D [HH:MM:SS] (带斜杠，可能带时间)
    match = re.match(r'^(\d{4})/(\d{1,2})/(\d{1,2})(?:\s+\d{1,2}:\d{2}:\d{2})?', date_str)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))
        
        # 验证月份和日期的合理性
        if 1 <= month <= 12 and 1 <= day <= 31:
            return (year, month, day)

    # 模式2：YYYY-MM-DD [HH:MM:SS]（str(datetime) 的默认格式）
    match = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})(?:[\s T]\d{1,2}:\d{2}:\d{2})?', date_str)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return (year, month, day)

    # 模式3：YYYY.M.D 或 YYYY.MM.DD 或 YYYY.M 或 YYYY.MM (带点)
    match = re.match(r'^(\d{4})\.(\d{1,2})(?:\.(\d{1,2}))?', date_str)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3)) if match.group(3) else None
        
        # 验证月份和日期的合理性
        if 1 <= month <= 12:
            if day is None or (1 <= day <= 31):
                return (year, month, day)
    
    return (None, None, None)

def infer_date_for_project(project_data):
    """
    为同一项目内的空白日期根据前后行推断填充
    project_data: [(row_idx, year, month, day), ...]
    """
    result = list(project_data)
    
    # 正向填充：从前往后，如果遇到空日期，从前面的行复制
    for i in range(1, len(result)):
        row_idx_i, year_i, month_i, day_i = result[i]
        if year_i is None:
            # 往前找到最近的非空日期
            for j in range(i - 1, -1, -1):
                row_idx_j, year_j, month_j, day_j = result[j]
                if year_j is not None:
                    result[i] = (row_idx_i, year_j, month_j, day_j)
                    break
    
    # 反向填充：从后往前，如果还有空日期，从后面的行复制
    for i in range(len(result) - 2, -1, -1):
        row_idx_i, year_i, month_i, day_i = result[i]
        if year_i is None:
            # 往后找到最近的非空日期
            for j in range(i + 1, len(result)):
                row_idx_j, year_j, month_j, day_j = result[j]
                if year_j is not None:
                    result[i] = (row_idx_i, year_j, month_j, day_j)
                    break
    
    return result

def format_date_str(year, month, day):
    """格式化日期为标准字符串"""
    if year is None:
        return ''
    if day is None:
        return f'{year}-{month:02d}'
    return f'{year}-{month:02d}-{day:02d}'

def fix_unit_quantity_swap(ws):
    """
    检测并纠正单位和数量列填反的情况
    逻辑：单位应该是文本，数量应该是数字
    如果发现单位是数字而数量是文本，则交换它们
    """
    # 找到单位和数量列
    header_row = ws[1]
    unit_col_idx = None
    quantity_col_idx = None
    
    for idx, cell in enumerate(header_row, start=1):
        # 优先查找"标准单位"，其次查找"单位"（不含"标准"前缀）
        if cell.value and '标准单位' in str(cell.value):
            unit_col_idx = idx
        elif cell.value and '单位' in str(cell.value) and '标准' not in str(cell.value) and unit_col_idx is None:
            unit_col_idx = idx
        if cell.value and '数量' in str(cell.value):
            quantity_col_idx = idx
    
    if not unit_col_idx or not quantity_col_idx:
        print("  未找到单位或数量列，跳过纠错")
        return
    
    swap_count = 0
    for row_idx in range(2, ws.max_row + 1):
        unit_cell = ws.cell(row_idx, unit_col_idx)
        quantity_cell = ws.cell(row_idx, quantity_col_idx)
        
        unit_value = unit_cell.value
        quantity_value = quantity_cell.value
        
        # 跳过空行
        if unit_value is None and quantity_value is None:
            continue
        
        # 检查是否填反：单位是数字且数量是文本
        unit_is_number = isinstance(unit_value, (int, float))
        quantity_is_text = isinstance(quantity_value, str) and quantity_value.strip() and not quantity_value.replace('.', '', 1).replace('-', '', 1).isdigit()
        
        if unit_is_number and quantity_is_text:
            # 交换单位和数量
            unit_cell.value, quantity_cell.value = quantity_cell.value, unit_cell.value
            swap_count += 1
    
    if swap_count > 0:
        print(f"  已纠正 {swap_count} 行的单位和数量填反问题")
    else:
        print("  未发现单位和数量填反的情况")

def add_standardized_unit_column(ws):
    """
    在单位列后添加标准单位列（使用手动右移方式，避免公式循环引用）
    """
    # 找到单位列
    header_row = ws[1]
    unit_col_idx = None
    
    for idx, cell in enumerate(header_row, start=1):
        if cell.value and '单位' in str(cell.value) and '标准' not in str(cell.value):
            unit_col_idx = idx
            break
    
    if not unit_col_idx:
        print("  未找到单位列，跳过单位标准化")
        return
    
    print(f"  单位列位置：第 {unit_col_idx} 列")
    
    # 在单位列后插入标准单位列
    new_col_idx = unit_col_idx + 1
    
    # 获取当前最后一列
    max_col = ws.max_column
    
    # 手动右移列：从最后一列开始往右移动1列（保留公式）
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(max_col, new_col_idx - 1, -1):
            src_cell = ws.cell(row_idx, col_idx)
            dst_cell = ws.cell(row_idx, col_idx + 1)
            
            # 直接复制值和公式，openpyxl会自动处理列号调整
            dst_cell.value = src_cell.value
            
            # 复制格式属性
            dst_cell.number_format = src_cell.number_format
            dst_cell.font = src_cell.font.copy() if src_cell.font else None
            dst_cell.fill = src_cell.fill.copy() if src_cell.fill else None
            dst_cell.border = src_cell.border.copy() if src_cell.border else None
            dst_cell.alignment = src_cell.alignment.copy() if src_cell.alignment else None
    
    # 设置新列的表头
    ws.cell(1, new_col_idx).value = '标准单位'
    ws.cell(1, new_col_idx).alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    
    # 填充标准化的单位
    standardized_count = 0
    error_count = 0
    
    for row_idx in range(2, ws.max_row + 1):
        unit_value = ws.cell(row_idx, unit_col_idx).value
        standardized_unit = standardize_unit(unit_value)
        
        ws.cell(row_idx, new_col_idx).value = standardized_unit
        
        # 统计标准化情况
        if unit_value and str(unit_value).strip():
            if standardized_unit != str(unit_value).strip():
                if standardized_unit == '':
                    error_count += 1
                else:
                    standardized_count += 1
    
    print(f"  ✓ 已添加标准单位列（第 {new_col_idx} 列）")
    print(f"  标准化修正：{standardized_count} 个单位")
    print(f"  错误单位清除：{error_count} 个")

def standardize_dates_in_workbook(output_wb):
    """
    对workbook进行日期标准化处理
    在签订日期后插入标准日期、年、月三列，并对空白日期进行推断
    直接修改workbook对象，不涉及文件I/O
    """
    print("\n=== 开始日期标准化处理 ===")
    
    ws = output_wb.active
    
    # 找到关键列的索引
    header_row = ws[1]
    date_col_idx = None
    project_name_col_idx = None
    
    for idx, cell in enumerate(header_row, start=1):
        if cell.value and '签订日期' in str(cell.value):
            date_col_idx = idx
        if cell.value and ('项目名称' in str(cell.value) or '标准项目名称' in str(cell.value)):
            # 优先使用标准项目名称，其次使用项目名称
            if '标准项目名称' in str(cell.value) or project_name_col_idx is None:
                project_name_col_idx = idx
    
    if not date_col_idx:
        print("未找到签订日期列")
        return
    
    if not project_name_col_idx:
        print("未找到项目名称列")
        return
    
    print(f"签订日期列：{date_col_idx}，项目名称列：{project_name_col_idx}")
    
    # 第一遍扫描：解析所有日期
    print("第一遍扫描：解析日期...")
    date_data = {}  # row_idx → (year, month, day)
    project_rows = {}  # project_name → [row_indices]
    
    for row_idx in range(2, ws.max_row + 1):
        project_name = ws.cell(row_idx, project_name_col_idx).value
        # 直接从ws读取日期值（此时ws已加载了data_only=True的值）
        date_str = ws.cell(row_idx, date_col_idx).value
        
        year, month, day = parse_date(date_str)
        date_data[row_idx] = (year, month, day)
        
        # 按项目分组行
        if project_name not in project_rows:
            project_rows[project_name] = []
        project_rows[project_name].append(row_idx)
    
    print(f"总行数：{ws.max_row - 1}，总项目数：{len(project_rows)}")
    
    # 第二遍：对每个项目进行日期推断
    print("第二遍：对每个项目进行日期推断...")
    inferred_dates = {}  # row_idx → (year, month, day)
    
    for project_name, row_indices in project_rows.items():
        # 收集该项目的日期数据
        project_data = [(row_idx, date_data[row_idx][0], date_data[row_idx][1], date_data[row_idx][2]) 
                        for row_idx in row_indices]
        
        # 推断该项目的日期
        inferred = infer_date_for_project(project_data)
        
        # 统计项目的推断情况
        original_empty = sum(1 for _, y, m, d in project_data if y is None)
        inferred_empty = sum(1 for _, y, m, d in inferred if y is None)
        inferred_count = original_empty - inferred_empty
        
        if inferred_count > 0:
            project_display = str(project_name)[:30] if project_name else '(未命名)'
            print(f"  项目 '{project_display}...': 推断了 {inferred_count} 个日期")
        
        for row_idx, year, month, day in inferred:
            inferred_dates[row_idx] = (year, month, day)
    
    # 插入新列
    print("插入新列（标准日期、年月）...")
    
    # 在签订日期后面插入两列
    new_col_start = date_col_idx + 1
    
    # 获取当前最后一列
    max_col = ws.max_column
    
    # 右移列：从最后一列开始往右移动2列（保留公式）
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(max_col, new_col_start - 1, -1):
            src_cell = ws.cell(row_idx, col_idx)
            dst_cell = ws.cell(row_idx, col_idx + 2)
            
            # 直接复制值和公式，openpyxl会自动处理列号调整
            dst_cell.value = src_cell.value
            
            # 复制格式属性
            dst_cell.number_format = src_cell.number_format
            dst_cell.font = src_cell.font.copy() if src_cell.font else None
            dst_cell.fill = src_cell.fill.copy() if src_cell.fill else None
            dst_cell.border = src_cell.border.copy() if src_cell.border else None
            dst_cell.alignment = src_cell.alignment.copy() if src_cell.alignment else None
    
    # 添加新列的表头
    ws.cell(1, new_col_start).value = '标准日期'
    ws.cell(1, new_col_start + 1).value = '年月'
    
    # 设置表头格式
    for col_idx in range(new_col_start, new_col_start + 2):
        cell = ws.cell(1, col_idx)
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    
    # 填充数据
    print("填充日期数据...")
    empty_count_before = sum(1 for row_idx in range(2, ws.max_row + 1) if date_data[row_idx][0] is None)
    empty_count_after = 0
    
    for row_idx in range(2, ws.max_row + 1):
        year, month, day = inferred_dates[row_idx]
        
        # 标准日期列
        ws.cell(row_idx, new_col_start).value = format_date_str(year, month, day)
        
        # 年月列（格式：YYYY-MM）
        if year and month:
            ws.cell(row_idx, new_col_start + 1).value = f'{year}-{month:02d}'
        else:
            ws.cell(row_idx, new_col_start + 1).value = ''
        
        if year is None:
            empty_count_after += 1
    
    print(f"日期填充完成：原始空白行数 {empty_count_before} → {empty_count_after}")
    print(f"推断填充的日期数：{empty_count_before} - {empty_count_after}")
    
    # 自动纠正单位和数量填反的问题
    print("\n检查并纠正单位和数量填反的情况...")
    fix_unit_quantity_swap(ws)
    
    # 注：标准单位列已在merge阶段添加，无需再添加
    
    # 更新筛选范围（包含新增的列）
    last_col = ws.max_column
    last_row = ws.max_row
    ws.auto_filter.ref = f'A1:{get_column_letter(last_col)}{last_row}'
    
    print(f"\n日期标准化完成！")
    print(f"新增列：")
    print(f"  - 第 {new_col_start} 列：标准日期")
    print(f"  - 第 {new_col_start + 1} 列：年月（YYYY-MM）")
    print("=" * 50)
    
    return new_col_start


def finalize_output_layout(output_wb):
    """统一整理导出表头、列顺序和对齐格式。"""
    source_ws = output_wb.active
    header_values = []
    for cell in source_ws[1]:
        value = '' if cell.value is None else str(cell.value).strip()
        if value == '标准项目名称':
            value = '项目名称'
            cell.value = value
        header_values.append(value)

    header_to_col = {name: idx for idx, name in enumerate(header_values, start=1) if name}
    missing_columns = [col for col in FINAL_OUTPUT_COLUMNS if col not in header_to_col]
    if missing_columns:
        print(f"警告：输出表缺少以下列，无法完整重排：{missing_columns}")
        return

    target_ws = output_wb.create_sheet('合并数据_整理', 0)

    for target_col, column_name in enumerate(FINAL_OUTPUT_COLUMNS, start=1):
        source_col = header_to_col[column_name]
        source_header_cell = source_ws.cell(1, source_col)
        target_header_cell = target_ws.cell(1, target_col)
        target_header_cell.value = column_name
        target_header_cell.font = copy(source_header_cell.font)
        target_header_cell.fill = copy(source_header_cell.fill)
        target_header_cell.border = copy(source_header_cell.border)
        target_header_cell.number_format = source_header_cell.number_format
        target_header_cell.protection = copy(source_header_cell.protection)
        target_header_cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        source_letter = get_column_letter(source_col)
        target_letter = get_column_letter(target_col)
        target_ws.column_dimensions[target_letter].width = source_ws.column_dimensions[source_letter].width

    for row_idx in range(2, source_ws.max_row + 1):
        for target_col, column_name in enumerate(FINAL_OUTPUT_COLUMNS, start=1):
            source_col = header_to_col[column_name]
            source_cell = source_ws.cell(row_idx, source_col)
            target_cell = target_ws.cell(row_idx, target_col)
            target_cell.value = source_cell.value
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)

            if column_name in CENTER_ALIGN_COLUMNS:
                target_cell.alignment = Alignment(horizontal='center', vertical='center')
            elif column_name in LEFT_ALIGN_COLUMNS:
                target_cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                target_cell.alignment = copy(source_cell.alignment)

    target_ws.row_dimensions[1].height = 40
    target_ws.freeze_panes = 'A2'
    target_ws.auto_filter.ref = f"A1:{get_column_letter(target_ws.max_column)}{target_ws.max_row}"

    output_wb.remove(source_ws)
    target_ws.title = '合并数据'


# 标准化流水线脚本所在目录
STANDARDIZATION_PIPELINE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '物品名称规格标准化')


def apply_standardization_pipeline(excel_path):
    """加载物品名称规格标准化流水线，对已保存的 Excel 就地运行所有默认规则。"""
    import importlib.util
    from pathlib import Path

    pipeline_dir = Path(STANDARDIZATION_PIPELINE_DIR)
    pipeline_script = pipeline_dir / '0_统一标准化流水线.py'

    if not pipeline_script.exists():
        print(f"\n⚠️  未找到标准化流水线脚本，跳过: {pipeline_script}")
        return

    # 动态加载流水线模块
    spec = importlib.util.spec_from_file_location('_standardization_pipeline', pipeline_script)
    if spec is None or spec.loader is None:
        print(f'\n⚠️  无法加载流水线模块: {pipeline_script}')
        return
    pipeline_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(pipeline_module)

    rule_names = getattr(pipeline_module, 'DEFAULT_RULES', [])
    if not rule_names:
        print('\n⚠️  流水线中 DEFAULT_RULES 为空，跳过标准化')
        return

    rule_files = [pipeline_dir / name for name in rule_names]

    print(f'\n{"=" * 60}')
    print('开始物品名称规格标准化流水线')
    print(f'目标文件: {Path(excel_path).name}')
    print(f'规则数量: {len(rule_files)}')
    print('=' * 60)

    workbook = openpyxl.load_workbook(str(excel_path))

    for rule_file in rule_files:
        if not rule_file.exists():
            print(f'  ⚠️  规则文件不存在，跳过: {rule_file.name}')
            continue
        try:
            rule_module = pipeline_module.load_module_from_file(rule_file)
            stats = pipeline_module.run_rule_on_workbook(rule_module, workbook, rule_file.name)
            print(
                f'  ✓ {rule_file.name}: '
                f'命中 {stats["matched"]} 行, 更新 {stats["updated"]} 行, '
                f'未识别 {stats["unresolved"]} 行'
            )
        except Exception as e:
            print(f'  ✗ 规则 {rule_file.name} 执行失败: {e}')

    workbook.save(str(excel_path))
    print(f'\n标准化完成，已覆盖保存: {excel_path}')
    print('=' * 60)


def merge_excels(folder_path, project_config, allow_batch_prompt=False):
    excel_files = list_source_excel_files(folder_path)
    
    # ===== 诊断日志初始化 =====
    processing_log = {
        'total_files': len(excel_files),
        'processed_files': [],
        'skipped_files': [],
        'error_files': [],
        'total_rows_processed': 0
    }
    
    print("\n" + "=" * 80)
    print(f"【诊断模式】开始处理文件夹: {folder_path}")
    print(f"【诊断模式】发现 {len(excel_files)} 个Excel文件")
    print("=" * 80 + "\n")
    
    if len(excel_files) == 0:
        print("⚠️  警告：未找到任何Excel文件！")
        return None, None
    
    # 列出所有文件
    print("【文件列表】:")
    for idx, file in enumerate(excel_files, 1):
        print(f"  {idx}. {file}")
    print()
    
    if PROJECT_MATCH_DB is not None:
        print(
            f"已载入共享匹配库: 别名 {len(PROJECT_MATCH_DB.get('alias_to_standard', {}))} 个，"
            f"留空 {len(PROJECT_MATCH_DB.get('unmatched_aliases', []))} 个"
        )
    
    # 第一遍扫描：收集所有需要匹配的项目
    unmatched_projects = []
    project_match_results = {}
    
    print("\n=== 第一遍扫描：提取项目名称 ===")
    for file in excel_files:
        extracted_name = extract_project_name(file)
        print(f"  文件: {file} → 项目名: '{extracted_name}'")

        resolved = resolve_project_match(extracted_name, project_config)
        project_match_results[file] = resolved
        if resolved is None:
            if extracted_name not in unmatched_projects:
                unmatched_projects.append(extracted_name)
            continue

        if resolved['status'] == 'matched':
            print(f"    → 使用共享库：[{resolved['project_code']}] {resolved['project_name']}")
        else:
            print("    → 已记录为留空，直接使用原始项目名称")
    
    # 如果有未匹配的项目，批量显示匹配对话框
    if allow_batch_prompt and unmatched_projects and project_config:
        print(f"\n发现 {len(unmatched_projects)} 个未匹配项目，打开批量匹配窗口...")
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
                project_match_results.update({
                    file_name: {
                        'status': 'matched',
                        'project_code': code,
                        'project_name': std_name,
                    }
                    for file_name, result in project_match_results.items()
                    if result is None and extract_project_name(file_name) == project_name
                })
                print(f"  '{project_name}' → [{code}] {std_name}")
            else:
                if PROJECT_MATCH_DB is not None:
                    mark_unmatched(PROJECT_MATCH_DB, project_name)
                project_match_results.update({
                    file_name: {
                        'status': 'unmatched',
                        'project_code': '',
                        'project_name': project_name,
                    }
                    for file_name, result in project_match_results.items()
                    if result is None and extract_project_name(file_name) == project_name
                })
                print(f"  '{project_name}' → （未选择）")

        if PROJECT_MATCH_DB is not None:
            save_db(PROJECT_MATCH_DB_FILE, PROJECT_MATCH_DB)
    
    print(f"\n" + "=" * 80)
    print("【开始合并Excel文件】")
    print("=" * 80 + "\n")
    
    output_wb = openpyxl.Workbook()
    merged_ws = output_wb.active
    merged_ws.title = '合并数据'
    
    # Write header - 添加项目编码和项目名称列
    header = ['项目编码', '项目名称'] + standard_columns
    merged_ws.append(header)
    
    current_row = 2  # Next row after header
    
    project_totals = defaultdict(float)
    summary_notes = {}
    
    for file_idx, file in enumerate(excel_files, 1):
        print(f"\n{'='*80}")
        print(f"【处理文件 {file_idx}/{len(excel_files)}】: {file}")
        print(f"{'='*80}")
        file_path = os.path.join(folder_path, file)
        
        file_log = {
            'filename': file,
            'status': 'unknown',
            'rows_processed': 0,
            'rows_total': 0,
            'rows_skipped': 0,
            'skip_reasons': [],
            'error_message': None
        }
        
        try:
            # ── Pass 1：公式版扫描（表头识别 + 汇总行预提取），用完即释放 ──
            print(f"  ✓ 正在加载文件（公式扫描）...")
            formula_wb = openpyxl.load_workbook(file_path, data_only=False)
            formula_sheet = formula_wb.active
            total_rows = formula_sheet.max_row
            print(f"  ✓ 文件加载成功，总行数: {total_rows}")
            
            header_row_idx = find_header_row(formula_sheet)
            if header_row_idx is None:
                skip_reason = "未找到表头行（序号列）"
                print(f"  ✗ 跳过文件: {skip_reason}")
                file_log['status'] = 'skipped'
                file_log['skip_reasons'].append(skip_reason)
                processing_log['skipped_files'].append(file_log)
                del formula_sheet, formula_wb
                continue
            
            print(f"  ✓ 找到表头行: 第{header_row_idx}行")
            
            # Get and normalize headers
            headers = [cell.value for cell in formula_sheet[header_row_idx]]
            headers = normalize_headers(headers)
            print(f"  ✓ 标准化表头完成")
            
            # Find indices
            try:
                amount_idx = headers.index('含税金额')  # 0-based
                amount_col = amount_idx + 1  # 1-based for sheet
                amount_col_letter = get_column_letter(amount_col)
                print(f"  ✓ 找到含税金额列: 第{amount_col}列")
            except ValueError:
                skip_reason = "未找到'含税金额'列"
                print(f"  ✗ 跳过文件: {skip_reason}")
                file_log['status'] = 'skipped'
                file_log['skip_reasons'].append(skip_reason)
                processing_log['skipped_files'].append(file_log)
                del formula_sheet, formula_wb
                continue
            
            # Ensure all core required columns present
            missing_core_cols = [col for col in CORE_REQUIRED_COLUMNS if col not in headers]
            if missing_core_cols:
                skip_reason = f"缺少核心必需列: {', '.join(missing_core_cols)}"
                print(f"  ✗ 跳过文件: {skip_reason}")
                file_log['status'] = 'skipped'
                file_log['skip_reasons'].append(skip_reason)
                processing_log['skipped_files'].append(file_log)
                del formula_sheet, formula_wb
                continue
            
            # Check for optional columns (允许缺少，但需要记录)
            missing_optional_cols = [col for col in OPTIONAL_COLUMNS if col not in headers]
            if missing_optional_cols:
                print(f"  ⚠️  缺少可选列: {', '.join(missing_optional_cols)}（继续处理）")

            # 预扫描汇总行：一次性从公式版提取，随后释放公式 workbook
            print(f"  ⏳ 预扫描汇总行...")
            summary_rows = set()
            for _row_idx in range(header_row_idx + 1, total_rows + 1):
                if is_summary_row(formula_sheet, None, _row_idx, headers, amount_col):
                    summary_rows.add(_row_idx)
            if summary_rows:
                print(f"  ✓ 检测到 {len(summary_rows)} 个汇总行: {sorted(summary_rows)}")
            del formula_sheet, formula_wb  # 释放公式 workbook，降低内存占用

            # ── Pass 2：数值版读取实际计算结果 ──
            print(f"  ✓ 正在加载文件（数值读取）...")
            value_wb = openpyxl.load_workbook(file_path, data_only=True)
            value_sheet = value_wb.active

            # 获取项目名称（从第一遍扫描的映射中获取）
            extracted_name = extract_project_name(file)
            matched_info = project_match_results.get(file)
            if matched_info is None:
                matched_info = resolve_project_match(extracted_name, project_config)
            
            # 设置项目信息
            project_name = extracted_name
            if matched_info and matched_info.get('status') == 'matched':
                project_code = matched_info.get('project_code', '')
                standard_project_name = matched_info.get('project_name', '')
                print(f"  ✓ 提取项目名: '{extracted_name}'")
                print(f"  ✓ 匹配结果: [{project_code}] {standard_project_name}")
            else:
                project_code = ''
                # 如果没有匹配，使用原始项目名称
                standard_project_name = extracted_name
                print(f"  ✓ 提取项目名: '{extracted_name}'")
                print(f"  ⚠️  未找到项目匹配，使用原始项目名称")
            
            excluded_summary = False
            summary_row_count = 0
            empty_row_count = 0
            blank_data_row_count = 0
            
            # First, collect all valid rows' data to determine the range for formulas
            valid_rows = []
            file_log['rows_total'] = total_rows - header_row_idx
            
            print(f"  ⏳ 开始扫描数据行...")
            for row_idx in range(header_row_idx + 1, total_rows + 1):
                row_values = [cell.value for cell in value_sheet[row_idx]]
                non_empty_count = sum(1 for val in row_values if val is not None and str(val).strip() != '')
                
                # 跳过完全空行
                if non_empty_count == 0:
                    empty_row_count += 1
                    continue
                
                # 检查是否为汇总行
                # 检查是否为汇总行（使用预扫描结果，无需再持有公式 workbook）
                if row_idx in summary_rows:
                    excluded_summary = True
                    summary_row_count += 1
                    print(f"    [跳过] 第{row_idx}行 - 汇总行")
                    continue
                
                # 检查是否为"空白数据行"
                # 即：关键采购数据列都是空（除了序号）
                # 关键列：物品名称、规格、单位、数量、单价、不含税金额等
                key_data_cols = ['物品名称', '规格', '单位', '数量', '单价', '不含税金额']
                key_data_empty = True
                for col_name in key_data_cols:
                    try:
                        col_idx = headers.index(col_name)
                        cell_value = value_sheet.cell(row_idx, col_idx + 1).value
                        if cell_value is not None and str(cell_value).strip() != '':
                            key_data_empty = False
                            break
                    except ValueError:
                        continue
                
                if key_data_empty:
                    blank_data_row_count += 1
                    continue
                
                valid_rows.append(row_idx)
            
            file_log['rows_skipped'] = empty_row_count + summary_row_count + blank_data_row_count
            file_log['rows_processed'] = len(valid_rows)
            print(f"  ⏳ 开始复制数据到合并表...")
            file_start_row = current_row
            
            # 金额列名称（需要保留2位小数）
            amount_columns = ['单价', '不含税金额', '7%增值税', '含税金额']
            
            for local_idx, row_idx in enumerate(valid_rows):
                local_row = file_start_row + local_idx
                
                # Write project code, standard project name
                merged_ws.cell(local_row, 1).value = project_code
                merged_ws.cell(local_row, 2).value = standard_project_name
                
                # Copy calculated values (不保留原有公式，直接使用计算结果)
                for col_idx, col_name in enumerate(standard_columns, start=0):
                    target_col = col_idx + 3  # After project_code, standard_project_name
                    
                    # 如果列存在于该文件的表头中，则复制数据；否则留空
                    if col_name == '标准单位':
                        # 特殊处理：标准单位从原始单位列标准化得出
                        if '单位' in headers:
                            source_col = headers.index('单位') + 1  # 1-based
                            unit_value = value_sheet.cell(row_idx, source_col).value
                            value = standardize_unit(unit_value)
                        else:
                            value = None
                    elif col_name in headers:
                        source_col = headers.index(col_name) + 1  # 1-based
                        value = value_sheet.cell(row_idx, source_col).value
                    else:
                        value = None  # 列不存在，用None填充
                    
                    # 处理金额列的格式（保留2位小数）
                    if col_name in amount_columns and isinstance(value, (int, float)):
                        # 先存储数值，格式化留在后续处理
                        merged_ws.cell(local_row, target_col).value = value
                        merged_ws.cell(local_row, target_col).number_format = '0.00'
                    else:
                        merged_ws.cell(local_row, target_col).value = value
                
                # Add to total using value
                amount_value = value_sheet.cell(row_idx, amount_col).value
                if isinstance(amount_value, (int, float)):
                    key = (project_code, standard_project_name or project_name)
                    project_totals[key] += amount_value
            
            if excluded_summary:
                summary_notes[project_code] = f"Excluded summary row from file {file}"
            
            current_row = file_start_row + len(valid_rows)
            
            # 记录成功处理的文件
            file_log['status'] = 'success'
            processing_log['processed_files'].append(file_log)
            processing_log['total_rows_processed'] += len(valid_rows)
            
            print(f"  ✅ 文件处理成功！已复制 {len(valid_rows)} 行数据")
                
        except Exception as e:
            error_msg = f"{type(e).__name__}: {str(e)}"
            print(f"  ✗ 错误: {error_msg}")
            file_log['status'] = 'error'
            file_log['error_message'] = error_msg
            processing_log['error_files'].append(file_log)
    # ===== 生成诊断报告 =====
    print("\n\n" + "=" * 80)
    print("【诊断报告】")
    print("=" * 80)
    print(f"\n📊 文件处理统计:")
    print(f"  • 总文件数: {processing_log['total_files']}")
    print(f"  • ✅ 成功处理: {len(processing_log['processed_files'])} 个")
    print(f"  • ⚠️  跳过文件: {len(processing_log['skipped_files'])} 个")
    print(f"  • ❌ 错误文件: {len(processing_log['error_files'])} 个")
    print(f"  • 📝 总数据行: {processing_log['total_rows_processed']} 行")
    
    if processing_log['processed_files']:
        print(f"\n✅ 成功处理的文件 ({len(processing_log['processed_files'])} 个):")
        for log in processing_log['processed_files']:
            skip_info = f" (跳过: {', '.join(log['skip_reasons'])})" if log['skip_reasons'] else ""
            print(f"  ✓ {log['filename']}")
            print(f"    └─ 处理 {log['rows_processed']}/{log['rows_total']} 行{skip_info}")
    
    if processing_log['skipped_files']:
        print(f"\n⚠️  跳过的文件 ({len(processing_log['skipped_files'])} 个):")
        for log in processing_log['skipped_files']:
            print(f"  ⊘ {log['filename']}")
            print(f"    └─ 原因: {'; '.join(log['skip_reasons'])}")
    
    if processing_log['error_files']:
        print(f"\n❌ 处理失败的文件 ({len(processing_log['error_files'])} 个):")
        for log in processing_log['error_files']:
            print(f"  ✗ {log['filename']}")
            print(f"    └─ 错误: {log['error_message']}")
    
    print("\n" + "=" * 80 + "\n")
    
    if current_row == 2:
        print("⚠️  没有找到任何有效数据！")
        return None, None
    
    # 设置合并数据表的格式
    # 1. 首行高度40
    merged_ws.row_dimensions[1].height = 40
    
    # 2. 首行自动换行和居中对齐
    for cell in merged_ws[1]:
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    
    # 3. 冻结首行
    merged_ws.freeze_panes = 'A2'
    
    # 4. 启用筛选（从A1到最后一列的第1行）
    last_col = get_column_letter(len(header))
    merged_ws.auto_filter.ref = f'A1:{last_col}{current_row - 1}'
    
    # 对金额列应用2位小数格式
    amount_col_names = ['单价', '不含税金额', '7%增值税', '含税金额']
    for col_idx, col_name in enumerate(standard_columns, start=3):
        if col_name in amount_col_names:
            for row_idx in range(2, current_row):
                merged_ws.cell(row_idx, col_idx).number_format = '0.00'
    
    # Create summary sheet
    summary_ws = output_wb.create_sheet('项目汇总')
    summary_ws.append(['项目编码', '项目名称', '总采购金额', '备注'])
    
    # 按照项目编码、标准项目名称排序
    for row_idx, ((project_code_key, project_name_std), total) in enumerate(sorted(project_totals.items()), start=2):
        note = summary_notes.get(project_code_key, '')
        summary_ws.append([project_code_key, project_name_std, total, note])
        # 对汇总表的总采购金额列应用2位小数格式
        summary_ws.cell(row_idx, 3).number_format = '0.00'
    
    return output_wb, output_wb.sheetnames

def main():
    global PROJECT_MATCH_DB

    raw_root = RAW_DATA_ROOT
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
    if not targets:
        print("未选择任何待处理文件夹，程序结束。")
        return

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

    print("\n双输出模式：")
    print(f"- A_合并汇总 使用原始数据: {raw_root}")
    print(f"- B_标准化汇总 使用标准化数据: {standardized_root}")

    for name, folder_path, ym_key in targets:
        print("\n" + "#" * 100)
        print(f"开始处理月份目录: {name}")
        print("#" * 100)

        month_tag = f"{ym_key[0]}年{ym_key[1]:02d}月" if ym_key else name
        output_filename = f"{month_tag}_{OUTPUT_MERGED_FILENAME}"

        # --- A: 合并汇总（来自A_原始数据）---
        output_wb_raw, _ = merge_excels(
            folder_path,
            project_config,
            allow_batch_prompt=False,
        )
        if output_wb_raw is None:
            print(f"A_合并汇总跳过：{name} 未产生有效汇总数据")
        else:
            raw_out_dir = os.path.join(out_root, "A_合并汇总")
            os.makedirs(raw_out_dir, exist_ok=True)
            raw_output_path = os.path.join(raw_out_dir, output_filename)
            apply_unified_workbook_format(output_wb_raw)
            output_wb_raw.save(raw_output_path)
            print(f"\n已输出（A_合并汇总）: {raw_output_path}")

        # --- B: 标准化汇总（来自B_采购部标准化）---
        std_folder_path = os.path.join(standardized_root, name)
        if not os.path.isdir(std_folder_path):
            print(f"B_标准化汇总跳过：未找到对应目录 {std_folder_path}")
            continue

        output_wb_std, _ = merge_standardized_excels(std_folder_path)
        if output_wb_std is None:
            print(f"B_标准化汇总跳过：{name} 未产生有效汇总数据")
            continue

        std_out_dir = os.path.join(out_root, "B_标准化汇总")
        os.makedirs(std_out_dir, exist_ok=True)
        std_output_path = os.path.join(std_out_dir, output_filename)
        apply_unified_workbook_format(output_wb_std)
        output_wb_std.save(std_output_path)
        print(f"已输出（B_标准化汇总）: {std_output_path}")

    if PROJECT_MATCH_DB is not None:
        save_db(PROJECT_MATCH_DB_FILE, PROJECT_MATCH_DB)
        print(f"共享项目匹配库已保存: {PROJECT_MATCH_DB_FILE}")

if __name__ == "__main__":
    main()