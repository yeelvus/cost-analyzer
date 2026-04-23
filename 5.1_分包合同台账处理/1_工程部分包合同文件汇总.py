import pandas as pd
import re
import os
import sys
import importlib.util
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

DATA_ROOT = Path('/Users/wyattsun/Desktop/邯郸中材/01成本统计')
CODE_ROOT = Path('/Users/wyattsun/Desktop/code')
MODULE_DIRS = [
    DATA_ROOT / '0_项目基本信息情况',
    CODE_ROOT,
    CODE_ROOT / '0_项目基本信息',
]
for module_dir in MODULE_DIRS:
    if module_dir.exists() and str(module_dir) not in sys.path:
        sys.path.insert(0, str(module_dir))

def _load_module_from_path(module_name: str, file_path: Path):
    if not file_path.exists():
        raise FileNotFoundError(f"未找到模块文件: {file_path}")
    spec = importlib.util.spec_from_file_location(module_name, str(file_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MODULE_BASE = Path('/Users/wyattsun/Desktop/code/0_项目基本信息')
project_name_matcher = _load_module_from_path(
    'project_name_matcher',
    MODULE_BASE / 'project_name_matcher.py'
)
excel_formatting_utils = _load_module_from_path(
    'excel_formatting_utils',
    MODULE_BASE / 'excel_formatting_utils.py'
)

build_standard_projects_from_dataframe = project_name_matcher.build_standard_projects_from_dataframe
load_or_init_db = project_name_matcher.load_or_init_db
save_db = project_name_matcher.save_db
get_project_lookup = project_name_matcher.get_project_lookup
add_mapping = project_name_matcher.add_mapping
match_project = project_name_matcher.match_project
apply_common_excel_layout = excel_formatting_utils.apply_common_excel_layout

# ================= 配置区域 =================
EXCEL_INFO_PATH = Path('/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx')

# 新目录结构：5_合同管理/B_分包合同
NEW_SUBPACKAGE_ROOT = Path('/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同')
NEW_INPUT_DIR = NEW_SUBPACKAGE_ROOT / '/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/1_分包合同台账_raw'
NEW_OUTPUT_DIR = NEW_SUBPACKAGE_ROOT / '2_分包合同台账_std'

INPUT_DIR = NEW_INPUT_DIR
OUTPUT_DIR = NEW_OUTPUT_DIR

MATCH_DB_PATH = Path('/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json')
# ===========================================

def smart_process_and_match():
    print("=" * 60)
    print("🚀 批处理模式：自动扫描目录内所有月度台账并汇总")

    if not os.path.exists(EXCEL_INFO_PATH) or not os.path.exists(INPUT_DIR):
        print(f"❌ 错误：请检查路径设置是否正确！\n项目信息表: {EXCEL_INFO_PATH}\n输入目录: {INPUT_DIR}")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        # 1. 从项目基本信息表 Excel 加载配置
        info_df = pd.read_excel(EXCEL_INFO_PATH, sheet_name='在建项目')
        info_df['项目编码'] = info_df['项目编码'].fillna('需手动核对').astype(str)
        if '项目状态' in info_df.columns:
            info_df['项目状态'] = info_df['项目状态'].fillna('需手动核对').astype(str)
            # 兼容 project_name_matcher 仍读取“备注”列
            info_df['备注'] = info_df['项目状态']
        elif '备注' in info_df.columns:
            info_df['备注'] = info_df['备注'].fillna('需手动核对').astype(str)
        else:
            info_df['备注'] = '需手动核对'
        info_df = info_df.dropna(subset=['项目名称'])
        config_data = build_standard_projects_from_dataframe(info_df)
        match_db = load_or_init_db(MATCH_DB_PATH, config_data)
        project_lookup = get_project_lookup(match_db)

        # 2. 自动扫描所有台账文件（支持 xlsx/xlsm/xls）
        ledger_files = []
        for fn in os.listdir(INPUT_DIR):
            if fn.startswith('~$'):
                continue
            lower_fn = fn.lower()
            if lower_fn.endswith(('.xlsx', '.xlsm', '.xls')) and '台账' in fn:
                ledger_files.append(os.path.join(INPUT_DIR, fn))

        if not ledger_files:
            print(f"❌ 未在目录中找到可处理的台账文件：{INPUT_DIR}")
            return

        print(f"📂 检测到台账文件：{len(ledger_files)} 个")
        print(f"🗂️ 匹配库路径: {MATCH_DB_PATH}")

        # 按源文件聚合结果（一个输入文件对应一个输出文件）
        file_result_map = {}

        def extract_month_from_filename(file_path):
            file_name = os.path.basename(file_path)
            # 优先匹配 yyyy-mm-dd / yyyy_mm_dd / yyyy年m月d日
            m = re.search(r'(20\d{2})[-_年\.](\d{1,2})(?:[-_月\.](\d{1,2}))?', file_name)
            if m:
                year = m.group(1)
                month = int(m.group(2))
                return f"{year}-{month:02d}"
            return None

        # 3. 匹配逻辑函数
        def get_match_info(raw_header_name):
            if pd.isna(raw_header_name) or not str(raw_header_name).strip():
                return "未分配", "未分配", "未分配", "未分配"
            
            # 其次使用持久化匹配数据库
            db_hit = match_project(raw_header_name, match_db)
            if db_hit:
                return (
                    raw_header_name,
                    db_hit['project_name'],
                    db_hit['project_code'],
                    db_hit['target_status'],
                )
            
            # 清洗标题名称：去除括号及冗余词
            clean_n = re.sub(r'（.*?）|\(.*?\)', '', str(raw_header_name))
            clean_n = clean_n.replace("泰国", "").replace("项目", "").replace("工程", "").strip()
            
            # 与 JSON 配置进行模糊比对
            for item in config_data:
                std_name = item.get('project_name', "")
                if std_name in clean_n or clean_n in std_name:
                    # 自动匹配命中也增量沉淀到匹配库
                    add_mapping(match_db, raw_header_name, std_name)
                    return (
                        raw_header_name,
                        std_name,
                        item.get('project_code', ""),
                        item.get('target_status', "")
                    )
            return raw_header_name, "需手动核对", "需手动核对", "需手动核对"

        # 5. 核心处理流程
        processed_rows = []
        current_header_name = None
        # 扩展序号识别模式
        # 匹配：中文数字(一二三)、阿拉伯数字(1、2、3、1.、2.)、罗马数字(I、II、III)
        header_pattern = re.compile(r'^([一二三四五六七八九十百千万]+|[\dIVXLCDM]+[\.\、]?)$')

        # 定义关键数据列（用于判断是否为明细行）
        # 金额列：优先识别“分包合同额（泰铢）”，兼容历史列名
        amount_columns = [
            '分包合同额（泰铢）', '分包合同额(泰铢）', '分包合同额(泰铢)',
            '合同额（泰铢）', '合同额(泰铢）', '合同额(泰铢)',
            '合同金额', '金额', '结算总金额(泰铢）', '结算总金额（泰铢）'
        ]
        # 其他关键列
        key_columns = ['合同编号', '协作体名称', '工程范围']

        # 结算模式关键词：命中这些文本时，不应当被识别为项目标题
        settlement_keywords = [
            '包干总价', '总价包干', '总包干价', '固定包干总价',
            '固定综合单价', '综合单价', '清包工', '全包模式', '大包模式', '合同结算模式'
        ]

        def has_text(value):
            if pd.isna(value):
                return False
            text = str(value).strip()
            return text not in ('', 'nan', 'none', 'None')

        def parse_amount_number(value):
            if pd.isna(value):
                return None
            text = str(value).strip()
            if text in ('', 'nan', 'None', 'none'):
                return None
            cleaned = text.replace(',', '').replace(' ', '')
            cleaned = re.sub(r'[^\d\.-]', '', cleaned)
            if cleaned in ('', '-', '.', '-.'):
                return None
            try:
                return float(cleaned)
            except Exception:
                return None

        def format_output_excel(output_path, amount_col_name='合同额（泰铢）'):
            wb = load_workbook(output_path)
            ws = wb.active
            
            # 使用统一的格式化工具
            apply_common_excel_layout(
                ws,
                numeric_start_col=5,  # 从第5列开始应用千分位格式（金额相关列）
                header_row=1,
                header_height=40,
                freeze_panes='A2',
                thousand_number_format='#,##0.00',
                min_col_width=10,
                max_col_width=60,
                extra_padding=2,
            )
            
            wb.save(output_path)

        def find_amount_from_row(row):
            # 1) 先按显式列名优先级识别
            for col in amount_columns:
                if col in row.index and pd.notna(row[col]):
                    amt_val = parse_amount_number(row[col])
                    if amt_val is not None and amt_val != 0:
                        return row[col], col

            # 2) 再按列名关键词兜底识别，适配轻微变体
            for col in row.index:
                col_text = str(col)
                if ('分包合同额' in col_text) or ('合同额' in col_text and '主合同额' not in col_text):
                    if pd.notna(row[col]):
                        amt_val = parse_amount_number(row[col])
                        if amt_val is not None and amt_val != 0:
                            return row[col], col_text

            return None, None

        def normalize_contract_amount(new_row, preferred_value=None):
            # 合同额统一回填到“合同额（泰铢）”，避免金额落在别的列导致漏汇总
            primary = new_row.get('合同额（泰铢）')
            if has_text(primary):
                return

            if has_text(preferred_value):
                new_row['合同额（泰铢）'] = preferred_value
                return

            for alt_col in ('合同金额', '金额'):
                alt_val = new_row.get(alt_col)
                if has_text(alt_val):
                    new_row['合同额（泰铢）'] = alt_val
                    return

        def is_settlement_mode_text(text):
            return any(k in text for k in settlement_keywords)

        def is_project_title_candidate(raw_name):
            text = str(raw_name).strip()
            if not text:
                return False
            if is_settlement_mode_text(text):
                return False

            # 先尝试使用匹配库/别名库识别，兼容“建华管桩”这类不含“项目/工程”的标题。
            if match_project(text, match_db):
                return True

            # 优先按项目信息表做模糊匹配
            clean_text = re.sub(r'（.*?）|\(.*?\)', '', text)
            clean_text = clean_text.replace('泰国', '').replace('项目', '').replace('工程', '').strip()
            for item in config_data:
                std_name = item.get('project_name', '')
                clean_std = re.sub(r'（.*?）|\(.*?\)', '', str(std_name))
                clean_std = clean_std.replace('泰国', '').replace('项目', '').replace('工程', '').strip()
                if clean_std and (clean_std in clean_text or clean_text in clean_std):
                    return True

            # 未命中映射时，至少需要具备项目类语义
            return ('项目' in text or '工程' in text)

        def is_detail_row(row, amount_raw_value=None):
            if amount_raw_value is not None:
                return True

            for col in (contract_no_col, coop_col, scope_col):
                if col and col in row.index and has_text(row.get(col)):
                    return True

            mode_val = row.get(mode_col) if mode_col and mode_col in row.index else None
            if has_text(mode_val) and not is_project_title_candidate(mode_val):
                return True

            return False

        def find_col_name(columns, include_keywords, exclude_keywords=None):
            exclude_keywords = exclude_keywords or []
            for col in columns:
                col_text = str(col).strip()
                if any(k in col_text for k in include_keywords) and not any(ek in col_text for ek in exclude_keywords):
                    return col
            return None

        # 4. 逐文件处理（一个输入文件对应一个输出文件）
        for ledger_path in sorted(ledger_files):
            target_month = extract_month_from_filename(ledger_path)
            if not target_month:
                print(f"⚠️ 跳过（文件名未识别月份）: {os.path.basename(ledger_path)}")
                continue

            print(f"\n⌛ 处理文件: {os.path.basename(ledger_path)} -> 月份 {target_month}")
            excel_file = pd.ExcelFile(ledger_path)
            processed_rows = []

            # 优先主台账sheet，避免读取个人拆分sheet导致重复
            primary_sheets = [s for s in excel_file.sheet_names if '分包结算台账' in str(s)]
            sheets_to_process = primary_sheets if primary_sheets else excel_file.sheet_names

            for sheet_name in sheets_to_process:
                df = pd.read_excel(ledger_path, sheet_name=sheet_name)
                current_header_name = None

                # 按列语义自动识别关键列，兼容不同 sheet 结构
                serial_col = find_col_name(df.columns, ['序号'])
                mode_col = find_col_name(df.columns, ['合同结算模式'])
                project_col = find_col_name(df.columns, ['项目名称'])
                contract_no_col = find_col_name(df.columns, ['合同编号'])
                coop_col = find_col_name(df.columns, ['协作体名称'])
                scope_col = find_col_name(df.columns, ['工程范围'])

                for _, row in df.iterrows():
                    first_raw = row.get(serial_col) if serial_col in row.index else (row.iloc[0] if len(row) > 0 else None)
                    second_raw = row.get(mode_col) if mode_col in row.index else (row.iloc[1] if len(row) > 1 else None)
                    first_col_val = str(first_raw).strip() if pd.notna(first_raw) else ""
                    second_col_val = str(second_raw).strip() if pd.notna(second_raw) else ""

                    has_contract_meta = has_text(row.get(contract_no_col)) or has_text(row.get(coop_col))

                    # 优先读取显式项目名称列（若该列存在）
                    explicit_project_name = ""
                    if project_col and project_col in row.index and has_text(row.get(project_col)):
                        candidate = str(row.get(project_col)).strip()
                        if not is_settlement_mode_text(candidate):
                            explicit_project_name = candidate

                    # 显式项目列优先更新当前项目上下文
                    if explicit_project_name and is_project_title_candidate(explicit_project_name):
                        current_header_name = explicit_project_name

                    amount_raw_value, _ = find_amount_from_row(row)
                    has_amount = amount_raw_value is not None

                    if has_amount:
                        effective_project_name = explicit_project_name or current_header_name
                        if effective_project_name:
                            orig, standard, code, _ = get_match_info(effective_project_name)
                            new_row = row.to_dict()
                            normalize_contract_amount(new_row, preferred_value=amount_raw_value)
                            new_row['统计月份'] = target_month
                            new_row['项目名称'] = standard if standard != '需手动核对' else orig
                            new_row['项目编码'] = code
                            processed_rows.append(new_row)
                        continue

                    is_header = False
                    if (
                        header_pattern.match(first_col_val)
                        and second_col_val
                        and not has_contract_meta
                        and is_project_title_candidate(second_col_val)
                    ):
                        is_header = True
                        current_header_name = second_col_val
                    elif len(first_col_val) <= 3 and second_col_val and not has_contract_meta:
                        key_empty_count = sum(1 for col in key_columns if col in row.index and pd.isna(row[col]))
                        if key_empty_count >= len(key_columns) * 0.7 and is_project_title_candidate(second_col_val):
                            is_header = True
                            current_header_name = second_col_val if len(second_col_val) > len(first_col_val) else first_col_val
                    elif 4 <= len(first_col_val) <= 20 and not second_col_val and not has_contract_meta:
                        key_empty_count = sum(1 for col in key_columns if col in row.index and pd.isna(row[col]))
                        if key_empty_count >= len(key_columns) * 0.7 and is_project_title_candidate(first_col_val):
                            is_header = True
                            current_header_name = first_col_val

                    if is_header:
                        print(f"  识别项目标题: {current_header_name}")
                        continue

                    effective_project_name = explicit_project_name or current_header_name
                    if effective_project_name and is_detail_row(row, amount_raw_value=amount_raw_value):
                        orig, standard, code, _ = get_match_info(effective_project_name)
                        new_row = row.to_dict()
                        normalize_contract_amount(new_row, preferred_value=amount_raw_value)
                        new_row['统计月份'] = target_month
                        new_row['项目名称'] = standard if standard != '需手动核对' else orig
                        new_row['项目编码'] = code
                        processed_rows.append(new_row)

            if processed_rows:
                file_result_df = pd.DataFrame(processed_rows)
                file_result_map[ledger_path] = file_result_df
                print(f"✅ 明细行提取: {len(file_result_df)}")
            else:
                print("⚠️ 未提取到明细行")

        if not file_result_map:
            print("❌ 未生成任何月度汇总结果。")
            return

        # 将本次新增匹配沉淀到持久化数据库
        save_db(MATCH_DB_PATH, match_db)
        print(f"💾 已更新匹配库: {MATCH_DB_PATH}")

        # 6. 仅保留指定列并导出（一个源文件对应一个输出文件）
        source_columns = [
            '项目编码', '项目名称', '合同结算模式', '协作体名称', '工程范围',
            '合同额（泰铢）', '合同编号', '统计月份'
        ]

        target_columns = [
            '项目编码', '项目名称', '合同结算模式', '协作体名称', '工程范围',
            '合同额（泰铢）', '合同编号', '统计月份'
        ]

        print("\n" + "-" * 60)
        for ledger_path in sorted(file_result_map.keys()):
            source_df = file_result_map[ledger_path]
            target_month = extract_month_from_filename(ledger_path)
            for col in source_columns:
                if col not in source_df.columns:
                    source_df[col] = ''

            export_df = source_df[source_columns].copy()
            export_df = export_df[target_columns]

            # 按项目编码排序，相同项目编码聚合在一起
            export_df = export_df.sort_values(by='项目编码', kind='stable')
            export_df = export_df.reset_index(drop=True)

            # 保留原始顺序与文本样貌，不做额外数值/日期格式转换
            output_name = f"分包签订合同汇总_{target_month}" if target_month else f"分包签订合同汇总_{os.path.splitext(os.path.basename(ledger_path))[0]}"
            output_path = os.path.join(OUTPUT_DIR, f"{output_name}.xlsx")
            export_df.to_excel(output_path, index=False)
            format_output_excel(output_path)
            print(f"📁 已导出: {output_path}  (行数: {len(export_df)})")

        print("-" * 60)
        print(f"✅ 批处理完成！共输出 {len(file_result_map)} 个文件")

    except Exception as e:
        print(f"❌ 程序运行出错: {e}")

if __name__ == "__main__":
    smart_process_and_match()