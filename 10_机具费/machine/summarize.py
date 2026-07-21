# -*- coding: utf-8 -*-
"""机具使用汇总：std → sum（项目汇总 / 月份宽表 / 处理报告 / 明细）。"""
from __future__ import annotations

import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

from machine.config import (
    CODE_ROOT,
    EXCEL_FILE_OVERRIDE,
    MACHINE_MAPPING_JSON,
    MACHINE_MAPPING_TXT,
    OUTPUT_FILE_SUFFIX,
    OUTPUT_FILENAME_PREFIX,
    PROCESS_REPORT_SHEET_NAME,
    PROJECT_MATCH_DB,
    PROJECT_MONTHLY_CHANGE_SHEET_NAME,
    PROJECT_MONTHLY_WIDE_SHEET_NAME,
    PROJECT_SUMMARY_SHEET_NAME,
    RAW_SPLIT_DIR,
    REGION_MAPPING_JSON,
    REGION_MAPPING_TXT,
    RUN_MODE,
    STD_DIR,
    STD_SUM_DIR,
    SUM_DIR,
    SUMMARY_SHEET_NAME,
    ensure_data_dirs,
)
from machine.columns import (
    extract_project_info,
    extract_project_name_from_sheet,
    find_header_row,
    find_matching_column,
    find_use_project_column,
)
from machine.discover import (
    PROJECT_SPLIT_FILE_RE,
    build_project_output_basename,
    build_summary_output_basename,
    extract_project_code_and_name_from_filename,
    extract_project_name_from_filename,
    extract_year_month_from_filename,
    group_files_by_project,
    list_all_standardized_files,
    pick_latest_standardized_file,
)
from machine.mapping import (
    apply_smart_mapping_rules,
    guess_region_from_project_or_sheet,
    load_machine_mapping,
    load_project_config,
    load_project_mapping,
    load_unified_project_mapping,
    normalize_special_vehicle_kind,
    save_machine_mapping,
    save_project_mapping,
)
from machine.textutil import (
    is_month_like_text,
    normalize_lookup_text,
    normalize_year_month_text,
    safe_to_float,
)
from machine.validate import (
    detect_and_fix_supplier_month_swap,
    detect_and_fix_swapped_values,
    validate_overtime_fee,
    validate_settlement_amount,
    validate_total_amount,
)

# 输入：原始拆分；输出：标准汇总
ensure_data_dirs()
STANDARDIZED_INPUT_DIR = str(RAW_SPLIT_DIR if RAW_SPLIT_DIR.exists() else STD_DIR)
OUTPUT_FOLDER = str(STD_SUM_DIR)
PROJECT_MAPPING_JSON = str(REGION_MAPPING_JSON)
MACHINE_MAPPING_JSON_PATH = str(MACHINE_MAPPING_JSON)
PROJECT_MAPPING_TXT_LEGACY = str(REGION_MAPPING_TXT) if REGION_MAPPING_TXT.exists() else None
MACHINE_MAPPING_TXT_LEGACY = str(MACHINE_MAPPING_TXT) if MACHINE_MAPPING_TXT.exists() else None
PROJECT_MATCH_DB_PATH = str(PROJECT_MATCH_DB)

if str(CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(CODE_ROOT))

from common.project import (  # noqa: E402
    build_standard_projects_from_dataframe,
    load_or_init_db,
)


def main(excel_path_override=None, output_basename_override=None):
    if isinstance(excel_path_override, (list, tuple)) and excel_path_override:
        excel_paths = [Path(p).expanduser().resolve() for p in excel_path_override]
        print("批量模式-当前处理文件组:")
        for p in excel_paths:
            print(f"  - {p.name}")
    elif excel_path_override:
        excel_path = Path(str(excel_path_override)).expanduser().resolve()
        excel_paths = [excel_path]
        print(f"批量模式-当前处理文件: {excel_path.name}")
    elif EXCEL_FILE_OVERRIDE and str(EXCEL_FILE_OVERRIDE).strip():
        excel_path = Path(EXCEL_FILE_OVERRIDE).expanduser().resolve()
        excel_paths = [excel_path]
        print(f"使用手工指定文件: {excel_path}")
    else:
        picked = pick_latest_standardized_file(STANDARDIZED_INPUT_DIR)
        if not picked:
            print(f"找不到标准化文件：{STANDARDIZED_INPUT_DIR}")
            return
        excel_path = picked.resolve()
        excel_paths = [excel_path]
        print(f"自动选择最新标准化文件: {excel_path.name}")

    for excel_path in excel_paths:
        if not excel_path.exists():
            print(f"找不到文件：{excel_path}")
            return

    # 汇总输出命名默认使用文件组中“最新月份文件”
    primary_excel_path = excel_paths[-1]

    # 1. 加载映射表和项目配置
    project_mapping = load_project_mapping(PROJECT_MAPPING_JSON, PROJECT_MAPPING_TXT_LEGACY)          # 项目→区域
    kind_mapping, model_mapping = load_machine_mapping(MACHINE_MAPPING_JSON_PATH, MACHINE_MAPPING_TXT_LEGACY)  # 机具种类、型号
    kind_mapping = {
        normalize_lookup_text(raw): str(std).strip()
        for raw, std in kind_mapping.items()
        if normalize_lookup_text(raw)
    }
    model_mapping = {
        (normalize_lookup_text(kind), normalize_lookup_text(raw_model)): str(std_model).strip()
        for (kind, raw_model), std_model in model_mapping.items()
        if normalize_lookup_text(kind) and normalize_lookup_text(raw_model)
    }
    
    # 加载项目配置：优先使用统一的项目编码映射表
    project_config = load_unified_project_mapping()
    
    # 如果统一映射表为空，回退到旧配置
    if not project_config:
        print("统一映射表加载失败或为空，回退到旧项目配置")
        project_config = load_project_config()

    # 项目匹配仅在 merge-raw（机具费_raw合并）完成；汇总阶段不再弹窗/再匹配
    matcher_info_df = pd.DataFrame({
        '项目名称': [p.get('project_name', '') for p in project_config],
        '项目编码': [p.get('project_code', '') for p in project_config],
        '备注': [p.get('status', p.get('target_status', '需手动核对')) for p in project_config],
    })
    standard_projects = build_standard_projects_from_dataframe(matcher_info_df)
    match_db = load_or_init_db(PROJECT_MATCH_DB_PATH, standard_projects)
    print(
        f"✅ 项目主数据已加载（仅用于输出命名）：标准项目 {len(standard_projects)} 个；"
        f"汇总阶段不弹窗匹配项目"
    )

    # 用于收集新发现的机具种类、型号（不再收集/补全项目映射）
    new_projects = set()
    new_kinds = set()
    new_models = set()
    
    all_data = []
    report_data = []  # 用于生成报告

    internal_generated_sheets = {
        SUMMARY_SHEET_NAME,
        PROJECT_SUMMARY_SHEET_NAME,
        PROJECT_MONTHLY_CHANGE_SHEET_NAME,
        PROJECT_MONTHLY_WIDE_SHEET_NAME,
        PROCESS_REPORT_SHEET_NAME,
        "处理报告",
    }

    def _is_data_sheet(sheet_name: str, all_names: list[str]) -> bool:
        """只处理数据 sheet：优先「明细」；跳过月份名 sheet 与内部汇总 sheet。"""
        if sheet_name in internal_generated_sheets:
            return False
        if is_month_like_text(sheet_name):
            return False
        # 若工作簿含「明细」，只读明细，避免旧文件里按月 sheet 重复计入
        if "明细" in all_names:
            return sheet_name == "明细"
        return True
    
    # 直接处理项目明细文件（项目编码/名称已在 raw 合并阶段写好）
    for source_excel_path in excel_paths:
        file_code, file_name = extract_project_code_and_name_from_filename(source_excel_path)
        wb = openpyxl.load_workbook(source_excel_path, data_only=True)
        workbook_looks_generated = len(set(wb.sheetnames) & internal_generated_sheets) >= 2
        print(f"\n=== 处理文件: {source_excel_path.name} ===")
        print(f"  文件名解析项目: [{file_code or '—'}] {file_name or '—'}")
        sheet_names = list(wb.sheetnames)
        for sheet_name in sheet_names:
            if workbook_looks_generated and sheet_name in internal_generated_sheets:
                continue
            if not _is_data_sheet(sheet_name, sheet_names):
                print(f"  跳过 sheet：{sheet_name}（非明细/像月份名）")
                continue
        
            print(f"\n正在处理 sheet：{sheet_name}")
            ws = wb[sheet_name]
            sheet_key = f"{source_excel_path.name}::{sheet_name}"
        
            # 仅取开始月份等辅助信息，不再用 sheet 名做项目匹配
            _, start_month = extract_project_info(ws)
            print(f"  开始月份: {start_month if start_month else '未找到'}")
            
            # 查找表头行
            header_row = find_header_row(ws)
            if not header_row:
                print(f"  ⚠️  未找到表头行，跳过此sheet")
                report_data.append({
                    "sheet名称": sheet_name,
                    "项目名称": sheet_name,
                    "项目编码": '',
                    "开始月份": start_month,
                    "状态": "失败：未找到表头",
                    "数据行数": 0,
                    "加班费错误行数": 0,
                    "加班费错误率(%)": 0,
                    "加班费差异金额(泰铢)": 0,
                    "结算费用错误行数": 0,
                    "结算费用错误率(%)": 0,
                    "结算费用差异金额(泰铢)": 0,
                    "本次费用合计错误行数": 0,
                    "本次费用合计错误率(%)": 0,
                    "本次费用合计差异金额(泰铢)": 0,
                    "费用合计(泰铢)": 0
                })
                continue
            
            print(f"  表头位于第 {header_row} 行")
            
            # 使用pandas读取，从表头行开始
            # 先不指定dtype让pandas读取原始数据类型，这样有助于公式计算
            df = pd.read_excel(source_excel_path, sheet_name=sheet_name, header=header_row-1, engine='openpyxl')
            
            # 清理列名：去除换行符和多余空格
            df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
            
            # 定义标准列名及其可能的变体（用于模糊匹配）
            column_mapping = {
                "序号": ["序号"],
                # 项目字段：merge-raw 已匹配，必须保留（勿把「项目名称」吸到使用项目）
                "项目编码": ["项目编码", "编码"],
                "项目名称": ["项目名称", "标准项目名称"],
                "使用项目": ["使用项目", "使用项目名称"],
                "机具种类": ["机具种类"],
                "型号规格": ["型号规格"],
                "台班费_泰铢": ["台班费", "台班费泰铢", "台班费(泰铢)", "台班费 泰铢", "台班费（泰铢）"],
                "本次结算台班": ["本次结算台班", "结算台班", "本次结算台班（小时、天、月）", "本次结算台班(小时、天、月)", "本次结算台班 （小时、天、月)"],
                "加班小时": ["加班", "加班(小时)", "加班（小时）", "加班 (小时)"],
                "加班费": ["加班费", "加班费/超时费", "加班费（小时）", "加班费(小时)", "加班费 （小时）", "加班费 "],
                "进出场费运费": ["进出场费", "运费", "进出场费/运费", "运费＋拆装费", "运费+拆装费"],
                "移动费拆装费": ["移动费", "拆装费", "移动费＋拆装费", "移动费+拆装费", "移动费＋拆装费+接管费"],
                "结算费用_泰铢": ["结算费用", "本次结算费用", "结算费用(泰铢)", "结算费用（泰铢）", "本次结算费用(泰铢)", "结算费用 (泰铢)", "本次结算费用 (泰铢)"],
                "增值税7%": ["增值税", "增值税7%", "增值税7％"],
                "机具租赁预扣税3%": ["预扣税", "机具租赁预扣税", "机具租赁预扣税3%", "机具租赁预扣税3％", "机具租赁预 扣税"],
                "机具租赁预扣税5%": ["机具租赁预扣税5%", "机具租赁预扣税5％"],
                "本次费用合计": ["本次费用合计", "本次费用合计（泰铢）", "本次费用合计(泰铢)", "本次费用合计 （泰铢）", "费用合计", "费用 合计"],
                "使用日期": ["使用日期", "使用时间"],
                "出租单位": ["出租单位"],
                "月份": ["月份"],
                "年月": ["年月", "年-月", "月度", "期间"],
                "备注": ["备注", "项目说明"]
            }
            
            # 创建新的DataFrame，使用标准列名
            new_df = pd.DataFrame()
            unmatched_cols = []

            for std_col, patterns in column_mapping.items():
                # 项目编码/名称/使用项目：严格精确匹配，避免互相误吸
                exact_only = std_col in ("使用项目", "项目编码", "项目名称")
                matched_col = find_matching_column(df.columns, patterns, exact_only=exact_only)
                if matched_col:
                    new_df[std_col] = df[matched_col]
                    print(f"    '{matched_col}' -> '{std_col}'")
                else:
                    new_df[std_col] = ""
                    unmatched_cols.append(std_col)
            
            if unmatched_cols:
                print(f"  ⚠️  未匹配到的列: {', '.join(unmatched_cols)}")
            
            df = new_df
            
            # 重新计算费用相关列（处理公式未缓存的情况）
            # Excel公式结构：
            # 1. 结算费用 = 台班费 × 台班数 + 加班费 + 运费
            # 2. 增值税7% = 结算费用 × 0.07
            # 3. 预扣税3% = 结算费用 × 0.03
            # 4. 本次费用合计 = 结算费用 + 增值税 - 预扣税3% - 预扣税5%
            print(f"  正在检查并重新计算费用...")
            
            settlement_recalc = 0
            vat_recalc = 0
            tax_recalc = 0
            fee_recalc = 0

            def write_recalc_value(target_df, row_idx, col_name, numeric_value):
                try:
                    target_df.at[row_idx, col_name] = numeric_value
                except Exception:
                    target_df.at[row_idx, col_name] = str(numeric_value)
            
            for idx in range(len(df)):
                # 第一步：重新计算"结算费用"
                settlement = safe_to_float(df.iloc[idx]["结算费用_泰铢"])
                
                if settlement == 0:
                    # 尝试从基础数据计算
                    unit_price = safe_to_float(df.iloc[idx]["台班费_泰铢"])
                    quantity = safe_to_float(df.iloc[idx]["本次结算台班"])
                    overtime_fee = safe_to_float(df.iloc[idx]["加班费"])
                    entry_fee = safe_to_float(df.iloc[idx]["进出场费运费"])
                    move_fee = safe_to_float(df.iloc[idx]["移动费拆装费"])
                    
                    # 计算结算费用 = 台班费×台班数 + 加班费 + 各种运费
                    calc_settlement = unit_price * quantity + overtime_fee + entry_fee + move_fee
                    
                    if calc_settlement != 0:
                        settlement = calc_settlement
                        write_recalc_value(df, idx, "结算费用_泰铢", settlement)
                        settlement_recalc += 1
                
                # 第二步：重新计算"增值税7%"
                vat = safe_to_float(df.iloc[idx]["增值税7%"])
                if vat == 0 and settlement != 0:
                    calc_vat = settlement * 0.07
                    write_recalc_value(df, idx, "增值税7%", calc_vat)
                    vat = calc_vat
                    vat_recalc += 1
                
                # 第三步：重新计算"预扣税3%"
                tax3 = safe_to_float(df.iloc[idx]["机具租赁预扣税3%"])
                if tax3 == 0 and settlement != 0:
                    calc_tax3 = settlement * 0.03
                    write_recalc_value(df, idx, "机具租赁预扣税3%", calc_tax3)
                    tax3 = calc_tax3
                    tax_recalc += 1
                
                # 第四步：重新计算"本次费用合计"
                current_fee = safe_to_float(df.iloc[idx]["本次费用合计"])
                if current_fee == 0 and settlement != 0:
                    tax5 = safe_to_float(df.iloc[idx]["机具租赁预扣税5%"])
                    calculated_fee = settlement + vat - tax3 - tax5
                    write_recalc_value(df, idx, "本次费用合计", calculated_fee)
                    fee_recalc += 1
            
            if settlement_recalc > 0 or vat_recalc > 0 or tax_recalc > 0 or fee_recalc > 0:
                print(f"  ✓ 重新计算: 结算费用{settlement_recalc}行, 增值税{vat_recalc}行, 预扣税{tax_recalc}行, 费用合计{fee_recalc}行")
            
            # 过滤空行：保留序号有效的行，或保留序号为空但有关键费用数据的行
            key_cols = ["台班费_泰铢", "本次结算台班", "加班费", "结算费用_泰铢", "本次费用合计"]
            
            # 先保留序号有效的行
            has_seq = df["序号"].notna() & (df["序号"].astype(str).str.strip() != "")
            
            # 对于序号为空的行，检查是否有关键列有非空数据
            has_data = pd.Series(False, index=df.index)
            for col in key_cols:
                if col in df.columns:
                    has_data |= (df[col].notna() & (df[col].astype(str).str.strip() != ""))
            
            # 保留有序号或有数据的行
            df = df[has_seq | has_data].copy()  # 使用.copy()避免SettingWithCopyWarning
            
            if len(df) == 0:
                print(f"  ⚠️  没有有效数据行")
                report_data.append({
                    "sheet名称": sheet_name,
                    "项目名称": sheet_name,
                    "项目编码": '',
                    "开始月份": start_month,
                    "状态": "失败：无有效数据",
                    "数据行数": 0,
                    "加班费错误行数": 0,
                    "加班费错误率(%)": 0,
                    "加班费差异金额(泰铢)": 0,
                    "结算费用错误行数": 0,
                    "结算费用错误率(%)": 0,
                    "结算费用差异金额(泰铢)": 0,
                    "本次费用合计错误行数": 0,
                    "本次费用合计错误率(%)": 0,
                    "本次费用合计差异金额(泰铢)": 0,
                    "费用合计(泰铢)": 0
                })
                continue
            
            print(f"  有效数据行数: {len(df)}")
            
            # 项目编码/名称：仅使用 raw 合并阶段已写好的列，或文件名；不再 match_project / 弹窗
            default_project_code = file_code or ""
            default_project_name = file_name or ""

            def _cell_str(series_name: str) -> list[str]:
                if series_name not in df.columns:
                    return [""] * len(df)
                return (
                    df[series_name]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace({"nan": "", "None": ""})
                    .tolist()
                )

            codes_in = _cell_str("项目编码")
            names_in = _cell_str("项目名称")
            # 若列映射时「使用项目」被当成了项目名来源，也只作回退，不再匹配库
            use_in = _cell_str("使用项目")

            row_project_codes = []
            row_project_names = []
            for i in range(len(df)):
                code = codes_in[i] if i < len(codes_in) else ""
                name = names_in[i] if i < len(names_in) else ""
                # 优先：明细里已有项目编码/名称（merge-raw 结果）
                # 其次：文件名解析
                # 不再用「使用项目」做匹配库查询（使用项目仅保留原始台账字段）
                if not code:
                    code = default_project_code
                if not name or is_month_like_text(name):
                    name = default_project_name
                if is_month_like_text(code):
                    code = default_project_code or ""
                row_project_codes.append(code)
                row_project_names.append(name)

            df["项目编码"] = row_project_codes
            df["项目名称"] = row_project_names
            df["开始月份"] = start_month
            print(
                f"  项目字段(不弹窗): [{default_project_code or '—'}] {default_project_name or '—'} "
                f"| 行内已有编码 {sum(1 for c in codes_in if c)} 行"
            )
    
            # 区域：仅查映射表，未知则标未知，不弹窗、不写入项目映射
            df["区域"] = df["项目名称"].apply(
                lambda x: guess_region_from_project_or_sheet(x, sheet_name, project_mapping)
            )
    
            # 报告展示项目名
            unique_projects = sorted({p for p in row_project_names if p and p != "nan"})
            report_project_name = (
                unique_projects[0]
                if len(unique_projects) == 1
                else (default_project_name or f"多项目({len(unique_projects)})")
            )
            unique_codes = sorted({c for c in row_project_codes if c})
            report_project_code = (
                unique_codes[0]
                if len(unique_codes) == 1
                else (default_project_code or "")
            )
            unique_regions = sorted({r for r in df["区域"].astype(str).str.strip().tolist() if r and r != "nan"})
            report_region = unique_regions[0] if len(unique_regions) == 1 else f"多区域({len(unique_regions)})"
            
            # 映射标准机具种类并收集新种类
            def map_kind(x):
                # 先做专项归一：拉土车/随车吊 + 车牌号 => 基础名称
                val = normalize_lookup_text(normalize_special_vehicle_kind(x))
                if val and val != "nan" and val not in kind_mapping:
                    new_kinds.add(val)
                return kind_mapping.get(val, val)
            
            df["标准机具种类"] = df["机具种类"].apply(map_kind)
            
            # 映射标准型号规格并收集新型号（基于标准机具种类）
            def map_model(row):
                std_kind = normalize_lookup_text(row["标准机具种类"])
                raw_model = normalize_lookup_text(row["型号规格"])
                
                if not raw_model or raw_model == "nan":
                    return raw_model
                
                # 使用(标准机具种类, 原始型号)作为键
                key = (std_kind, raw_model)
                if key in model_mapping:
                    return model_mapping[key]
                else:
                    # 收集新型号（记录在哪个机具种类下）
                    new_models.add((std_kind, raw_model))
                    return str(row["型号规格"]).strip()
            
            df["标准型号规格"] = df.apply(map_model, axis=1)
            
            # 计算本sheet的费用合计
            sheet_total = df["本次费用合计"].apply(safe_to_float).sum()
            
            all_data.append(df)
            
            # 计算各类费用
            overtime_fee_total = df["加班费"].apply(safe_to_float).sum()
            settlement_total = df["结算费用_泰铢"].apply(safe_to_float).sum()
            
            # 统计有数据的行数
            overtime_count = (df["加班费"].apply(safe_to_float) != 0).sum()
            settlement_count = (df["结算费用_泰铢"].apply(safe_to_float) != 0).sum()
            total_count = (df["本次费用合计"].apply(safe_to_float) != 0).sum()
            
            # 统计验算错误的行数和金额差异
            overtime_error_rows = 0
            settlement_error_rows = 0
            total_error_rows = 0
            overtime_diff_sum = 0.0
            settlement_diff_sum = 0.0
            total_diff_sum = 0.0
            
            for idx in range(len(df)):
                row = df.iloc[idx]
                
                # 检查加班费错误
                is_correct, actual, calculated = validate_overtime_fee(row)
                if not is_correct and actual != 0:
                    overtime_error_rows += 1
                    overtime_diff_sum += abs(actual - calculated)
                
                # 检查结算费用错误
                is_correct, actual, calc1, calc2 = validate_settlement_amount(row)
                if not is_correct and actual != 0:
                    settlement_error_rows += 1
                    # 选择差异较小的计算方式
                    diff = min(abs(actual - calc1), abs(actual - calc2))
                    settlement_diff_sum += diff
                
                # 检查本次费用合计错误
                is_correct, actual, calculated = validate_total_amount(row)
                if not is_correct:
                    total_error_rows += 1
                    total_diff_sum += abs(actual - calculated)
            
            # 计算错误率
            data_rows = len(df)
            overtime_error_rate = (overtime_error_rows / data_rows * 100) if data_rows > 0 else 0
            settlement_error_rate = (settlement_error_rows / data_rows * 100) if data_rows > 0 else 0
            total_error_rate = (total_error_rows / data_rows * 100) if data_rows > 0 else 0
            
            # 添加报告数据
            report_data.append({
                "sheet名称": sheet_name,
                "项目名称": report_project_name,
                "项目编码": report_project_code,
                "开始月份": start_month,
                "状态": "成功",
                "数据行数": len(df),
                "区域": report_region,
                "机具种类数": df["标准机具种类"].nunique(),
                "型号规格数": df["标准型号规格"].nunique(),
                "加班费错误行数": overtime_error_rows,
                "加班费错误率(%)": round(overtime_error_rate, 2),
                "加班费差异金额(泰铢)": round(overtime_diff_sum, 2),
                "结算费用错误行数": settlement_error_rows,
                "结算费用错误率(%)": round(settlement_error_rate, 2),
                "结算费用差异金额(泰铢)": round(settlement_diff_sum, 2),
                "本次费用合计错误行数": total_error_rows,
                "本次费用合计错误率(%)": round(total_error_rate, 2),
                "本次费用合计差异金额(泰铢)": round(total_diff_sum, 2),
                "加班费行数": overtime_count,
                "加班费合计(泰铢)": round(overtime_fee_total, 2),
                "结算费用行数": settlement_count,
                "结算费用合计(泰铢)": round(settlement_total, 2),
                "本次费用行数": total_count,
                "费用合计(泰铢)": round(sheet_total, 2)
            })
    
        wb.close()

    # 3. 仅补充机具种类/型号映射（项目匹配只在 merge-raw 完成，此处不写项目映射、不弹窗）
    if new_kinds or new_models:
        print(f"\n发现新机具数据：")
        if new_kinds:
            print(f"  新机具种类（{len(new_kinds)}个）：")
            smart_mapped = 0
            for k in sorted(new_kinds):
                smart_result = apply_smart_mapping_rules(k)
                if smart_result:
                    kind_mapping[k] = smart_result
                    print(f"    {k} -> {smart_result} (智能映射)")
                    smart_mapped += 1
                else:
                    kind_mapping[k] = k
                    print(f"    {k} (未匹配规则，保持原名)")
            if smart_mapped > 0:
                print(f"  ✅ 智能映射成功 {smart_mapped}/{len(new_kinds)} 个机具种类")
        if new_models:
            print(f"  新型号规格（{len(new_models)}个）：")
            models_by_kind = {}
            for std_kind, raw_model in sorted(new_models):
                models_by_kind.setdefault(std_kind, []).append(raw_model)
            for kind in sorted(models_by_kind.keys()):
                print(f"    [{kind}]")
                for model in models_by_kind[kind]:
                    print(f"      {model}")
            for std_kind, raw_model in new_models:
                model_mapping[(std_kind, raw_model)] = raw_model
        save_machine_mapping(MACHINE_MAPPING_JSON_PATH, kind_mapping, model_mapping)
        print(f"已自动补充机具映射表，请人工复核。\n")
    
    # 4. 合并所有 sheet
    if not all_data:
        print("\n没有读取到任何数据！")
        return
    
    print(f"\n合并 {len(all_data)} 个sheet的数据...")
    # 打印第一个DataFrame的列，用于调试
    if len(all_data) > 0:
        print(f"第一个sheet的列: {list(all_data[0].columns)}")
        
    summary_df = pd.concat(all_data, ignore_index=True)
    print(f"合并后的列: {list(summary_df.columns)}")
    
    # 清理数据格式
    print(f"\n清理数据格式...")
    
    # 1. 清理型号规格：删除换行符、制表符等空白字符
    def clean_model_spec(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""
        # 将所有换行符、制表符等替换为空格，然后合并多余空格
        cleaned = re.sub(r'[\n\r\t]+', ' ', str(value))
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned
    
    summary_df["型号规格"] = summary_df["型号规格"].apply(clean_model_spec)
    summary_df["标准型号规格"] = summary_df["标准型号规格"].apply(clean_model_spec)
    print(f"  ✅ 已清理型号规格列的换行符")
    
    # 2. 清理备注：将连续空格（3个以上）替换为分号
    def clean_remark(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""
        # 将3个以上的连续空格替换为分号
        cleaned = re.sub(r' {3,}', '; ', str(value))
        # 清理换行符为分号
        cleaned = re.sub(r'[\n\r]+', '; ', cleaned)
        # 合并多个分号
        cleaned = re.sub(r';\s*;+', ';', cleaned)
        return cleaned.strip()
    
    summary_df["备注"] = summary_df["备注"].apply(clean_remark)
    print(f"  ✅ 已清理备注列的空格")

    # 2.1 清理使用日期：将连续3个及以上空白替换为分号，避免多个日期黏连
    def clean_use_date(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""

        cleaned = str(value)
        # 先将换行、制表符标准化为空格
        cleaned = re.sub(r'[\n\r\t]+', ' ', cleaned)
        # 连续3个及以上空白替换为中文分号
        cleaned = re.sub(r'\s{3,}', '；', cleaned)
        # 规整分号周围空白并去重
        cleaned = re.sub(r'\s*；\s*', '；', cleaned)
        cleaned = re.sub(r'；{2,}', '；', cleaned)
        return cleaned.strip('； ').strip()

    summary_df["使用日期"] = summary_df["使用日期"].apply(clean_use_date)
    print(f"  ✅ 已清理使用日期列的连续空白")
    
    # 3. 从备注列智能解析月租金额和加班费单价
    def extract_monthly_rent(remark):
        """从备注中提取月租金额（单位：泰铢）"""
        if not remark or str(remark).strip() == "" or str(remark) == "nan":
            return ""
        
        remark_str = str(remark)
        
        # 匹配模式（按优先级）
        patterns = [
            # 模式1: 月租 250000/月
            r'月租[\s]*([\d,.]+)[\s]*/[\s]*月',
            # 模式2: 月租250000/月含油
            r'月租[\s]*([\d,.]+)[\s]*/[\s]*月',
            # 模式3: 月租/70000
            r'月租[\s]*/[\s]*([\d,.]+)',
            # 模式4: 月租125000
            r'月租[\s]*([\d,.]+)(?![\d/天小时])',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, remark_str, re.IGNORECASE)
            if match:
                amount_str = match.group(1)
                # 移除逗号
                amount_str = amount_str.replace(',', '')
                try:
                    amount = float(amount_str)
                    # 验证金额合理性（通常月租在几千到几百万泰铢之间）
                    if 1000 <= amount <= 10000000:
                        return amount
                except ValueError:
                    continue
        
        return ""

    def extract_daily_rent(remark):
        """从备注中提取日租金额（单位：泰铢）"""
        if not remark or str(remark).strip() == "" or str(remark) == "nan":
            return ""

        remark_str = str(remark)

        patterns = [
            r'日租[\s]*([\d,.]+)[\s]*/[\s]*天',
            r'日租[\s]*([\d,.]+)[\s]*/[\s]*日',
            r'每天[\s]*([\d,.]+)[\s]*泰铢',
            r'每天[\s]*([\d,.]+)',
            r'日租[\s]*/[\s]*([\d,.]+)',
            r'日租[\s]*([\d,.]+)(?![\d/天日小时])',
        ]

        for pattern in patterns:
            match = re.search(pattern, remark_str, re.IGNORECASE)
            if match:
                amount_str = match.group(1).replace(',', '')
                try:
                    amount = float(amount_str)
                    if 100 <= amount <= 1000000:
                        return amount
                except ValueError:
                    continue

        return ""
    
    def extract_overtime_rate(remark):
        """从备注中提取加班费单价（单位：泰铢/小时）"""
        if not remark or str(remark).strip() == "" or str(remark) == "nan":
            return ""
        
        remark_str = str(remark)
        
        # 匹配模式（按优先级）
        patterns = [
            # 模式1: 加班费1250/小时
            r'加班费[\s]*([\d,.]+)[\s]*/[\s]*小时',
            # 模式2: 加班 833.33/小时
            r'加班[\s]+([\d,.]+)[\s]*/[\s]*小时',
            # 模式3: 加班费 1041.66/小时已退租
            r'加班费[\s]+([\d,.]+)[\s]*/[\s]*小时',
            # 模式4: 加班 1小时/1200泰铢
            r'加班[\s]+[\d]+[\s]*小时[\s]*/[\s]*([\d,.]+)[\s]*泰铢',
            r'加班[\s]+[\d]+[\s]*小时[\s]*/[\s]*([\d,.]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, remark_str, re.IGNORECASE)
            if match:
                rate_str = match.group(1)
                # 移除逗号
                rate_str = rate_str.replace(',', '')
                try:
                    rate = float(rate_str)
                    # 验证单价合理性（通常加班费单价在几百到几千泰铢/小时）
                    if 100 <= rate <= 50000:
                        return rate
                except ValueError:
                    continue
        
        return ""
    
    summary_df["月租金额_泰铢"] = summary_df["备注"].apply(extract_monthly_rent)
    summary_df["日租金额_泰铢"] = summary_df["备注"].apply(extract_daily_rent)
    summary_df["备注加班费单价"] = summary_df["备注"].apply(extract_overtime_rate)
    
    # 统计提取成功率
    monthly_rent_count = (summary_df["月租金额_泰铢"] != "").sum()
    daily_rent_count = (summary_df["日租金额_泰铢"] != "").sum()
    overtime_rate_count = (summary_df["备注加班费单价"] != "").sum()
    print(f"  ✅ 已从备注中提取月租金额：{monthly_rent_count}/{len(summary_df)} 行")
    print(f"  ✅ 已从备注中提取日租金额：{daily_rent_count}/{len(summary_df)} 行")
    print(f"  ✅ 已从备注中提取加班费单价：{overtime_rate_count}/{len(summary_df)} 行")
    
    # 4. 生成年月列：根据使用日期智能推断
    def extract_year_month(row):
        year_month_field = str(row.get('年月', '')).strip()
        use_date = str(row.get('使用日期', '')).strip()
        month_field = str(row.get('月份', '')).strip()

        # 优先使用源数据中的年月列，避免已有年月被二次解析丢失。
        normalized_from_field = normalize_year_month_text(year_month_field)
        if normalized_from_field:
            return normalized_from_field

        # 其次使用月份列（常见格式：2026年4月、2026-04、202604）
        normalized_from_month = normalize_year_month_text(month_field)
        if normalized_from_month:
            return normalized_from_month
        
        if not use_date or use_date == 'nan' or use_date == '':
            return ""
        
        # 尝试多种日期格式解析
        try:
            # 情况1：Excel日期序列号（纯数字且大于1000）
            if use_date.replace('.', '').replace('-', '').isdigit():
                date_num = float(use_date)
                if date_num > 1000:  # Excel日期序列号
                    from datetime import datetime, timedelta
                    # Excel日期基准：1900-01-01（注意Excel bug：1900不是闰年但Excel认为是）
                    base_date = datetime(1899, 12, 30)
                    actual_date = base_date + timedelta(days=date_num)
                    return actual_date.strftime('%Y-%m')
            
            # 情况2：标准日期格式 YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
            date_patterns = [
                r'(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})',  # YYYY-MM-DD
                r'(\d{1,2})[-/\.](\d{1,2})[-/\.]?(\d{4})',  # DD-MM-YYYY
                r'(\d{4})(\d{2})(\d{2})',  # YYYYMMDD
            ]
            
            for pattern in date_patterns:
                match = re.search(pattern, use_date)
                if match:
                    groups = match.groups()
                    # 判断哪个是年份（4位数字）
                    if len(groups[0]) == 4:
                        year, month = groups[0], groups[1]
                    elif len(groups[2]) == 4:
                        year, month = groups[2], groups[1]
                    else:
                        continue
                    
                    # 验证月份范围
                    month_int = int(month)
                    if 1 <= month_int <= 12:
                        return f"{year}-{month.zfill(2)}"
            
            # 情况3：从月份字段辅助推断年份
            if month_field and month_field != 'nan':
                # 提取月份中的年份信息
                year_match = re.search(r'(20\d{2})', month_field)
                if year_match:
                    year = year_match.group(1)
                    # 提取月份数字
                    month_match = re.search(r'(\d{1,2})月', month_field)
                    if month_match:
                        month = month_match.group(1).zfill(2)
                        return f"{year}-{month}"
        
        except Exception as e:
            pass
        
        return ""  # 无法解析则返回空
    
    summary_df["年月"] = summary_df.apply(extract_year_month, axis=1)
    year_month_count = (summary_df["年月"] != "").sum()
    print(f"  ✅ 已生成年月列，成功解析 {year_month_count}/{len(summary_df)} 行")
    
    # 检测并修正台班费和结算台班填反的情况
    print(f"\n检测数据填反情况...")
    summary_df, swapped_indices = detect_and_fix_swapped_values(summary_df)
    if swapped_indices:
        print(f"✅ 已自动修正 {len(swapped_indices)} 行台班费/结算台班填反的数据，将用浅蓝色标注")
    else:
        print(f"✅ 未发现台班费/结算台班填反的数据")
    
    # 检测并修正出租单位和月份填反的情况
    print(f"\n检测出租单位和月份填反情况...")
    summary_df, swapped_month_indices = detect_and_fix_supplier_month_swap(summary_df)
    if swapped_month_indices:
        print(f"✅ 已自动修正 {len(swapped_month_indices)} 行出租单位/月份填反的数据，将用黄色标注")
    else:
        print(f"✅ 未发现出租单位/月份填反的数据")

    # 月份列可能在上一步被纠正，需重算年月，确保汇总列使用最新值
    summary_df["年月"] = summary_df.apply(extract_year_month, axis=1)
    refreshed_year_month_count = (summary_df["年月"] != "").sum()
    print(f"✅ 已按最新月份信息重算年月列，成功解析 {refreshed_year_month_count}/{len(summary_df)} 行")
    
    # 调整列顺序（基础字段统一为：项目编码、项目名称）
    col_order = [
        "项目编码", "项目名称", "标准机具种类", "标准型号规格",
        "序号", "台班费_泰铢", "本次结算台班",
        "加班小时", "加班费", "进出场费运费", "移动费拆装费", "结算费用_泰铢",
        "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%", "本次费用合计",
        "使用日期", "出租单位", "月份", "年月", "备注", "月租金额_泰铢", "日租金额_泰铢", "备注加班费单价"
    ]
    
    # 检查是否所有列都存在
    missing_cols = [col for col in col_order if col not in summary_df.columns]
    if missing_cols:
        print(f"\n⚠️  警告：以下列在数据中不存在: {missing_cols}")
        print(f"实际存在的列: {list(summary_df.columns)}")
    
    # 只选择存在的列进行重新排序
    existing_cols = [col for col in col_order if col in summary_df.columns]
    summary_df = summary_df[existing_cols]
    summary_df = summary_df.rename(columns={"标准机具种类": "机具种类", "标准型号规格": "型号规格"})
    print(f"重新排序后的列: {list(summary_df.columns)}")

    # 5. 计算汇总金额并验证
    summary_total = summary_df["本次费用合计"].apply(safe_to_float).sum()
    report_total = sum(item.get("费用合计(泰铢)", 0) for item in report_data if item["状态"] == "成功")
    
    print(f"\n{'='*60}")
    print(f"金额验证：")
    print(f"  各sheet费用合计: {report_total:,.2f} 泰铢")
    print(f"  汇总表费用合计: {summary_total:,.2f} 泰铢")
    if abs(summary_total - report_total) < 0.01:
        print(f"  ✅ 金额验证通过，差异: {abs(summary_total - report_total):.2f} 泰铢")
    else:
        print(f"  ⚠️  金额存在差异: {abs(summary_total - report_total):.2f} 泰铢")
    print(f"{'='*60}\n")
    
    # 添加汇总行到报告
    total_data_rows = sum(item.get("数据行数", 0) for item in report_data)
    total_overtime_errors = sum(item.get("加班费错误行数", 0) for item in report_data)
    total_settlement_errors = sum(item.get("结算费用错误行数", 0) for item in report_data)
    total_amount_errors = sum(item.get("本次费用合计错误行数", 0) for item in report_data)
    
    report_data.append({
        "sheet名称": "【合计】",
        "项目名称": "",
        "项目编码": "",
        "开始月份": "",
        "状态": "",
        "数据行数": total_data_rows,
        "区域": "",
        "机具种类数": "",
        "型号规格数": "",
        "加班费错误行数": total_overtime_errors,
        "加班费错误率(%)": round((total_overtime_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "加班费差异金额(泰铢)": round(sum(item.get("加班费差异金额(泰铢)", 0) for item in report_data), 2),
        "结算费用错误行数": total_settlement_errors,
        "结算费用错误率(%)": round((total_settlement_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "结算费用差异金额(泰铢)": round(sum(item.get("结算费用差异金额(泰铢)", 0) for item in report_data), 2),
        "本次费用合计错误行数": total_amount_errors,
        "本次费用合计错误率(%)": round((total_amount_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "本次费用合计差异金额(泰铢)": round(sum(item.get("本次费用合计差异金额(泰铢)", 0) for item in report_data), 2),
        "加班费行数": sum(item.get("加班费行数", 0) for item in report_data),
        "加班费合计(泰铢)": round(sum(item.get("加班费合计(泰铢)", 0) for item in report_data), 2),
        "结算费用行数": sum(item.get("结算费用行数", 0) for item in report_data),
        "结算费用合计(泰铢)": round(sum(item.get("结算费用合计(泰铢)", 0) for item in report_data), 2),
        "本次费用行数": sum(item.get("本次费用行数", 0) for item in report_data),
        "费用合计(泰铢)": round(report_total, 2)
    })
    
    # 6. 生成处理报告（导出时不保留sheet名称列）
    report_df = pd.DataFrame(report_data).drop(columns=["sheet名称"], errors="ignore")
    
    # 8. 处理数值列（转换为浮点数，保留原始数值类型）
    numeric_columns = [
        "台班费_泰铢", "本次结算台班", "加班小时", "加班费",
        "进出场费运费", "移动费拆装费", "结算费用_泰铢",
        "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%", "本次费用合计",
        "月租金额_泰铢", "日租金额_泰铢", "备注加班费单价"
    ]
    
    # 注意：年月列不进行数值转换，保持文本格式
    
    # 定义需要货币格式的列
    currency_columns = [
        "台班费_泰铢", "加班费", "进出场费运费", "移动费拆装费",
        "结算费用_泰铢", "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%",
        "月租金额_泰铢", "日租金额_泰铢", "备注加班费单价"
    ]
    
    for col in numeric_columns:
        if col in summary_df.columns:
            summary_df[col] = summary_df[col].apply(safe_to_float)

    # 生成项目汇总机具费表
    project_summary_df = summary_df.groupby(['项目编码', '项目名称']).agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count',
        '机具种类': 'nunique',
        '出租单位': 'nunique',
        '月份': 'nunique'
    }).round(2).reset_index()
    project_summary_df.columns = ['项目编码', '项目名称', '费用合计(泰铢)', '使用次数', '机具种类数', '出租单位数', '覆盖月份数']
    project_summary_df = project_summary_df.sort_values('费用合计(泰铢)', ascending=False).reset_index(drop=True)

    # 新增：按“年月(YYYY-MM)”生成项目月份金额宽表
    monthly_wide_source = summary_df.copy()
    monthly_wide_source["月度项目编码"] = monthly_wide_source["项目编码"].fillna("").astype(str).str.strip()
    monthly_wide_source["月度项目名称"] = monthly_wide_source["项目名称"].fillna("未匹配项目").astype(str).str.strip()
    monthly_wide_source["月度年月"] = monthly_wide_source.apply(extract_year_month, axis=1)
    monthly_wide_source["月度年月"] = monthly_wide_source["月度年月"].fillna("").astype(str).str.strip()
    monthly_wide_source = monthly_wide_source[
        monthly_wide_source["月度年月"].str.match(r"^20\d{2}-\d{2}$", na=False)
    ].copy()

    project_monthly_wide_report = monthly_wide_source.pivot_table(
        index=["月度项目编码", "月度项目名称"],
        columns="月度年月",
        values="本次费用合计",
        aggfunc="sum",
        fill_value=0
    )

    wide_month_columns = list(project_monthly_wide_report.columns)
    project_monthly_wide_report = project_monthly_wide_report.reindex(
        columns=sorted(wide_month_columns)
    )
    project_monthly_wide_report = project_monthly_wide_report.reset_index().round(2)
    project_monthly_wide_report = project_monthly_wide_report.rename(columns={
        "月度项目编码": "项目编码",
        "月度项目名称": "项目名称"
    })
    
    # 8. 保存到新Excel（输出到汇总文件夹，带时间戳）
    # 确保输出文件夹存在
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    if output_basename_override:
        output_basename = output_basename_override
    elif len(excel_paths) > 1:
        grouped_project_names = {
            extract_project_name_from_filename(path)
            for path in excel_paths
            if extract_project_name_from_filename(path)
        }
        if len(grouped_project_names) == 1:
            output_basename = build_project_output_basename(grouped_project_names.pop(), match_db)
        else:
            output_basename = f"{OUTPUT_FILENAME_PREFIX}_全量标准化"
    else:
        output_basename = build_summary_output_basename(primary_excel_path)
    output_filename = f"{output_basename}_{OUTPUT_FILE_SUFFIX}.xlsx"
    output_path = Path(OUTPUT_FOLDER) / output_filename

    # D 目录策略：写完目标文件后，只保留全量标准化汇总（见 _keep_only_full_summary_in_d）
    old_summary_files = []
    for old_file in Path(OUTPUT_FOLDER).glob("*.xlsx"):
        if old_file.name.startswith("~$"):
            continue
        if old_file != output_path:
            old_summary_files.append(old_file)

    if old_summary_files:
        print(f"检测到 {len(old_summary_files)} 个旧/多余汇总文件，输出成功后将删除")
    
    # 定义填充颜色
    pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")  # 粉红色：本次费用合计错误
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # 浅蓝色：台班费/结算台班已交换
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # 黄色：出租单位/月份已交换
    orange_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # 金色：加班费错误
    light_green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # 浅绿色：结算费用错误
    
    # 验证数据并记录错误
    overtime_error_count = 0
    settlement_error_count = 0
    total_error_count = 0
    
    print(f"\n验证费用计算...")
    
    # 验证加班费
    print(f"\n1. 验证加班费...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calculated = validate_overtime_fee(row)
        if not is_correct and actual != 0:  # 只报告非零的错误
            overtime_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行加班费错误：实际={actual:.2f}，应为={calculated:.2f}，差异={abs(actual-calculated):.2f}")
    
    if overtime_error_count > 0:
        print(f"发现 {overtime_error_count} 行加班费计算错误，将用金色标注")
    else:
        print(f"✅ 所有加班费验证通过")
    
    # 验证结算费用
    print(f"\n2. 验证结算费用...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calc1, calc2 = validate_settlement_amount(row)
        if not is_correct and actual != 0:  # 只报告非零的错误
            settlement_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行结算费用错误：实际={actual:.2f}，方式1={calc1:.2f}，方式2={calc2:.2f}")
    
    if settlement_error_count > 0:
        print(f"发现 {settlement_error_count} 行结算费用计算错误，将用浅绿色标注")
    else:
        print(f"✅ 所有结算费用验证通过")
    
    # 验证本次费用合计
    print(f"\n3. 验证本次费用合计...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calculated = validate_total_amount(row)
        if not is_correct:
            total_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行本次费用合计错误：实际={actual:.2f}，应为={calculated:.2f}，差异={abs(actual-calculated):.2f}")
    
    if total_error_count > 0:
        print(f"发现 {total_error_count} 行本次费用合计计算错误，将用粉红色标注")
    else:
        print(f"✅ 所有本次费用合计验证通过")
    
    # 总结
    total_errors = overtime_error_count + settlement_error_count + total_error_count
    if total_errors > 0:
        print(f"\n⚠️  共发现 {total_errors} 处计算错误")
    else:
        print(f"\n✅ 所有数据验证通过，无计算错误")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 顺序：汇总 置首，再项目汇总 / 月份宽表 / 处理报告
        summary_df.to_excel(writer, sheet_name=SUMMARY_SHEET_NAME, index=False)
        project_summary_df.to_excel(writer, sheet_name=PROJECT_SUMMARY_SHEET_NAME, index=False)
        project_monthly_wide_report.to_excel(writer, sheet_name=PROJECT_MONTHLY_WIDE_SHEET_NAME, index=False)
        report_df.to_excel(writer, sheet_name=PROCESS_REPORT_SHEET_NAME, index=False)
        

        from openpyxl.styles import Alignment, Font

        # 定义加粗字体
        bold_font = Font(bold=True)

        def format_summary_like_sheet(ws, display_df, left_align_use_date=False):
            """按汇总页规则格式化工作表；可选将使用日期列左对齐"""
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

            for col_idx, col_name in enumerate(display_df.columns, 1):
                max_length = len(str(col_name))
                for row_idx in range(2, len(display_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell_value = cell.value
                    df_row_idx = row_idx - 2

                    # 为货币列设置货币格式（2位小数，千分位）
                    if col_name in currency_columns:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'
                    # 为其他数值列设置普通数值格式
                    elif col_name in numeric_columns:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'
                    # 为使用日期列设置日期格式
                    elif col_name == "使用日期":
                        if isinstance(cell_value, (int, float)) and cell_value > 0:
                            cell.number_format = 'YYYY-MM-DD'
                        elif isinstance(cell_value, str) and cell_value.strip():
                            try:
                                if cell_value.strip().replace('.', '').replace('-', '').replace('/', '').isdigit():
                                    date_value = float(cell_value)
                                    if date_value > 0:
                                        cell.value = date_value
                                        cell.number_format = 'YYYY-MM-DD'
                            except:
                                pass

                        if left_align_use_date:
                            cell.alignment = Alignment(horizontal='left')
                    # 年月列设置居中
                    elif col_name == "年月":
                        cell.alignment = Alignment(horizontal='center')

                    # 为交换过的行标注浅蓝色（台班费和结算台班列）
                    if df_row_idx in swapped_indices and col_name in ['台班费_泰铢', '本次结算台班']:
                        cell.fill = light_blue_fill

                    # 为交换过的行标注黄色（出租单位和月份列）
                    if df_row_idx in swapped_month_indices and col_name in ['出租单位', '月份']:
                        cell.fill = yellow_fill

                    # 验证错误着色（基于summary_df同索引行）
                    if df_row_idx >= 0 and df_row_idx < len(summary_df):
                        source_row = summary_df.iloc[df_row_idx]

                        if col_name == "加班费":
                            is_correct, actual, calculated = validate_overtime_fee(source_row)
                            if not is_correct and actual != 0:
                                cell.fill = orange_fill

                        if col_name == "结算费用_泰铢":
                            is_correct, actual, calc1, calc2 = validate_settlement_amount(source_row)
                            if not is_correct and actual != 0:
                                cell.fill = light_green_fill

                        if col_name == "本次费用合计":
                            cell.font = bold_font
                            is_correct, actual, calculated = validate_total_amount(source_row)
                            if not is_correct:
                                cell.fill = pink_fill

                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))

                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions

        def format_monthly_report_sheet(ws, display_df):
            """格式化项目月度机具费汇总表"""
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

            for col_idx, col_name in enumerate(display_df.columns, 1):
                max_length = len(str(col_name))
                is_amount_col = col_name not in ["项目编码", "项目名称"]

                for row_idx in range(2, len(display_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell_value = cell.value

                    if is_amount_col:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'

                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))

                if col_name == "项目编码":
                    adjusted_width = min(max(max_length + 2, 16), 24)
                elif col_name == "项目名称":
                    adjusted_width = min(max(max_length + 2, 18), 40)
                else:
                    adjusted_width = max(14, min(max_length + 2, 18))
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            ws.freeze_panes = 'C2'
            ws.auto_filter.ref = ws.dimensions
        
        # 美化汇总表
        ws_summary = writer.sheets[SUMMARY_SHEET_NAME]
        format_summary_like_sheet(ws_summary, summary_df, left_align_use_date=False)

        # 美化项目汇总机具费表
        ws_project_summary = writer.sheets[PROJECT_SUMMARY_SHEET_NAME]
        ws_project_summary.row_dimensions[1].height = 40
        for cell in ws_project_summary[1]:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        for col_idx, col_name in enumerate(project_summary_df.columns, 1):
            max_length = len(str(col_name))
            for row_idx in range(2, len(project_summary_df) + 2):
                cell = ws_project_summary.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                if col_name == '费用合计(泰铢)':
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    cell.number_format = '#,##0.00'
                elif col_name in ['使用次数', '机具种类数', '出租单位数', '覆盖月份数']:
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    cell.number_format = '0'
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            adjusted_width = min(max_length + 2, 40)
            ws_project_summary.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        ws_project_summary.freeze_panes = 'A2'
        ws_project_summary.auto_filter.ref = ws_project_summary.dimensions

        # 美化项目月份金额宽表
        ws_project_monthly_wide = writer.sheets[PROJECT_MONTHLY_WIDE_SHEET_NAME]
        format_monthly_report_sheet(ws_project_monthly_wide, project_monthly_wide_report)
        
        # 美化报告表
        ws_report = writer.sheets[PROCESS_REPORT_SHEET_NAME]
        
        # 设置首行格式：高度40，自动换行
        ws_report.row_dimensions[1].height = 40
        for cell in ws_report[1]:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        
        for col_idx, col_name in enumerate(report_df.columns, 1):
            max_length = len(str(col_name))
            for row_idx in range(2, len(report_df) + 2):
                cell = ws_report.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                
                # 为费用合计列设置货币格式（带泰铢符号）
                if col_name == "费用合计(泰铢)":
                    # 转换字符串为数字
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    # 应用货币格式
                    cell.number_format = '#,##0.00'
                
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            adjusted_width = min(max_length + 2, 40)
            ws_report.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        
        # 冻结首行
        ws_report.freeze_panes = 'A2'
        
        # 添加自动筛选
        ws_report.auto_filter.ref = ws_report.dimensions
        

    # 新文件写入成功后，删除旧汇总文件
    deleted_old_files = 0
    for old_file in old_summary_files:
        try:
            old_file.unlink()
            deleted_old_files += 1
            print(f"已删除旧汇总文件: {old_file.name}")
        except Exception as e:
            print(f"⚠️ 删除旧汇总文件失败: {old_file.name}，原因: {e}")

    print(f"\n{'='*60}")
    print(f"汇总完成！文件已保存为：")
    print(f"  {output_path}")
    if deleted_old_files > 0:
        print(f"已删除 {deleted_old_files} 个旧汇总文件")
    print(f"{'='*60}")
    print(f"共处理 {len(report_data)-1} 个 sheet")  # -1因为包含了合计行
    print(f"成功汇总 {len(all_data)} 个 sheet，共 {len(summary_df)} 行数据")
    print(f"\n处理报告：")
    for item in report_data[:-1]:  # 不显示合计行
        status_icon = "✅" if item["状态"] == "成功" else "❌"
        amount_str = f"{item.get('费用合计(泰铢)', 0):,.2f}" if item["状态"] == "成功" else "0.00"
        print(f"  {status_icon} {item['sheet名称']}: {item['状态']} - {item.get('数据行数', 0)} 行 - {amount_str} 泰铢")
    print(f"\n  💰 总计: {report_total:,.2f} 泰铢")
    print(f"\n已生成4个工作表（汇总置首）：")
    print(f"  1. '{SUMMARY_SHEET_NAME}' - 所有数据明细 (共 {len(summary_df)} 行)")
    print(f"  2. '{PROJECT_SUMMARY_SHEET_NAME}' - 项目汇总机具使用费")
    print(f"  3. '{PROJECT_MONTHLY_WIDE_SHEET_NAME}' - 各项目按年月横向金额宽表")
    print(f"  4. '{PROCESS_REPORT_SHEET_NAME}' - 各项目处理情况统计")
    
    # 9. 生成筛选分析
    print(f"\n{'='*60}")
    print(f"数据分析：")
    print(f"{'='*60}")
    
    # 按标准机具种类统计
    print(f"\n🔧 按机具种类统计（Top 10）：")
    kind_stats = summary_df.groupby('机具种类').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count'
    }).round(2)
    kind_stats.columns = ['费用合计(泰铢)', '使用次数']
    kind_stats = kind_stats.sort_values('费用合计(泰铢)', ascending=False).head(10)
    for idx, (kind, row) in enumerate(kind_stats.iterrows(), 1):
        print(f"  {idx}. {kind}: {row['费用合计(泰铢)']:,.2f} 泰铢 ({int(row['使用次数'])} 次)")
    
    # 按项目统计
    print(f"\n📊 按项目统计：")
    project_stats = summary_df.groupby('项目名称').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '机具种类': 'nunique'
    }).round(2)
    project_stats.columns = ['费用合计(泰铢)', '机具种类数']
    project_stats = project_stats.sort_values('费用合计(泰铢)', ascending=False)
    for project, row in project_stats.iterrows():
        if project:  # 排除空项目名
            print(f"  {project}: {row['费用合计(泰铢)']:,.2f} 泰铢 ({int(row['机具种类数'])} 种机具)")
    
    print(f"\n{'='*60}")



def _keep_only_full_summary_in_d() -> None:
    """D_机具费_std_sum 仅保留 机具费汇总_全量标准化_汇总.xlsx。"""
    from machine.config import FULL_SUMMARY_FILENAME, STD_SUM_DIR

    keep_name = FULL_SUMMARY_FILENAME
    out_dir = Path(STD_SUM_DIR)
    if not out_dir.exists():
        return
    removed = 0
    for f in out_dir.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        if f.name != keep_name:
            try:
                f.unlink()
                removed += 1
                print(f"  已清理 D 多余文件: {f.name}")
            except Exception as e:
                print(f"  ⚠️ 无法删除 {f.name}: {e}")
    if removed:
        print(f"  D_机具费_std_sum 现仅保留: {keep_name}（已删 {removed} 个）")
    else:
        print(f"  D_机具费_std_sum 仅保留: {keep_name}")


def run_summarize_batch() -> None:
    """按 RUN_MODE 批量汇总：D 仅输出全量标准化汇总。"""
    if EXCEL_FILE_OVERRIDE and str(EXCEL_FILE_OVERRIDE).strip():
        main()
        _keep_only_full_summary_in_d()
        return

    all_input_files = list_all_standardized_files(STANDARDIZED_INPUT_DIR)
    if not all_input_files:
        print(f"找不到标准化文件：{STANDARDIZED_INPUT_DIR}")
        return

    if RUN_MODE == "all":
        # D_机具费_std_sum 只保留「机具费汇总_全量标准化_汇总.xlsx」
        # 按项目的结果在 E_机具费_std_split，不再在 D 写单项目汇总
        print(
            f"自动发现原始拆分文件 {len(all_input_files)} 个，"
            f"一次汇总写入 D_机具费_std_sum（仅全量标准化汇总）..."
        )
        main(all_input_files, output_basename_override=f"{OUTPUT_FILENAME_PREFIX}_全量标准化")
        _keep_only_full_summary_in_d()
        try:
            from machine.split import split_std_summary_by_project
            print(f"\n{'='*60}")
            print("标准拆分：D_std_sum → E_std_split/")
            print("=" * 60)
            split_std_summary_by_project()
        except Exception as e:
            print(f"⚠️ 标准拆分跳过/失败: {e}")
    else:
        print(f"自动发现标准化文件 {len(all_input_files)} 个，开始按月份组合汇总...")
        month_to_files = {}
        undated_files = []
        for input_file in all_input_files:
            ym = extract_year_month_from_filename(input_file)
            if ym is None:
                undated_files.append(input_file)
            else:
                month_to_files.setdefault(ym, []).append(input_file)

        for ym, files in month_to_files.items():
            files.sort(key=lambda p: (p.stat().st_mtime, p.name))

        jobs = []
        sorted_months = sorted(month_to_files.keys())

        def prev_month(ym):
            y, m = ym
            return (y - 1, 12) if m == 1 else (y, m - 1)

        for idx, ym in enumerate(sorted_months):
            prev_ym = prev_month(ym)
            has_next = idx < len(sorted_months) - 1
            current_file = month_to_files[ym][-1]
            if prev_ym in month_to_files:
                prev_file = month_to_files[prev_ym][-1]
                jobs.append(
                    (
                        [prev_file, current_file],
                        f"{prev_ym[0]}-{prev_ym[1]:02d} + {ym[0]}-{ym[1]:02d} => {ym[0]}-{ym[1]:02d}",
                    )
                )
            elif not has_next:
                jobs.append((current_file, f"{ym[0]}-{ym[1]:02d}（单月）"))
            else:
                print(f"跳过单独输出：{current_file.name}（将并入下月汇总）")

        for f in undated_files:
            jobs.append((f, f"未识别月份（单文件）：{f.name}"))

        if not jobs:
            print("未构建到可执行任务，请检查文件命名中的年月格式（如 2026年4月）。")
        else:
            print(f"共构建任务 {len(jobs)} 个。")
            for idx, (job_arg, job_desc) in enumerate(jobs, start=1):
                print(f"\n[{idx}/{len(jobs)}] {job_desc}")
                main(job_arg)


if __name__ == "__main__":
    run_summarize_batch()
