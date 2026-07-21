# -*- coding: utf-8 -*-
from __future__ import annotations
import re
from machine.textutil import normalize_column_name

def find_matching_column(df_columns, target_patterns, exact_only=False):
    """
    在DataFrame列中查找匹配目标模式的列名
    target_patterns: 列表，包含可能的列名关键词
    返回：匹配到的列名，如果没找到返回None
    """
    normalized_cols = {normalize_column_name(col): col for col in df_columns}
    
    for pattern in target_patterns:
        pattern_norm = normalize_column_name(pattern)
        # 精确匹配
        if pattern_norm in normalized_cols:
            return normalized_cols[pattern_norm]
        if exact_only:
            continue
        # 包含匹配
        for norm_col, orig_col in normalized_cols.items():
            if pattern_norm in norm_col or norm_col in pattern_norm:
                return orig_col
    return None


def find_use_project_column(df_columns):
    """更稳健识别“使用项目”列，优先命中项目字段，排除项目说明/项目月份等干扰列。"""
    normalized_pairs = [(normalize_column_name(col), col) for col in df_columns]

    # 1) 先做精确优先
    exact_candidates = ["使用项目", "使用项目名称", "项目名称", "项目"]
    for pattern in exact_candidates:
        pattern_norm = normalize_column_name(pattern)
        for norm_col, orig_col in normalized_pairs:
            if norm_col == pattern_norm:
                return orig_col

    # 2) 再做包含匹配，但排除明显非项目名字段
    include_tokens = ["使用项目", "项目名称", "项目"]
    exclude_tokens = ["月份", "说明", "备注", "编码", "状态", "经理", "区域"]
    for norm_col, orig_col in normalized_pairs:
        if not any(token in norm_col for token in include_tokens):
            continue
        if any(token in norm_col for token in exclude_tokens):
            continue
        return orig_col

    return None


def find_header_row(ws):
    """在工作表中查找包含目标列名的行"""
    for row_idx in range(1, min(20, ws.max_row + 1)):  # 搜索前20行
        row_values = []
        for cell in ws[row_idx]:
            val = str(cell.value) if cell.value else ""
            # 清理换行符和多余空格
            val = re.sub(r'\s+', '', val)
            row_values.append(val)
        
        # 检查是否包含关键列名
        if any("序号" in val for val in row_values) and any("机具种类" in val for val in row_values):
            return row_idx
    return None


def extract_project_info(ws):
    """从第二行提取项目名称和开始月份"""
    project_name = ""
    start_month = ""
    
    # 搜索前10行中包含"项目名称"和"开始月份"的单元格
    for row_idx in range(1, min(11, ws.max_row + 1)):
        for cell in ws[row_idx]:
            cell_value = str(cell.value).strip() if cell.value else ""
            
            # 查找"项目名称：XXX"
            if "项目名称" in cell_value and "：" in cell_value:
                project_name = cell_value.split("：", 1)[1].strip()
            elif "项目名称" in cell_value:
                # 如果冒号在相邻单元格
                next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                if next_cell.value:
                    project_name = str(next_cell.value).strip()
            
            # 查找"开始月份：XXXX"
            if "开始月份" in cell_value and "：" in cell_value:
                start_month = cell_value.split("：", 1)[1].strip()
            elif "开始月份" in cell_value:
                # 如果冒号在相邻单元格
                next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                if next_cell.value:
                    start_month = str(next_cell.value).strip()
    
    return project_name, start_month


def extract_project_name_from_sheet(sheet_name):
    """从sheet名称提取项目名称"""
    # 尝试匹配 "XX项目" 模式
    match = re.match(r'^(.+?项目)', sheet_name)
    if match:
        return match.group(1)
    
    # 如果有下划线或空格，取第一部分
    if '_' in sheet_name:
        return sheet_name.split('_')[0]
    if ' ' in sheet_name:
        return sheet_name.split(' ')[0]
    
    return sheet_name  # 否则使用整个sheet名
