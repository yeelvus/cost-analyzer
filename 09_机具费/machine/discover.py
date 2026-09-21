# -*- coding: utf-8 -*-
from __future__ import annotations
import re
import os
import sys
from pathlib import Path
from datetime import datetime

code_root = Path(__file__).resolve().parents[2]
if str(code_root) not in sys.path:
    sys.path.insert(0, str(code_root))

from common.project import match_project

LATEST_FILE_MONTH_RE = re.compile(r"(20\d{2})\s*年\s*(1[0-2]|0?[1-9])\s*月")
PROJECT_SPLIT_FILE_RE = re.compile(r"^(20\d{2}(?:0[1-9]|1[0-2]))_(.+?)机具费(?:_标准化)?\.xlsx$")


PROCESS_FILE_MARKERS = ("合并", "clean_机具费", "按项目拆分")
PROJECT_CODE_FILE_RE = re.compile(
    r"^(?P<code>THCC-\d{4}-\d{3}|XCY-\d+|未编码)_(?P<name>.+?)_机具费(?:_标准化)?\.xlsx$",
    re.IGNORECASE,
)


def is_process_or_junk_file(file_path: Path) -> bool:
    """过程文件 / 临时文件，不参与 summarize。"""
    name = file_path.name
    if name.startswith("~$"):
        return True
    if name.startswith("clean_"):
        return True
    for mk in PROCESS_FILE_MARKERS:
        if mk in name:
            return True
    return False


def list_all_standardized_files(input_dir):
    """列出标准化目录下「项目明细」xlsx，排除过程合并文件。"""
    input_path = Path(input_dir)
    if not input_path.exists():
        return []

    all_xlsx = [
        p for p in input_path.glob("*.xlsx")
        if p.is_file() and not is_process_or_junk_file(p)
    ]

    preferred = [p for p in all_xlsx if p.name.endswith("_标准化.xlsx")]
    preferred_set = {p.resolve() for p in preferred}
    others = [p for p in all_xlsx if p.resolve() not in preferred_set]

    return sorted(preferred, key=lambda p: p.name) + sorted(others, key=lambda p: p.name)


def extract_project_code_and_name_from_filename(file_path: Path) -> tuple[str, str]:
    """从项目明细文件名提取 (项目编码, 项目名称)。

    支持：
      THCC-2025-002_泰华电子项目_机具费.xlsx
      202604_伟隆项目机具费.xlsx  → 编码为空，名称=伟隆项目
    """
    match2 = PROJECT_CODE_FILE_RE.match(file_path.name)
    if match2:
        return match2.group("code").strip(), match2.group("name").strip()

    match = PROJECT_SPLIT_FILE_RE.match(file_path.name)
    if match:
        return "", match.group(2).strip()

    stem = file_path.stem.strip()
    stem = re.sub(r"_机具费$", "", stem)
    stem = re.sub(r"_标准化$", "", stem)
    return "", stem


def extract_project_name_from_filename(file_path: Path) -> str | None:
    """从项目明细文件名提取项目名。"""
    _, name = extract_project_code_and_name_from_filename(file_path)
    return name if name else None


def extract_year_month_from_filename(file_path):
    """从文件名提取年月，用于“最新月份”判定。"""
    match = LATEST_FILE_MONTH_RE.search(file_path.name)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def pick_latest_standardized_file(input_dir):
    """自动选择标准化目录中最新文件：优先按文件名年月，其次按修改时间。"""
    input_path = Path(input_dir)
    if not input_path.exists():
        return None

    candidates = [
        p for p in input_path.glob("*.xlsx")
        if p.is_file() and not is_process_or_junk_file(p)
    ]
    if not candidates:
        return None

    dated_candidates = []
    undated_candidates = []
    for file_path in candidates:
        ym = extract_year_month_from_filename(file_path)
        mtime = file_path.stat().st_mtime
        if ym is not None:
            dated_candidates.append((ym, mtime, file_path))
        else:
            undated_candidates.append((mtime, file_path))

    if dated_candidates:
        dated_candidates.sort(key=lambda item: (item[0], item[1]))
        return dated_candidates[-1][2]

    undated_candidates.sort(key=lambda item: item[0])
    return undated_candidates[-1][1]



def build_summary_output_basename(excel_path: Path) -> str:
    """输出名使用原台账名，去掉“_标准化”后再拼接时间后缀。"""
    stem = excel_path.stem
    if stem.endswith("_标准化"):
        stem = stem[:-4]
    return stem



def group_files_by_project(file_paths):
    """将标准化目录中的月度项目文件按项目名分组。"""
    project_to_files = {}
    unmatched_files = []

    for file_path in file_paths:
        project_name = extract_project_name_from_filename(file_path)
        if not project_name:
            unmatched_files.append(file_path)
            continue
        project_to_files.setdefault(project_name, []).append(file_path)

    for project_name, grouped_files in project_to_files.items():
        grouped_files.sort(key=lambda p: (extract_year_month_from_filename(p) or (0, 0), p.name))

    return project_to_files, unmatched_files


def build_project_output_basename(project_name, match_db):
    """生成按项目汇总输出名：项目编码_项目名称_机具费。"""
    project_name = str(project_name).strip()
    if not project_name:
        return f"需手动核对_未匹配项目_机具费"

    hit = match_project(project_name, match_db)
    if hit:
        project_code = str(hit.get('project_code', '')).strip() or '需手动核对'
        standard_name = str(hit.get('project_name', '')).strip() or project_name
    else:
        project_code = '需手动核对'
        standard_name = project_name

    safe_name = re.sub(r'[\\/:*?"<>|]', '_', standard_name)
    safe_code = re.sub(r'[\\/:*?"<>|]', '_', project_code)
    return f"{safe_code}_{safe_name}_机具费"
