# -*- coding: utf-8 -*-
"""按项目拆开一级二级科目宽表。

每个项目一个工作簿，放在
03_Sum/12_财务付款/一级二级科目宽/{组名}/ 下。

工作簿里有两张表：
- 汇总：这个项目的一级小计、二级科目和全部汇总
- 一级二级科目宽表：同样的科目结构，二级科目下面挂付款明细
"""

from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import Paths
from .payment_amount import ensure_payment_amount
from .summary_build import (
    SUBJECT_WIDE_SHEET,
    build_subject_wide_from_detail_df,
    write_subject_wide_sheet,
)
from .utils import level1_sort_key, safe_filename, safe_text

SPLIT_DIR_NAME = "一级二级科目宽"
SUMMARY_SHEET = "汇总"
AMOUNT_HEADER = "付款金额"
LINE_COLUMNS = [
    "三级科目",
    "日期",
    "摘要",
    "供应商名称",
    "应付金额",
    "预扣税额",
    "实付金额",
    "增值税额",
    "不含税额",
    "币别",
    "汇率",
    "年月",
]
AMOUNT_LINE_COLUMNS = {
    "应付金额",
    "预扣税额",
    "实付金额",
    "增值税额",
    "不含税额",
}
_UNCLASSIFIED_L1 = "未分类一级科目"
_UNCLASSIFIED_L2 = "未分类二级科目"


def subject_wide_split_root(paths: Paths) -> Path:
    return Path(paths.sum_dir) / SPLIT_DIR_NAME


def _prepare_detail(detail: pd.DataFrame) -> pd.DataFrame:
    work = ensure_payment_amount(detail).copy()
    if "付款金额" not in work.columns:
        return work.iloc[0:0].copy()
    work["付款金额"] = pd.to_numeric(work["付款金额"], errors="coerce").fillna(0.0)
    work = work[work["付款金额"] != 0]
    if work.empty:
        return work
    for col, fallback in (("一级科目", _UNCLASSIFIED_L1), ("二级科目", _UNCLASSIFIED_L2)):
        if col not in work.columns:
            work[col] = fallback
        work[col] = work[col].map(lambda v, fb=fallback: safe_text(v) or fb)
    for col in ("项目编码", "项目名称"):
        if col not in work.columns:
            work[col] = ""
        work[col] = work[col].map(safe_text)
    if "日期" in work.columns:
        work["_日期"] = pd.to_datetime(work["日期"], errors="coerce")
    else:
        work["_日期"] = pd.NaT
    return work


def _project_filename(code: str, name: str, used: set[str]) -> str:
    code_part = safe_filename(code) if safe_text(code) else "未识别编码"
    name_part = safe_filename(name) if safe_text(name) else "未命名项目"
    base = f"{code_part}_{name_part}"
    candidate = f"{base}.xlsx"
    if candidate not in used:
        used.add(candidate)
        return candidate
    n = 2
    while f"{base}_{n}.xlsx" in used:
        n += 1
    candidate = f"{base}_{n}.xlsx"
    used.add(candidate)
    return candidate


def _text_or_blank(value) -> str:
    text = safe_text(value)
    return text


def _date_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def _number_or_blank(value):
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return None
    return float(number)


def _paint(cell, fill, font, alignment, border) -> None:
    cell.fill = fill
    cell.font = font
    cell.alignment = alignment
    cell.border = border


def _write_detail_sheet(wb: Workbook, project_df: pd.DataFrame, code: str, name: str):
    headers = ["一级科目", "二级科目", AMOUNT_HEADER, *LINE_COLUMNS]
    ws = wb.create_sheet(title=SUBJECT_WIDE_SHEET)
    ws.freeze_panes = "A4"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor="DDEFE3")
    header_fill = PatternFill("solid", fgColor="C9E1D1")
    summary_fill = PatternFill("solid", fgColor="EAF4EC")
    line_fill = PatternFill("solid", fgColor="F7FBF8")
    bold_font = Font(bold=True, color="1F3A2E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B7C7BC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    _paint(ws.cell(1, 1, "项目编码"), title_fill, bold_font, center, border)
    _paint(ws.cell(1, 2, code or None), title_fill, bold_font, center, border)
    _paint(ws.cell(2, 1, "项目名称"), title_fill, bold_font, center, border)
    _paint(ws.cell(2, 2, name or None), title_fill, bold_font, center, border)
    for col, header in enumerate(headers, start=1):
        _paint(ws.cell(3, col, header), title_fill, bold_font, center, border)

    amount_col = 3
    line_start_col = 4
    level1_order = sorted(project_df["一级科目"].unique().tolist(), key=level1_sort_key)
    row_idx = 4
    grand = 0.0

    for level1 in level1_order:
        level1_df = project_df[project_df["一级科目"] == level1]
        subtotal_row = row_idx
        level1_total = float(level1_df["付款金额"].sum())
        grand += level1_total
        for col, val in ((1, level1), (2, "小计")):
            _paint(ws.cell(subtotal_row, col, val), header_fill, bold_font, center, border)
        amount_cell = ws.cell(subtotal_row, amount_col, level1_total)
        amount_cell.number_format = "#,##0.00"
        _paint(amount_cell, header_fill, bold_font, center, border)
        for col in range(line_start_col, len(headers) + 1):
            _paint(ws.cell(subtotal_row, col, None), header_fill, bold_font, center, border)
        row_idx += 1
        block_start = row_idx

        level2_order = sorted(level1_df["二级科目"].unique().tolist(), key=level1_sort_key)
        line_groups: List[tuple[int, int]] = []
        for level2 in level2_order:
            level2_df = level1_df[level1_df["二级科目"] == level2]
            level2_df = level2_df.sort_values("_日期", kind="mergesort", na_position="last")
            level2_total = float(level2_df["付款金额"].sum())
            level2_row = row_idx
            for col, val in ((1, level1), (2, level2)):
                _paint(ws.cell(level2_row, col, val), summary_fill, bold_font, center, border)
            amount_cell = ws.cell(level2_row, amount_col, level2_total)
            amount_cell.number_format = "#,##0.00"
            _paint(amount_cell, summary_fill, bold_font, center, border)
            for col in range(line_start_col, len(headers) + 1):
                _paint(ws.cell(level2_row, col, None), summary_fill, bold_font, center, border)
            row_idx += 1
            line_start = row_idx
            for record in level2_df.to_dict("records"):
                values = {
                    "一级科目": level1,
                    "二级科目": level2,
                    AMOUNT_HEADER: float(record.get("付款金额") or 0),
                }
                for col_name in LINE_COLUMNS:
                    raw = record.get(col_name)
                    if col_name == "日期":
                        values[col_name] = _date_value(record.get("_日期"))
                    elif col_name in AMOUNT_LINE_COLUMNS or col_name == "汇率":
                        values[col_name] = _number_or_blank(raw)
                    else:
                        values[col_name] = _text_or_blank(raw) or None
                for col, header in enumerate(headers, start=1):
                    cell = ws.cell(row_idx, col, values.get(header))
                    cell.border = border
                    cell.fill = line_fill
                    cell.alignment = left if header in {"摘要", "供应商名称", "三级科目"} else center
                    if header in {AMOUNT_HEADER, *AMOUNT_LINE_COLUMNS}:
                        cell.number_format = "#,##0.00"
                    elif header == "汇率" and isinstance(cell.value, float):
                        cell.number_format = "0.0000"
                    elif header == "日期" and cell.value is not None:
                        cell.number_format = "yyyy-mm-dd"
                row_idx += 1
            if row_idx - 1 >= line_start:
                line_groups.append((line_start, row_idx - 1))
        if row_idx - 1 >= block_start:
            ws.row_dimensions.group(block_start, row_idx - 1, outline_level=1, hidden=False)
        for line_start, line_end in line_groups:
            ws.row_dimensions.group(line_start, line_end, outline_level=2, hidden=False)

    for col, val in ((1, "全部汇总"), (2, "-")):
        _paint(ws.cell(row_idx, col, val), header_fill, bold_font, center, border)
    total_cell = ws.cell(row_idx, amount_col, grand)
    total_cell.number_format = "#,##0.00"
    _paint(total_cell, header_fill, bold_font, center, border)
    for col in range(line_start_col, len(headers) + 1):
        _paint(ws.cell(row_idx, col, None), header_fill, bold_font, center, border)

    widths = {
        "一级科目": 18,
        "二级科目": 22,
        "付款金额": 16,
        "三级科目": 16,
        "日期": 14,
        "摘要": 42,
        "供应商名称": 24,
        "币别": 10,
        "汇率": 12,
        "年月": 12,
    }
    for col, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(header, 14)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 30
    return ws


def _write_project_workbook(path: Path, project_df: pd.DataFrame, code: str, name: str) -> None:
    payload = build_subject_wide_from_detail_df(project_df)
    wb = Workbook()
    default = wb.active
    write_subject_wide_sheet(wb, payload)
    wb[SUBJECT_WIDE_SHEET].title = SUMMARY_SHEET
    _write_detail_sheet(wb, project_df, code, name)
    if default.title == "Sheet":
        wb.remove(default)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _cleanup_group_dir(group_dir: Path, keep: set[str]) -> None:
    if not group_dir.is_dir():
        return
    for child in list(group_dir.iterdir()):
        if not child.is_file() or child.name.startswith("~$"):
            continue
        if child.suffix.lower() == ".xlsx" and child.name not in keep:
            try:
                child.unlink()
                print(f"  🗑️ 移除过期项目: {group_dir.name}/{child.name}")
            except OSError as exc:
                print(f"  ⚠️ 无法删除 {child.name}: {exc}")


def _cleanup_stale_groups(root: Path, active: set[str]) -> None:
    if not root.is_dir():
        return
    for child in list(root.iterdir()):
        if not child.is_dir() or child.name.startswith(".") or child.name in active:
            continue
        leftover = [p for p in child.iterdir() if p.name not in {".DS_Store"} and not p.name.startswith("~$")]
        xlsx = [p for p in leftover if p.is_file() and p.suffix.lower() == ".xlsx"]
        if len(xlsx) != len(leftover):
            continue
        for path in xlsx:
            try:
                path.unlink()
            except OSError:
                break
        else:
            try:
                ds = child / ".DS_Store"
                if ds.exists():
                    ds.unlink()
                child.rmdir()
                print(f"  🗑️ 移除过期分组目录: {child.name}")
            except OSError as exc:
                print(f"  ⚠️ 无法移除目录 {child.name}: {exc}")


def run_subject_wide_split(paths: Paths) -> List[str]:
    """A2 已分类明细 → 03_Sum/一级二级科目宽/{组名}/{项目}.xlsx。"""
    from .classify_raw_std import list_classified_groups, load_classified_detail

    root = subject_wide_split_root(paths)
    root.mkdir(parents=True, exist_ok=True)
    groups = list_classified_groups(paths)
    if not groups:
        print(f"未找到 A2 分类结果: {paths.classified_dir}")
        return []

    print(f"一级二级科目宽按项目拆分: {', '.join(name for name, _ in groups)}")
    print(f"输出: {root}/{{组名}}/{{项目编码}}_{{项目名称}}.xlsx")
    written: List[str] = []
    active: set[str] = set()
    for group_name, class_xlsx in groups:
        print(f"\n{'=' * 64}")
        print(f"  科目宽拆分 · {group_name}")
        print(f"{'=' * 64}")
        detail = _prepare_detail(load_classified_detail(class_xlsx))
        group_dir = root / group_name
        group_dir.mkdir(parents=True, exist_ok=True)
        active.add(group_name)
        if detail.empty:
            print(f"  [{group_name}] 没有付款金额，跳过")
            _cleanup_group_dir(group_dir, set())
            continue
        used: set[str] = set()
        keep: set[str] = set()
        projects = detail.groupby(["项目编码", "项目名称"], dropna=False, sort=False)
        ordered = sorted(projects.groups, key=lambda key: (safe_text(key[0]), safe_text(key[1])))
        for code, name in ordered:
            project_df = projects.get_group((code, name))
            filename = _project_filename(safe_text(code), safe_text(name), used)
            out_path = group_dir / filename
            _write_project_workbook(out_path, project_df, safe_text(code), safe_text(name))
            keep.add(filename)
            written.append(str(out_path))
            print(f"  ✅ {group_name}/{filename}  ({len(project_df)} 行)")
        _cleanup_group_dir(group_dir, keep)
        print(f"  ✅ [{group_name}] {len(keep)} 个项目 → {group_dir}")

    _cleanup_stale_groups(root, active)
    print(f"\n✅ 科目宽拆分完成：{len(written)} 个项目文件 → {root}")
    return written
