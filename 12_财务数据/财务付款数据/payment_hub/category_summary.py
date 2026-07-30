# -*- coding: utf-8 -*-
"""
项目明细 Excel · 分类汇总（稳健版，避免 Excel「修复工作簿」）

能力（不依赖易损坏的 Table 结构化公式 / FILTER）：
1. 「分类汇总」：一/二/三级 + 实付/预扣/合计（静态值，可大纲折叠）
2. 「定位明细」超链接 → 跳到明细首条
3. 「透视源」：干净数据区，便于插入原生数据透视表
4. 「分类使用说明」

改明细分类后请重新运行 category-summary / reclassify 刷新汇总数字。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence
from zipfile import ZipFile, ZIP_DEFLATED
import io
import re
import shutil
import tempfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

SUMMARY_SHEET = "分类汇总"
PIVOT_SRC_SHEET = "透视源"
GUIDE_SHEET = "分类使用说明"
VIEW_SHEET = "分项查看"  # 保留兼容；改为说明页，去掉 FILTER
DETAIL_SHEET = "明细数据"
DETAIL_SHEET_CANDIDATES = ("明细数据", "明细数据_万泰铢")

_SUMMARY_HEADERS = [
    "一级科目",
    "二级科目",
    "三级科目",
    "笔数",
    "实付金额",
    "预扣税额",
    "实付加预扣税",
    "查看明细",
]

_PIVOT_COLS = [
    "一级科目",
    "二级科目",
    "三级科目",
    "实付金额",
    "预扣税额",
    "实付加预扣税",
    "摘要",
    "供应商名称",
    "日期",
    "科目",
]


def _blank_label(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "(空白)"
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return "(空白)"
    return s


def _cell_text(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return ""
    return s


def _to_float(v: Any) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def _header_map(ws: Worksheet) -> Dict[str, int]:
    m: Dict[str, int] = {}
    for c in range(1, (ws.max_column or 0) + 1):
        v = ws.cell(row=1, column=c).value
        if v is None:
            continue
        m[str(v).strip()] = c
    return m


def _strip_all_tables(wb: Workbook) -> None:
    """移除全部 Excel 表，避免绝对路径/中文表名导致修复提示。"""
    for ws in wb.worksheets:
        tables = getattr(ws, "tables", None)
        if not tables:
            continue
        try:
            names = list(tables.keys())
        except Exception:
            names = []
        for name in names:
            try:
                del tables[name]
            except Exception:
                pass


def build_category_summary_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df is None or df.empty:
        return []

    d = df.copy()
    for col in ("一级科目", "二级科目", "三级科目"):
        if col not in d.columns:
            d[col] = ""
        d[col] = d[col].map(_blank_label)

    pay = _num(d["实付金额"]) if "实付金额" in d.columns else pd.Series(0.0, index=d.index)
    wht = _num(d["预扣税额"]) if "预扣税额" in d.columns else pd.Series(0.0, index=d.index)
    d = d.assign(_pay=pay, _wht=wht, _sum=pay + wht, _n=1)

    g = (
        d.groupby(["一级科目", "二级科目", "三级科目"], dropna=False, sort=False)
        .agg(
            笔数=("_n", "sum"),
            实付金额=("_pay", "sum"),
            预扣税额=("_wht", "sum"),
            实付加预扣税=("_sum", "sum"),
        )
        .reset_index()
    )
    g = g.sort_values(["一级科目", "二级科目", "三级科目"], kind="mergesort").reset_index(
        drop=True
    )

    rows: List[Dict[str, Any]] = []
    for l1, g1 in g.groupby("一级科目", sort=False):
        for l2, g2 in g1.groupby("二级科目", sort=False):
            for _, r in g2.iterrows():
                rows.append(
                    {
                        "row_type": "detail",
                        "outline": 2,
                        "一级科目": str(l1),
                        "二级科目": str(l2),
                        "三级科目": str(r["三级科目"]),
                        "笔数": int(r["笔数"]),
                        "实付金额": float(r["实付金额"]),
                        "预扣税额": float(r["预扣税额"]),
                        "实付加预扣税": float(r["实付加预扣税"]),
                    }
                )
            rows.append(
                {
                    "row_type": "l2_sub",
                    "outline": 1,
                    "一级科目": str(l1),
                    "二级科目": f"{l2} 小计",
                    "三级科目": "",
                    "笔数": int(g2["笔数"].sum()),
                    "实付金额": float(g2["实付金额"].sum()),
                    "预扣税额": float(g2["预扣税额"].sum()),
                    "实付加预扣税": float(g2["实付加预扣税"].sum()),
                }
            )
        rows.append(
            {
                "row_type": "l1_sub",
                "outline": 0,
                "一级科目": f"{l1} 合计",
                "二级科目": "",
                "三级科目": "",
                "笔数": int(g1["笔数"].sum()),
                "实付金额": float(g1["实付金额"].sum()),
                "预扣税额": float(g1["预扣税额"].sum()),
                "实付加预扣税": float(g1["实付加预扣税"].sum()),
            }
        )

    if rows:
        rows.append(
            {
                "row_type": "grand",
                "outline": 0,
                "一级科目": "总计",
                "二级科目": "",
                "三级科目": "",
                "笔数": int(g["笔数"].sum()),
                "实付金额": float(g["实付金额"].sum()),
                "预扣税额": float(g["预扣税额"].sum()),
                "实付加预扣税": float(g["实付加预扣税"].sum()),
            }
        )
    return rows


def _find_first_detail_row(
    ws: Worksheet,
    l1: str,
    l2: str,
    l3: str,
) -> Optional[int]:
    headers = _header_map(ws)
    c1 = headers.get("一级科目")
    c2 = headers.get("二级科目")
    c3 = headers.get("三级科目")
    if not c1:
        return 2 if (ws.max_row or 0) >= 2 else None

    def norm(v: Any) -> str:
        t = _cell_text(v)
        return t if t else "(空白)"

    mode = "detail"
    want_l1 = str(l1 or "").strip()
    want_l2 = str(l2 or "").strip()
    want_l3 = str(l3 or "").strip() or "(空白)"
    if want_l1 == "总计":
        return 2 if (ws.max_row or 0) >= 2 else None
    if want_l1.endswith("合计"):
        mode = "l1"
        want_l1 = want_l1[: -len("合计")].strip()
    elif want_l2.endswith("小计"):
        mode = "l2"
        want_l2 = want_l2[: -len("小计")].strip()

    for r in range(2, (ws.max_row or 1) + 1):
        v1 = norm(ws.cell(row=r, column=c1).value)
        v2 = norm(ws.cell(row=r, column=c2).value) if c2 else "(空白)"
        v3 = norm(ws.cell(row=r, column=c3).value) if c3 else "(空白)"
        if v1 != want_l1:
            continue
        if mode == "l1":
            return r
        if v2 != want_l2:
            continue
        if mode == "l2":
            return r
        if v3 == want_l3:
            return r
    return None


def write_category_summary_sheet(
    wb: Workbook,
    rows: Sequence[Dict[str, Any]],
    detail_ws: Optional[Worksheet] = None,
    *,
    sheet_name: str = SUMMARY_SHEET,
) -> Worksheet:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    pos = wb.sheetnames.index(DETAIL_SHEET) + 1 if DETAIL_SHEET in wb.sheetnames else None
    ws = wb.create_sheet(sheet_name, pos)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    l1_fill = PatternFill("solid", fgColor="D6DCE4")
    l2_fill = PatternFill("solid", fgColor="E2EFDA")
    grand_fill = PatternFill("solid", fgColor="FFF2CC")
    link_font = Font(color="0563C1", underline="single")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    money_fmt = "#,##0.00"
    center = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(_SUMMARY_HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin

    for i, row in enumerate(rows, start=2):
        rtype = row.get("row_type", "detail")
        outline = int(row.get("outline", 2))
        l1 = row.get("一级科目", "")
        l2 = row.get("二级科目", "")
        l3 = row.get("三级科目", "")

        vals = [
            l1,
            l2,
            l3 if l3 else None,
            row.get("笔数", 0),
            row.get("实付金额", 0.0),
            row.get("预扣税额", 0.0),
            row.get("实付加预扣税", 0.0),
        ]
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = thin
            if col >= 5:
                cell.number_format = money_fmt

        fill = None
        bold = False
        if rtype == "l2_sub":
            fill, bold = l2_fill, True
        elif rtype == "l1_sub":
            fill, bold = l1_fill, True
        elif rtype == "grand":
            fill, bold = grand_fill, True
        if fill or bold:
            for col in range(1, 9):
                cell = ws.cell(row=i, column=col)
                if bold:
                    cell.font = Font(bold=True, size=12 if rtype == "grand" else 11)
                if fill:
                    cell.fill = fill

        link_cell = ws.cell(row=i, column=8)
        link_cell.border = thin
        if detail_ws is not None and rtype != "grand":
            target_row = _find_first_detail_row(
                detail_ws, str(l1), str(l2), str(l3 or "(空白)")
            )
            if target_row:
                link_cell.value = "定位明细"
                # 相对工作表定位，兼容性好
                link_cell.hyperlink = f"#{DETAIL_SHEET}!A{target_row}"
                link_cell.font = link_font

        ws.row_dimensions[i].outline_level = min(outline, 7)

    ws.freeze_panes = "A2"
    last = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:H{last}"
    try:
        ws.sheet_properties.outlinePr.summaryBelow = True
    except Exception:
        pass
    for i, w in enumerate([16, 20, 16, 8, 14, 14, 14, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    note = len(rows) + 3
    ws.cell(
        row=note,
        column=1,
        value=(
            "【定位】点「定位明细」跳到该分类在「明细数据」中的第一条，可改一/二/三级。"
            "【刷新】改完分类后请重新运行 category-summary 或 reclassify 以刷新本表数字。"
            "【透视】到「透视源」选中数据 → 插入 → 数据透视表（行=一二三级，值=金额求和）。"
        ),
    )
    ws.merge_cells(start_row=note, start_column=1, end_row=note, end_column=8)
    return ws


def write_pivot_source_sheet(wb: Workbook, detail_ws: Worksheet) -> Worksheet:
    if PIVOT_SRC_SHEET in wb.sheetnames:
        del wb[PIVOT_SRC_SHEET]
    ws = wb.create_sheet(PIVOT_SRC_SHEET)
    headers = _header_map(detail_ws)

    for c, name in enumerate(_PIVOT_COLS, start=1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E2F3")

    out_r = 2
    for r in range(2, (detail_ws.max_row or 1) + 1):
        any_val = False
        for src_c in headers.values():
            if detail_ws.cell(row=r, column=src_c).value not in (None, ""):
                any_val = True
                break
        if not any_val:
            continue
        pay = _to_float(
            detail_ws.cell(row=r, column=headers["实付金额"]).value
            if "实付金额" in headers
            else 0
        )
        wht = _to_float(
            detail_ws.cell(row=r, column=headers["预扣税额"]).value
            if "预扣税额" in headers
            else 0
        )
        for c, name in enumerate(_PIVOT_COLS, start=1):
            if name == "实付加预扣税":
                cell = ws.cell(row=out_r, column=c, value=pay + wht)
                cell.number_format = "#,##0.00"
            elif name in headers:
                val = detail_ws.cell(row=r, column=headers[name]).value
                if name in ("一级科目", "二级科目", "三级科目"):
                    val = _cell_text(val) or "(空白)"
                cell = ws.cell(row=out_r, column=c, value=val)
                if name in ("实付金额", "预扣税额"):
                    cell.number_format = "#,##0.00"
            else:
                ws.cell(row=out_r, column=c, value=None)
        out_r += 1

    last = max(1, out_r - 1)
    # 不用 Excel Table，避免修复提示；用户可选中 A1:J{last} 做透视
    for i, w in enumerate([14, 16, 14, 12, 12, 12, 24, 18, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.cell(
        row=last + 2,
        column=1,
        value=(
            f"【数据透视表】选中 A1:{get_column_letter(len(_PIVOT_COLS))}{last} → "
            "插入 → 数据透视表 → 行：一级/二级/三级科目 → 值：实付金额、预扣税额、实付加预扣税（求和）。"
            "改完「明细数据」分类后，请重新运行 category-summary 刷新本页，再对透视表点「刷新」。"
        ),
    )
    return ws


def write_item_view_sheet(wb: Workbook) -> Worksheet:
    """分项查看：改为操作说明（不再写 FILTER，避免 Excel 修复提示）。"""
    if VIEW_SHEET in wb.sheetnames:
        del wb[VIEW_SHEET]
    ws = wb.create_sheet(VIEW_SHEET)
    lines = [
        "分项查看（安全版）",
        "",
        "因部分 Excel 对 FILTER/结构化表公式会提示「修复工作簿」，本页改为操作指引：",
        "",
        "方法 A · 从分类汇总跳转",
        "  1. 打开「分类汇总」",
        "  2. 点击对应行的「定位明细」→ 跳到「明细数据」该分类第一条",
        "  3. 在明细中用筛选：数据 → 筛选 → 按一级/二级/三级筛选查看全部分项",
        "",
        "方法 B · 数据透视表切片",
        "  1. 打开「透视源」",
        "  2. 选中数据区 → 插入 → 数据透视表",
        "  3. 用透视表的筛选/切片器查看某一二级分类",
        "",
        "改科目：请在「明细数据」改一级/二级/三级，然后运行：",
        "  python 财务付款一体化.py category-summary",
        "或 reclassify，以刷新分类汇总与透视源。",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=line)
        if i == 1:
            ws.cell(row=i, column=1).font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 88
    return ws


def write_guide_sheet(wb: Workbook) -> Worksheet:
    if GUIDE_SHEET in wb.sheetnames:
        del wb[GUIDE_SHEET]
    ws = wb.create_sheet(GUIDE_SHEET, 0)
    lines = [
        "财务付款 · 分类汇总使用说明",
        "",
        "一、改科目",
        "  在「明细数据」修改「一级科目 / 二级科目 / 三级科目」（有下拉）",
        "  然后运行 category-summary 或 reclassify 刷新「分类汇总」「透视源」",
        "",
        "二、从汇总跳到明细核对",
        "  在「分类汇总」点「定位明细」→ 跳到该分类第一条明细",
        "  再用 Excel 筛选查看该分类全部行",
        "",
        "三、用 Excel 自带数据透视表",
        "  打开「透视源」→ 选中数据 → 插入 → 数据透视表",
        "  行：一级/二级/三级；值：实付金额、预扣税额、实付加预扣税",
        "",
        "四、若仍提示修复工作簿",
        "  运行：python 财务付款一体化.py category-summary",
        "  会重写为兼容性更好的结构（已去掉易损坏的 Table/FILTER 公式）",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=line)
        if i == 1:
            ws.cell(row=i, column=1).font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 90
    return ws


def _read_detail_df_from_wb(wb: Workbook) -> Optional[pd.DataFrame]:
    for name in DETAIL_SHEET_CANDIDATES:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        if (ws.max_row or 0) < 2:
            continue
        headers = []
        for c in range(1, (ws.max_column or 0) + 1):
            v = ws.cell(row=1, column=c).value
            headers.append(str(v).strip() if v is not None else f"col{c}")
        data = []
        for r in range(2, (ws.max_row or 0) + 1):
            row = {}
            empty = True
            for c, h in enumerate(headers, start=1):
                val = ws.cell(row=r, column=c).value
                # 跳过公式列污染：若表头是旧的「实付+预扣税」且是公式，读值即可
                if val is not None and str(val).strip() != "":
                    empty = False
                row[h] = val
            if not empty:
                data.append(row)
        if data:
            return pd.DataFrame(data)
    return None


def _clear_detail_formula_columns(detail_ws: Worksheet) -> None:
    """去掉明细中易出问题的公式辅助列（实付+预扣税 / 实付加预扣税）。"""
    headers = _header_map(detail_ws)
    # 从右往左删列，避免索引偏移
    drop_cols = sorted(
        [
            headers[name]
            for name in ("实付+预扣税", "实付加预扣税")
            if name in headers
        ],
        reverse=True,
    )
    for col in drop_cols:
        detail_ws.delete_cols(col, 1)
    # 清理表头为空的尾列
    max_c = detail_ws.max_column or 1
    while max_c > 1 and detail_ws.cell(row=1, column=max_c).value in (None, ""):
        detail_ws.delete_cols(max_c, 1)
        max_c -= 1


def add_category_summary_to_workbook(
    wb: Workbook,
    detail_df: Optional[pd.DataFrame] = None,
    *,
    sheet_name: str = SUMMARY_SHEET,
) -> int:
    _strip_all_tables(wb)

    if DETAIL_SHEET in wb.sheetnames:
        detail_ws = wb[DETAIL_SHEET]
    elif "明细数据_万泰铢" in wb.sheetnames:
        detail_ws = wb["明细数据_万泰铢"]
    else:
        return 0

    _clear_detail_formula_columns(detail_ws)

    if detail_df is None:
        detail_df = _read_detail_df_from_wb(wb)
    if detail_df is None:
        detail_df = pd.DataFrame()

    # 去掉辅助列再汇总
    drop_cols = [c for c in detail_df.columns if c in ("实付+预扣税", "实付加预扣税")]
    if drop_cols:
        detail_df = detail_df.drop(columns=drop_cols, errors="ignore")

    rows = build_category_summary_rows(detail_df)
    write_category_summary_sheet(wb, rows, detail_ws=detail_ws, sheet_name=sheet_name)
    write_item_view_sheet(wb)
    write_pivot_source_sheet(wb, detail_ws)
    write_guide_sheet(wb)
    return len(rows)


def _fix_xlsx_table_paths(path: Path) -> None:
    """把 table 关系从绝对路径 /xl/tables/ 改为相对路径 ../tables/。"""
    path = Path(path)
    tmp = path.with_suffix(".xlsx.tmpfix")
    with ZipFile(path, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/_rels/") and item.filename.endswith(
                ".rels"
            ):
                text = data.decode("utf-8")
                text2 = text.replace('Target="/xl/tables/', 'Target="../tables/')
                data = text2.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def refresh_category_summary_in_file(excel_path: Path) -> int:
    path = Path(excel_path)
    wb = load_workbook(path)
    n = add_category_summary_to_workbook(wb)
    wb.save(path)
    try:
        _fix_xlsx_table_paths(path)
    except Exception:
        pass
    return n


def refresh_category_summary_in_std_dir(std_dir: Path) -> Dict[str, int]:
    std_dir = Path(std_dir)
    stats = {"files": 0, "rows": 0, "errors": 0}
    for f in sorted(std_dir.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        try:
            n = refresh_category_summary_in_file(f)
            stats["files"] += 1
            stats["rows"] += n
        except Exception as exc:
            stats["errors"] += 1
            print(f"  ⚠️ 分类汇总失败 {f.name}: {exc}")
    return stats


def repair_std_workbook(excel_path: Path) -> bool:
    """
    修复单个可能损坏的工作簿：
    - 去掉 Excel Table
    - 去掉易损坏公式辅助列
    - 重写分类汇总/透视源/说明
    """
    path = Path(excel_path)
    try:
        refresh_category_summary_in_file(path)
        return True
    except Exception as exc:
        print(f"  ❌ 修复失败 {path.name}: {exc}")
        return False
