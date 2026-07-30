# -*- coding: utf-8 -*-
"""
项目明细 Excel · 分类汇总 / 分项跳转 / 明细表联动

能力：
1. 将「明细数据」格式化为 Excel 表「付款明细」（改一/二/三级后汇总公式自动变）
2. 「分类汇总」：一/二/三级结构 + SUMIFS 公式联动金额 + 超链接跳到明细对应行
3. 「分项查看」：选择一/二/三级后用 FILTER 列出该分项全部明细（Excel 365/2021+）
4. 「透视源」：干净数据表，便于插入 Excel 原生数据透视表
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

SUMMARY_SHEET = "分类汇总"
VIEW_SHEET = "分项查看"
PIVOT_SRC_SHEET = "透视源"
GUIDE_SHEET = "分类使用说明"
DETAIL_SHEET = "明细数据"
DETAIL_SHEET_CANDIDATES = ("明细数据", "明细数据_万泰铢")
TABLE_NAME = "付款明细"
PIVOT_TABLE_NAME = "透视明细"

_SUMMARY_HEADERS = [
    "一级科目",
    "二级科目",
    "三级科目",
    "笔数",
    "实付金额",
    "预扣税额",
    "实付+预扣税",
    "查看明细",
]

_PIVOT_COLS = [
    "一级科目",
    "二级科目",
    "三级科目",
    "实付金额",
    "预扣税额",
    "实付+预扣税",
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


def _ensure_col(ws: Worksheet, headers: Dict[str, int], name: str) -> int:
    if name in headers:
        return headers[name]
    idx = (ws.max_column or 0) + 1
    ws.cell(row=1, column=idx, value=name)
    headers[name] = idx
    return idx


def _remove_table_if_exists(ws: Worksheet, name: str) -> None:
    tables = getattr(ws, "tables", None)
    if not tables:
        return
    # openpyxl TableList supports del by name
    try:
        if name in tables:
            del tables[name]
    except Exception:
        try:
            keys = list(tables.keys()) if hasattr(tables, "keys") else []
            for k in keys:
                if str(k) == name:
                    del tables[k]
        except Exception:
            pass


def ensure_detail_excel_table(ws: Worksheet, table_name: str = TABLE_NAME) -> str:
    """把明细 sheet 做成 Excel 表，供 SUMIFS / 透视 / FILTER 引用。"""
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    if max_row < 1 or max_col < 1:
        return table_name

    # 规范空分类为空白字符串，便于公式
    headers = _header_map(ws)
    for col_name in ("一级科目", "二级科目", "三级科目"):
        if col_name not in headers:
            continue
        c = headers[col_name]
        for r in range(2, max_row + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None:
                continue
            s = str(val).strip()
            if not s or s.lower() in {"nan", "none"} or s == "(空白)":
                cell.value = None

    # 辅助列：实付+预扣税（写值，便于透视）
    pay_c = headers.get("实付金额")
    wht_c = headers.get("预扣税额")
    if pay_c and wht_c:
        sum_c = _ensure_col(ws, headers, "实付+预扣税")
        for r in range(2, max_row + 1):
            # 用公式联动明细自身
            pc = get_column_letter(pay_c)
            wc = get_column_letter(wht_c)
            ws.cell(row=r, column=sum_c).value = f"=IFERROR({pc}{r},0)+IFERROR({wc}{r},0)"
            ws.cell(row=r, column=sum_c).number_format = "#,##0.00"
        max_col = max(max_col, sum_c)

    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    _remove_table_if_exists(ws, table_name)
    # 表名全局唯一：同一工作簿不能重复
    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)
    return table_name


def _find_first_detail_row(
    ws: Worksheet,
    l1: str,
    l2: str,
    l3: str,
) -> Optional[int]:
    """在明细中找首条匹配行，供超链接定位。"""
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
            **{"实付+预扣税": ("_sum", "sum")},
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
                        "实付+预扣税": float(r["实付+预扣税"]),
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
                    "实付+预扣税": float(g2["实付+预扣税"].sum()),
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
                "实付+预扣税": float(g1["实付+预扣税"].sum()),
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
                "实付+预扣税": float(g["实付+预扣税"].sum()),
            }
        )
    return rows


def _sumifs_pay(table: str, l1_cell: str, l2_cell: str, l3_cell: str, amount_col: str) -> str:
    """生成 SUMIFS；三级 (空白) 匹配空单元格。"""
    # Excel: 空单元格用 "" 匹配
    # 三级条件：IF(C2="(空白)","",C2)
    l3_crit = f'IF({l3_cell}="(空白)","",{l3_cell})'
    return (
        f'IFERROR(SUMIFS({table}[{amount_col}],'
        f'{table}[一级科目],{l1_cell},'
        f'{table}[二级科目],{l2_cell},'
        f'{table}[三级科目],{l3_crit}),0)'
    )


def _sumif_l1(table: str, l1_cell: str, amount_col: str) -> str:
    return f'IFERROR(SUMIF({table}[一级科目],SUBSTITUTE({l1_cell}," 合计",""),{table}[{amount_col}]),0)'


def _sumifs_l2(table: str, l1_cell: str, l2_cell: str, amount_col: str) -> str:
    l2 = f'SUBSTITUTE({l2_cell}," 小计","")'
    return (
        f'IFERROR(SUMIFS({table}[{amount_col}],'
        f'{table}[一级科目],{l1_cell},'
        f'{table}[二级科目],{l2}),0)'
    )


def write_category_summary_sheet(
    wb: Workbook,
    rows: Sequence[Dict[str, Any]],
    detail_ws: Optional[Worksheet] = None,
    *,
    sheet_name: str = SUMMARY_SHEET,
    table_name: str = TABLE_NAME,
    use_formulas: bool = True,
) -> Worksheet:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    pos = None
    if DETAIL_SHEET in wb.sheetnames:
        pos = wb.sheetnames.index(DETAIL_SHEET) + 1
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

        ws.cell(row=i, column=1, value=l1).border = thin
        ws.cell(row=i, column=2, value=l2).border = thin
        ws.cell(row=i, column=3, value=l3 if l3 else None).border = thin

        # 金额：优先公式联动「付款明细」表（改明细分类后自动重算）
        if use_formulas and table_name and rtype == "detail":
            crit3 = f'IF(OR(C{i}="(空白)",C{i}=""),"",C{i})'
            n_f = (
                f'=IFERROR(COUNTIFS({table_name}[一级科目],A{i},'
                f'{table_name}[二级科目],B{i},'
                f'{table_name}[三级科目],{crit3}),0)'
            )
            pay_f = (
                f'=IFERROR(SUMIFS({table_name}[实付金额],'
                f'{table_name}[一级科目],A{i},'
                f'{table_name}[二级科目],B{i},'
                f'{table_name}[三级科目],{crit3}),0)'
            )
            wht_f = (
                f'=IFERROR(SUMIFS({table_name}[预扣税额],'
                f'{table_name}[一级科目],A{i},'
                f'{table_name}[二级科目],B{i},'
                f'{table_name}[三级科目],{crit3}),0)'
            )
            ws.cell(row=i, column=4, value=n_f)
            ws.cell(row=i, column=5, value=pay_f)
            ws.cell(row=i, column=6, value=wht_f)
            ws.cell(row=i, column=7, value=f"=E{i}+F{i}")
        elif use_formulas and table_name and rtype == "l2_sub":
            ws.cell(
                row=i,
                column=4,
                value=(
                    f'=IFERROR(COUNTIFS({table_name}[一级科目],A{i},'
                    f'{table_name}[二级科目],SUBSTITUTE(B{i}," 小计","")),0)'
                ),
            )
            ws.cell(
                row=i,
                column=5,
                value=(
                    f'=IFERROR(SUMIFS({table_name}[实付金额],'
                    f'{table_name}[一级科目],A{i},'
                    f'{table_name}[二级科目],SUBSTITUTE(B{i}," 小计","")),0)'
                ),
            )
            ws.cell(
                row=i,
                column=6,
                value=(
                    f'=IFERROR(SUMIFS({table_name}[预扣税额],'
                    f'{table_name}[一级科目],A{i},'
                    f'{table_name}[二级科目],SUBSTITUTE(B{i}," 小计","")),0)'
                ),
            )
            ws.cell(row=i, column=7, value=f"=E{i}+F{i}")
        elif use_formulas and table_name and rtype == "l1_sub":
            ws.cell(
                row=i,
                column=4,
                value=f'=IFERROR(COUNTIF({table_name}[一级科目],SUBSTITUTE(A{i}," 合计","")),0)',
            )
            ws.cell(
                row=i,
                column=5,
                value=f'=IFERROR(SUMIF({table_name}[一级科目],SUBSTITUTE(A{i}," 合计",""),{table_name}[实付金额]),0)',
            )
            ws.cell(
                row=i,
                column=6,
                value=f'=IFERROR(SUMIF({table_name}[一级科目],SUBSTITUTE(A{i}," 合计",""),{table_name}[预扣税额]),0)',
            )
            ws.cell(row=i, column=7, value=f"=E{i}+F{i}")
        elif use_formulas and table_name and rtype == "grand":
            ws.cell(row=i, column=4, value=f"=IFERROR(COUNTA({table_name}[实付金额]),0)")
            ws.cell(row=i, column=5, value=f"=IFERROR(SUM({table_name}[实付金额]),0)")
            ws.cell(row=i, column=6, value=f"=IFERROR(SUM({table_name}[预扣税额]),0)")
            ws.cell(row=i, column=7, value=f"=E{i}+F{i}")
        else:
            ws.cell(row=i, column=4, value=row.get("笔数", 0))
            ws.cell(row=i, column=5, value=row.get("实付金额", 0))
            ws.cell(row=i, column=6, value=row.get("预扣税额", 0))
            ws.cell(row=i, column=7, value=row.get("实付+预扣税", 0))

        for col in range(4, 8):
            ws.cell(row=i, column=col).border = thin
            if col >= 5:
                ws.cell(row=i, column=col).number_format = money_fmt

        # 样式
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

        # 超链接：跳到明细首条匹配
        link_cell = ws.cell(row=i, column=8)
        link_cell.border = thin
        if detail_ws is not None and rtype != "grand":
            target_row = _find_first_detail_row(detail_ws, str(l1), str(l2), str(l3 or "(空白)"))
            if target_row:
                # 分项查看：写入参数并跳转；同时超链接到明细行
                link_cell.value = "定位明细"
                link_cell.hyperlink = f"#'{DETAIL_SHEET}'!A{target_row}"
                link_cell.font = link_font
            else:
                link_cell.value = ""
        else:
            link_cell.value = ""

        ws.row_dimensions[i].outline_level = min(outline, 7)

    ws.freeze_panes = "A2"
    last = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:H{last}"
    ws.sheet_properties.outlinePr.summaryBelow = True
    for i, w in enumerate([16, 20, 16, 8, 14, 14, 14, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    note = len(rows) + 3
    ws.cell(
        row=note,
        column=1,
        value=(
            "【联动】金额为公式，引用表「付款明细」。在「明细数据」改一级/二级/三级或金额后，本表自动重算。"
            "【跳转】点「定位明细」跳到该分类第一条明细，可改科目；改完后汇总自动变。"
            "【透视】请到「透视源」sheet，选中表 → 插入 → 数据透视表。"
        ),
    )
    ws.merge_cells(start_row=note, start_column=1, end_row=note, end_column=8)
    return ws


def write_pivot_source_sheet(wb: Workbook, detail_ws: Worksheet) -> Worksheet:
    """干净数据区 + Excel 表，方便用户一键插数据透视表。"""
    if PIVOT_SRC_SHEET in wb.sheetnames:
        del wb[PIVOT_SRC_SHEET]
    ws = wb.create_sheet(PIVOT_SRC_SHEET)
    headers = _header_map(detail_ws)
    # 选用列
    cols: List[Tuple[str, Optional[int]]] = []
    for name in _PIVOT_COLS:
        cols.append((name, headers.get(name)))

    # 若无 实付+预扣税 列，后面用公式
    for c, (name, _) in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=name)

    out_r = 2
    max_r = detail_ws.max_row or 1
    for r in range(2, max_r + 1):
        # 跳过全空行
        any_val = False
        for src_c in headers.values():
            if detail_ws.cell(row=r, column=src_c).value not in (None, ""):
                any_val = True
                break
        if not any_val:
            continue
        for c, (name, src_idx) in enumerate(cols, start=1):
            if name == "实付+预扣税":
                # 从源行取实付、预扣
                pay_i = headers.get("实付金额")
                wht_i = headers.get("预扣税额")
                pay = detail_ws.cell(row=r, column=pay_i).value if pay_i else 0
                wht = detail_ws.cell(row=r, column=wht_i).value if wht_i else 0
                try:
                    pv = float(pay or 0)
                except Exception:
                    pv = 0.0
                try:
                    wv = float(wht or 0)
                except Exception:
                    wv = 0.0
                cell = ws.cell(row=out_r, column=c, value=pv + wv)
                cell.number_format = "#,##0.00"
            elif src_idx:
                val = detail_ws.cell(row=r, column=src_idx).value
                if name in ("一级科目", "二级科目", "三级科目"):
                    val = _cell_text(val) or "(空白)"
                cell = ws.cell(row=out_r, column=c, value=val)
                if name in ("实付金额", "预扣税额"):
                    cell.number_format = "#,##0.00"
            else:
                ws.cell(row=out_r, column=c, value=None)
        out_r += 1

    last = max(1, out_r - 1)
    ref = f"A1:{get_column_letter(len(cols))}{last}"
    _remove_table_if_exists(ws, PIVOT_TABLE_NAME)
    if last >= 1:
        tab = Table(displayName=PIVOT_TABLE_NAME, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tab)

    for i, w in enumerate([14, 16, 14, 12, 12, 12, 24, 18, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.cell(
        row=last + 2,
        column=1,
        value=(
            "【做数据透视表】选中上方表格任意单元格 → 插入 → 数据透视表 → "
            "行：一级科目/二级科目/三级科目 → 值：实付金额、预扣税额、实付+预扣税（求和）。"
            "改「明细数据」分类后：数据 → 全部刷新，透视表同步更新。"
        ),
    )
    return ws


def write_item_view_sheet(wb: Workbook, table_name: str = TABLE_NAME) -> Worksheet:
    """
    分项查看：在 B2/B3/B4 选一/二/三级，下方 FILTER 列出全部匹配明细。
    需要 Excel 365 / 2021+ 动态数组。
    """
    if VIEW_SHEET in wb.sheetnames:
        del wb[VIEW_SHEET]
    ws = wb.create_sheet(VIEW_SHEET)

    ws["A1"] = "分项查看（改科目请回「明细数据」改，本页自动跟着变）"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:F1")

    ws["A2"] = "一级科目"
    ws["A3"] = "二级科目"
    ws["A4"] = "三级科目"
    ws["B2"] = ""
    ws["B3"] = ""
    ws["B4"] = ""
    for r in (2, 3, 4):
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.cell(row=r, column=2).border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

    # 下拉：来自明细表唯一值（用公式 UNIQUE，365+）
    # 兼容：写死从明细数据列提取的 validation 在 refresh 时用数据验证列表
    ws["D2"] = "提示：可从「分类汇总」复制一/二/三级粘贴到 B2:B4"
    ws["A6"] = "匹配明细（Excel 365/2021+ 自动列出；旧版请用「透视源」做透视表筛选）"
    ws["A6"].font = Font(bold=True)

    # FILTER 动态数组：三级空白兼容
    # =FILTER(付款明细, (一级=B2)*(二级=B3)*((B4="")+((B4="(空白)")*(三级=""))+(三级=B4)))
    filt = (
        f'=IFERROR(FILTER({table_name},'
        f'({table_name}[一级科目]=B2)*'
        f'({table_name}[二级科目]=B3)*'
        f'(((B4="")+(B4="(空白)"))*(({table_name}[三级科目]="")+({table_name}[三级科目]="(空白)"))'
        f'+({table_name}[三级科目]=B4))),'
        f'"（无匹配行，或请用 Excel 365；可检查 B2/B3/B4）")'
    )
    ws["A7"] = filt

    ws["A5"] = "合计实付+预扣税"
    ws["B5"] = (
        f'=IFERROR(SUMIFS({table_name}[实付金额],{table_name}[一级科目],B2,{table_name}[二级科目],B3,'
        f'{table_name}[三级科目],IF(OR(B4="",B4="(空白)"),"",B4)),0)'
        f'+IFERROR(SUMIFS({table_name}[预扣税额],{table_name}[一级科目],B2,{table_name}[二级科目],B3,'
        f'{table_name}[三级科目],IF(OR(B4="",B4="(空白)"),"",B4)),0)'
    )
    ws["B5"].number_format = "#,##0.00"
    ws["B5"].font = Font(bold=True, color="C00000")

    for i, w in enumerate([18, 22, 16, 50], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def write_guide_sheet(wb: Workbook) -> Worksheet:
    if GUIDE_SHEET in wb.sheetnames:
        del wb[GUIDE_SHEET]
    ws = wb.create_sheet(GUIDE_SHEET, 0)
    lines = [
        "财务付款 · 分类汇总使用说明",
        "",
        "一、改科目（会联动汇总）",
        "  1. 打开 sheet「明细数据」",
        "  2. 直接改「一级科目 / 二级科目 / 三级科目」列（有下拉可选）",
        "  3. 回到「分类汇总」：金额是公式，自动按新分类重算（若未变请按 F9 或保存重开）",
        "",
        "二、从汇总跳到明细核对",
        "  1. 在「分类汇总」点「定位明细」→ 跳到该分类在明细中的第一条",
        "  2. 或：把一/二/三级复制到「分项查看」的 B2/B3/B4 → 下方列出该分项全部行（需 Excel 365/2021）",
        "",
        "三、用 Excel 自带数据透视表",
        "  1. 打开「透视源」sheet",
        "  2. 点表格内任意格 → 菜单【插入】→【数据透视表】",
        "  3. 行字段：一级科目、二级科目、三级科目",
        "  4. 值字段：实付金额、预扣税额、实付+预扣税（求和）",
        "  5. 改完明细分类后：【数据】→【全部刷新】",
        "",
        "四、注意",
        "  · 请在「明细数据」改分类，不要在汇总表改分类名称（汇总表是统计结果）",
        "  · 表名「付款明细」供公式引用，勿删",
        "  · 重新跑 std / reclassify / category-summary 会刷新这些 sheet",
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
                if val is not None and str(val).strip() != "":
                    empty = False
                row[h] = val
            if not empty:
                data.append(row)
        if data:
            return pd.DataFrame(data)
    return None


def add_category_summary_to_workbook(
    wb: Workbook,
    detail_df: Optional[pd.DataFrame] = None,
    *,
    sheet_name: str = SUMMARY_SHEET,
) -> int:
    """写入：明细表 + 分类汇总(公式+跳转) + 分项查看 + 透视源 + 说明。"""
    # 明细 sheet
    if DETAIL_SHEET not in wb.sheetnames:
        # 尝试用万泰铢
        if "明细数据_万泰铢" in wb.sheetnames:
            detail_ws = wb["明细数据_万泰铢"]
        else:
            return 0
    else:
        detail_ws = wb[DETAIL_SHEET]

    # 唯一表名（多 sheet 时避免冲突）：付款明细
    # 若工作簿已有同名表在其他 sheet，先清
    for ws in wb.worksheets:
        _remove_table_if_exists(ws, TABLE_NAME)
        _remove_table_if_exists(ws, PIVOT_TABLE_NAME)

    ensure_detail_excel_table(detail_ws, TABLE_NAME)

    if detail_df is None:
        detail_df = _read_detail_df_from_wb(wb)
    if detail_df is None:
        detail_df = pd.DataFrame()

    rows = build_category_summary_rows(detail_df)
    write_category_summary_sheet(
        wb,
        rows,
        detail_ws=detail_ws,
        sheet_name=sheet_name,
        table_name=TABLE_NAME,
        use_formulas=True,
    )
    write_item_view_sheet(wb, TABLE_NAME)
    write_pivot_source_sheet(wb, detail_ws)
    write_guide_sheet(wb)
    return len(rows)


def refresh_category_summary_in_file(excel_path: Path) -> int:
    path = Path(excel_path)
    wb = load_workbook(path)
    n = add_category_summary_to_workbook(wb)
    wb.save(path)
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
