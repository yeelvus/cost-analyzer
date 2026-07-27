#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按项目拆分 D_sum/采购部采购台账_std_sum.xlsx（已含标准化、汇率换算等）
→ E_采购部采购台账_std_split/{项目编码}_{项目名称}_采购数据.xlsx

默认数据源为最新 std_sum「合并数据」sheet（含汇率填入、含税金额换算及绿色填充）。
拆分时保留全部列与绿色高亮样式。

用法：
  python3 流水线/06_按项目拆分.py
  python3 流水线/06_按项目拆分.py --input ... --output-dir ...
  python3 流水线/06_按项目拆分.py --from-raw-sum   # 兼容旧行为：从 C_raw_sum 拆
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

_PKG_ROOT = Path(__file__).resolve().parents[1]
_CODE_ROOT = Path(__file__).resolve().parents[2]
for _p in (str(_PKG_ROOT), str(_CODE_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from common.excel import apply_common_excel_layout

from _采购路径 import (
    RAW_SUM_DIR,
    RAW_SUM_FILE_NAME,
    STD_SPLIT_DIR,
    STD_SUM_FILE_NAME,
    SUM_DIR,
    latest_sum_file,
)

# 与 04b 非泰铢换算标绿一致
_FX_GREEN_RGB = "C6EFCE"
_FX_GREEN_FILL = PatternFill(start_color=_FX_GREEN_RGB, end_color=_FX_GREEN_RGB, fill_type="solid")
_FX_GREEN_FONT = Font(color="006100")


def _safe_filename_part(text: str, *, max_len: int = 40) -> str:
    s = str(text or "").strip()
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s)
    s = re.sub(r"\s+", "", s)
    s = s.strip("._ ")
    if not s:
        s = "未命名"
    return s[:max_len]


def _s(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none") else t


def _cell_rgb(cell) -> str | None:
    """读取单元格填充色 RGB（去 alpha 前缀）。"""
    try:
        fill = cell.fill
        if fill is None or not fill.fill_type or fill.fill_type == "none":
            return None
        color = fill.fgColor
        if color is None:
            return None
        rgb = getattr(color, "rgb", None)
        if not rgb or not isinstance(rgb, str):
            return None
        s = rgb.upper()
        # openpyxl 常带 ARGB：00C6EFCE
        if len(s) == 8 and s.startswith("00"):
            s = s[2:]
        elif len(s) == 8:
            s = s[2:]
        return s
    except Exception:
        return None


def _is_fx_green_cell(cell) -> bool:
    rgb = _cell_rgb(cell)
    return bool(rgb and rgb.endswith(_FX_GREEN_RGB))


def _default_std_sum_path() -> Path:
    """优先固定名 std_sum，否则 SUM_DIR 下最新 *std_sum*。"""
    preferred = SUM_DIR / STD_SUM_FILE_NAME
    if preferred.is_file():
        return preferred
    latest = latest_sum_file("*std_sum*.xlsx")
    if latest is not None:
        return latest
    return preferred


def _load_merged_sheet_rows(path: Path) -> tuple[list[str], list[dict]]:
    """
    用 openpyxl 读「合并数据」，保留每格是否绿色高亮。

    返回:
      headers: 列名列表
      rows: [{ "values": [...], "green_cols": set[int 0-based], "code", "name" }, ...]
    """
    wb = load_workbook(path, data_only=False)
    try:
        if "合并数据" in wb.sheetnames:
            ws = wb["合并数据"]
        else:
            ws = wb.active
        if ws is None:
            raise ValueError(f"工作簿无活动表: {path}")

        max_col = ws.max_column or 1
        headers: list[str] = []
        for c in range(1, max_col + 1):
            h = ws.cell(1, c).value
            headers.append(str(h).strip() if h is not None else f"列{c}")

        # 定位项目列
        def _find_col(*names: str) -> int | None:
            for i, h in enumerate(headers):
                if h in names:
                    return i
            return None

        code_i = _find_col("项目编码")
        name_i = _find_col("项目名称")

        rows: list[dict] = []
        for r in range(2, (ws.max_row or 1) + 1):
            values: list = []
            green_cols: set[int] = set()
            all_empty = True
            for c in range(1, len(headers) + 1):
                cell = ws.cell(r, c)
                v = cell.value
                values.append(v)
                if v is not None and str(v).strip() != "":
                    all_empty = False
                if _is_fx_green_cell(cell):
                    green_cols.add(c - 1)
            if all_empty:
                continue
            code = _s(values[code_i]) if code_i is not None else ""
            name = _s(values[name_i]) if name_i is not None else ""
            rows.append(
                {
                    "values": values,
                    "green_cols": green_cols,
                    "code": code,
                    "name": name,
                    "key": code or name or "未知项目",
                }
            )
        return headers, rows
    finally:
        wb.close()


def _write_project_excel(
    headers: list[str],
    project_rows: list[dict],
    out_path: Path,
) -> None:
    """写出单项目文件，保留汇率/含税金额等绿色填充。"""
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 去掉原序号，重编
    drop_seq = "序号" in headers
    seq_src_i = headers.index("序号") if drop_seq else None
    out_headers = ["序号"] + [
        h for i, h in enumerate(headers) if seq_src_i is None or i != seq_src_i
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "采购数据"
    ws.append(out_headers)

    for seq, prow in enumerate(project_rows, start=1):
        src_vals = prow["values"]
        green = prow["green_cols"]
        out_vals: list = [seq]
        out_green: set[int] = set()  # 输出列 0-based（含序号）
        out_col = 1  # 从序号后的第一列开始映射时用
        for i, h in enumerate(headers):
            if seq_src_i is not None and i == seq_src_i:
                continue
            v = src_vals[i] if i < len(src_vals) else None
            try:
                if pd.isna(v):
                    v = None
            except Exception:
                pass
            out_vals.append(v)
            if i in green:
                # 输出列索引：0=序号，其后按顺序
                out_green.add(len(out_vals) - 1)
        excel_row = seq + 1
        for col_idx, v in enumerate(out_vals, start=1):
            cell = ws.cell(excel_row, col_idx, v)
            h = out_headers[col_idx - 1]
            if col_idx - 1 in out_green:
                cell.fill = _FX_GREEN_FILL
                cell.font = _FX_GREEN_FONT
            if h in ("含税金额", "不含税金额", "7%增值税", "单价", "计算数量") and isinstance(
                v, (int, float)
            ):
                cell.number_format = "#,##0.00"
            elif h == "汇率" and isinstance(v, (int, float)):
                cell.number_format = "0.0000"

    amount_keys = ("金额", "单价", "税", "汇率", "数量")
    amount_headers = [h for h in out_headers if any(k in h for k in amount_keys)]
    apply_common_excel_layout(
        ws,
        numeric_start_col=ws.max_column + 1,
        amount_headers=amount_headers,
        header_row=1,
        header_height=36,
        freeze_panes="A2",
    )

    # 布局可能重置部分样式；再次补回绿色（保证填充保留）
    green_header_idx = {
        i + 1 for i, h in enumerate(out_headers) if h in ("汇率", "含税金额")
    }
    # 按写入时记录的 green 再刷一遍更准确
    for seq, prow in enumerate(project_rows, start=1):
        excel_row = seq + 1
        # 重建 out_green
        green = prow["green_cols"]
        seq_src_i = headers.index("序号") if "序号" in headers else None
        out_col = 2  # 1=序号
        for i, h in enumerate(headers):
            if seq_src_i is not None and i == seq_src_i:
                continue
            if i in green:
                cell = ws.cell(excel_row, out_col)
                cell.fill = _FX_GREEN_FILL
                cell.font = _FX_GREEN_FONT
            out_col += 1

    left_keys = (
        "物品名称",
        "规格",
        "原始物品名称",
        "原始规格",
        "项目名称",
        "材料归类",
        "供货商",
        "备注",
    )
    left = Alignment(horizontal="left", vertical="center")
    center = Alignment(horizontal="center", vertical="center")
    for col_idx, h in enumerate(out_headers, start=1):
        for r in range(2, (ws.max_row or 1) + 1):
            cell = ws.cell(r, col_idx)
            if h == "序号":
                cell.alignment = center
            elif any(k in h for k in left_keys):
                cell.alignment = left

    if ws.max_row and ws.max_column:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    wb.save(out_path)


def main(argv: list[str] | None = None) -> int:
    default_std = _default_std_sum_path()
    parser = argparse.ArgumentParser(
        description=(
            "按项目拆分 std_sum → E_std_split/{编码}_{名称}_采购数据.xlsx"
            "（保留汇率换算与绿色填充）"
        )
    )
    parser.add_argument(
        "--input",
        default="",
        help=f"输入 Excel，默认 {default_std}",
    )
    parser.add_argument(
        "--output-dir",
        default="",
        help=f"输出目录，默认 {STD_SPLIT_DIR}",
    )
    parser.add_argument(
        "--from-raw-sum",
        action="store_true",
        help=f"兼容：改从 C_raw_sum/{RAW_SUM_FILE_NAME} 拆分（无汇率标绿）",
    )
    parser.add_argument(
        "--clear",
        action="store_true",
        help="写出前清空输出目录中已有 *_采购数据.xlsx",
    )
    args = parser.parse_args(argv)

    if args.input:
        src = Path(args.input).expanduser()
    elif args.from_raw_sum:
        src = RAW_SUM_DIR / RAW_SUM_FILE_NAME
    else:
        src = default_std

    out_dir = Path(args.output_dir).expanduser() if args.output_dir else STD_SPLIT_DIR
    if not src.is_file():
        raise FileNotFoundError(
            f"未找到汇总文件: {src}\n"
            f"请先运行 流水线/04b_std_sum汇总.py 生成 std_sum"
            + (f"\n或使用 --from-raw-sum 从 {RAW_SUM_DIR / RAW_SUM_FILE_NAME} 拆分" if not args.from_raw_sum else "")
        )
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.clear:
        for p in out_dir.glob("*_采购数据.xlsx"):
            try:
                p.unlink()
            except Exception:
                pass

    print("按项目拆分 std_sum（保留汇率/含税填入与绿色填充）")
    print(f"- 输入: {src}")
    print(f"- 输出: {out_dir}")

    headers, rows = _load_merged_sheet_rows(src)
    if not rows:
        print("无数据")
        return 1

    n_green_rows = sum(1 for r in rows if r["green_cols"])
    if n_green_rows:
        print(f"- 源表绿色高亮行: {n_green_rows}（将随拆分保留）")

    groups: dict[str, list[dict]] = defaultdict(list)
    group_meta: dict[str, tuple[str, str]] = {}
    for r in rows:
        key = r["key"]
        groups[key].append(r)
        if key not in group_meta:
            group_meta[key] = (r["code"], r["name"])

    n_files = 0
    n_green_kept = 0
    for key, prow_list in groups.items():
        code, name = group_meta[key]
        if not code and not name:
            code, name = "未知", str(key)
        if not code:
            code = "无编码"
        if not name:
            name = "未命名项目"

        fname = (
            f"{_safe_filename_part(code)}_"
            f"{_safe_filename_part(name)}_采购数据.xlsx"
        )
        out_path = out_dir / fname
        _write_project_excel(headers, prow_list, out_path)
        g_cnt = sum(1 for r in prow_list if r["green_cols"])
        n_green_kept += g_cnt
        n_files += 1
        extra = f"，其中汇率换算标绿 {g_cnt} 行" if g_cnt else ""
        print(f"  ✓ {fname}  ({len(prow_list)} 行{extra})")

    print(f"\n完成: 共 {n_files} 个项目文件 → {out_dir}")
    if n_green_kept:
        print(f"  绿色高亮合计保留 {n_green_kept} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
