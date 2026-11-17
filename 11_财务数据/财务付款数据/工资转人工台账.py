# -*- coding: utf-8 -*-
"""把工资明细收成人工费台账的一行一个月。

源表是「工资部分」。同一项目、同一工资月份（年月调整）合成一行：

  正式工    → 中方正式工铢，中方正式工元 = 铢 ÷ 当月人民币对泰铢
  中方劳务  → 中方劳务工铢，中方劳务工元 = 铢 ÷ 汇率
  属地员工  → 子公司属地工铢
  劳务公司  → 劳务公司铢
  签证费    → 签证费

铢用付款金额合计。金额为 0 的列留空。地区优先用人工费台账里已有的项目地区。

换文件后重跑，二选一：

  1. 改下面的 SOURCE、TARGET，然后
     python3 工资转人工台账.py
  2. 不改代码，直接传路径
     python3 工资转人工台账.py 源.xlsx 目标.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SOURCE = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/3-人工费/"
    "工资8月前开工-8月前（人力部未记录补充数据）.xlsx"
)
TARGET = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/3-人工费/"
    "工资8月前开工-8月前_人工台账.xlsx"
)
SOURCE_SHEET = "工资部分"
FX_PATH = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/汇率三表.xlsx")
REGION_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/11_人工费/"
    "01_人工费台账/泰国子公司项目人工费统计台账2026.10.17.xlsx"
)

# 台账里没有、但名称本身就是地区的项目。
REGION_BY_NAME = {
    "子公司": "子公司",
    "子公司-现场费用": "子公司",
}

COLUMNS = (
    "项目名称",
    "地区",
    "中方正式工元",
    "中方正式工铢",
    "中方劳务工元",
    "中方劳务工铢",
    "子公司属地工铢",
    "劳务公司铢",
    "签证费",
    "年",
    "月",
    "汇率",
    "项目编码",
)
AMOUNT_COLUMNS = (
    "中方正式工元",
    "中方正式工铢",
    "中方劳务工元",
    "中方劳务工铢",
    "子公司属地工铢",
    "劳务公司铢",
    "签证费",
)
# 工资类别 → (泰铢列, 元列或 None)
CATEGORY_COLUMNS = {
    "正式工": ("中方正式工铢", "中方正式工元"),
    "正式员工": ("中方正式工铢", "中方正式工元"),
    "中方劳务": ("中方劳务工铢", "中方劳务工元"),
    "中方劳务工": ("中方劳务工铢", "中方劳务工元"),
    "属地员工": ("子公司属地工铢", None),
    "属地工": ("子公司属地工铢", None),
    "子公司属地工": ("子公司属地工铢", None),
    "劳务公司": ("劳务公司铢", None),
    "签证费": ("签证费", None),
    "签证": ("签证费", None),
}


def _text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _year_month(value):
    text = _text(value)
    if len(text) >= 7 and text[4] == "-" and text[:4].isdigit():
        month_text = text[5:7]
        if month_text.isdigit():
            month = int(month_text)
            if 1 <= month <= 12:
                return int(text[:4]), month
    return None


def _load_rates(path: Path) -> dict:
    table = pd.read_excel(path, sheet_name="汇率表")
    ym_col = cny_col = None
    for col in table.columns:
        name = str(col).replace("\n", "").replace(" ", "")
        if ym_col is None and "年月" in name:
            ym_col = col
        if cny_col is None and "人民币对泰铢" in name:
            cny_col = col
    if ym_col is None or cny_col is None:
        raise ValueError(f"汇率表缺少年月或人民币对泰铢: {list(table.columns)}")
    rates = {}
    for ym, rate in zip(table[ym_col], table[cny_col]):
        parsed = _year_month(ym)
        if parsed is None or pd.isna(rate):
            continue
        rates[f"{parsed[0]:04d}-{parsed[1]:02d}"] = float(rate)
    return rates


def _load_regions(path: Path) -> tuple[dict, dict]:
    frame = pd.read_excel(path, sheet_name="基础数据")
    by_code: dict[str, str] = {}
    by_name: dict[str, str] = {}
    for _, row in frame.iterrows():
        region = _text(row.get("地区"))
        if not region:
            continue
        code = _text(row.get("项目编码"))
        name = _text(row.get("项目名称"))
        if code and code not in by_code:
            by_code[code] = region
        if name and name not in by_name:
            by_name[name] = region
    by_name.update(REGION_BY_NAME)
    return by_code, by_name


def _region(code: str, name: str, by_code: dict, by_name: dict) -> str:
    if code and code in by_code:
        return by_code[code]
    if name and name in by_name:
        return by_name[name]
    return ""


def build_ledger(source: Path) -> pd.DataFrame:
    detail = pd.read_excel(source, sheet_name=SOURCE_SHEET)
    detail["付款金额"] = pd.to_numeric(detail["付款金额"], errors="coerce").fillna(0.0)
    rates = _load_rates(FX_PATH)
    by_code, by_name = _load_regions(REGION_PATH)

    grouped: dict[tuple, dict] = {}
    unknown = {}
    missing_rate = set()
    for record in detail.to_dict("records"):
        parsed = _year_month(record.get("年月调整")) or _year_month(record.get("年月"))
        if parsed is None:
            unknown.setdefault("缺少年月", 0)
            unknown["缺少年月"] += 1
            continue
        year, month = parsed
        category = _text(record.get("工资类别"))
        mapped = CATEGORY_COLUMNS.get(category)
        if mapped is None:
            unknown[category or "空白类别"] = unknown.get(category or "空白类别", 0) + 1
            continue
        code = _text(record.get("项目编码"))
        name = _text(record.get("项目名称"))
        key = (code, name, year, month)
        bucket = grouped.get(key)
        if bucket is None:
            bucket = {col: 0.0 for col in AMOUNT_COLUMNS}
            grouped[key] = bucket
        baht_col, _yuan_col = mapped
        bucket[baht_col] += float(record.get("付款金额") or 0)

    rows = []
    for (code, name, year, month), amounts in grouped.items():
        rate_key = f"{year:04d}-{month:02d}"
        rate = rates.get(rate_key)
        if rate is None:
            missing_rate.add(rate_key)
        row = {
            "项目名称": name,
            "地区": _region(code, name, by_code, by_name),
            "年": year,
            "月": month,
            "汇率": rate,
            "项目编码": code,
        }
        for baht_col, yuan_col in (
            ("中方正式工铢", "中方正式工元"),
            ("中方劳务工铢", "中方劳务工元"),
        ):
            baht = round(amounts[baht_col], 2)
            if baht == 0:
                row[baht_col] = None
                row[yuan_col] = None
                continue
            row[baht_col] = baht
            row[yuan_col] = round(baht / rate, 2) if rate else None
        for col in ("子公司属地工铢", "劳务公司铢", "签证费"):
            baht = round(amounts[col], 2)
            row[col] = baht if baht else None
        rows.append(row)

    frame = pd.DataFrame(rows, columns=COLUMNS)
    frame = frame.sort_values(["年", "月", "项目编码", "项目名称"], kind="mergesort")
    frame = frame.reset_index(drop=True)
    if unknown:
        print("未归入台账的行:")
        for name, count in sorted(unknown.items()):
            print(f"  · {name} {count}")
    if missing_rate:
        print("汇率表没有这些月份，元留空:", ", ".join(sorted(missing_rate)))
    missing_region = (
        frame.loc[frame["地区"].map(_text) == "", ["项目编码", "项目名称"]]
        .drop_duplicates()
        .itertuples(index=False)
    )
    missing_region = [f"{_text(code) or '无编码'}_{name}" for code, name in missing_region]
    if missing_region:
        print("台账里没有地区，已留空:")
        for item in missing_region:
            print(f"  · {item}")
    return frame


def _write(frame: pd.DataFrame, target: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "基础数据"
    header_fill = PatternFill("solid", fgColor="DDEFE3")
    bold = Font(bold=True, color="1F3A2E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B7C7BC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, header in enumerate(COLUMNS, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = header_fill
        cell.font = bold
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = 16 if header in AMOUNT_COLUMNS else 18
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10
    for row_idx, record in enumerate(frame.itertuples(index=False), start=2):
        for col, header in enumerate(COLUMNS, start=1):
            value = record[col - 1]
            if value == "" or (isinstance(value, float) and pd.isna(value)):
                value = None
            cell = ws.cell(row_idx, col, value)
            cell.border = border
            cell.alignment = center
            if header in AMOUNT_COLUMNS and isinstance(value, (int, float)):
                cell.number_format = "#,##0.00"
            elif header == "汇率" and isinstance(value, float):
                cell.number_format = "0.0000"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{ws.max_row}"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)
    wb.close()


def main(argv: list[str]) -> None:
    source = Path(argv[1]) if len(argv) > 1 else SOURCE
    target = Path(argv[2]) if len(argv) > 2 else TARGET
    frame = build_ledger(source)
    _write(frame, target)
    money = frame[list(AMOUNT_COLUMNS)].apply(pd.to_numeric, errors="coerce").fillna(0)
    print(f"已写出 {len(frame)} 行 → {target}")
    print(f"  中方正式工铢 {money['中方正式工铢'].sum():,.2f}")
    print(f"  中方劳务工铢 {money['中方劳务工铢'].sum():,.2f}")
    print(f"  子公司属地工铢 {money['子公司属地工铢'].sum():,.2f}")
    print(f"  劳务公司铢 {money['劳务公司铢'].sum():,.2f}")
    print(f"  签证费 {money['签证费'].sum():,.2f}")


if __name__ == "__main__":
    main(sys.argv)
