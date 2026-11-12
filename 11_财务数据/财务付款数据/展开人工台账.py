# -*- coding: utf-8 -*-
"""把人工费台账按工种展开，写入财务付款模板。

一行台账最多写成四行：

  中方正式工    应付金额 = 中方正式工元，应付币种 RMB
                实付金额、付款金额 = 中方正式工铢，实付币种 泰铢
  中方劳务      应付金额 = 中方劳务工元，应付币种 RMB
                实付金额、付款金额 = 中方劳务工铢，实付币种 泰铢
  子公司属地工  应付金额、实付金额、付款金额 = 子公司属地工铢，币种 泰铢
  劳务公司      应付金额、实付金额、付款金额 = 劳务公司铢，币种 泰铢

金额为空或 0 的工种不写。币别留空：币别若写成 RMB，汇总时会把没有单独币种的金额再乘汇率。
预扣税额、增值税额、不含税额、供应商名称留空。日期用源表的年月日。

换文件后重跑，二选一：

  1. 改下面的 SOURCE、TARGET，然后
     python3 展开人工台账.py
  2. 不改代码，直接传路径
     python3 展开人工台账.py 源.xlsx 目标.xlsx
"""

from __future__ import annotations

import sys
from calendar import monthrange
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils.datetime import from_excel

SOURCE = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/工资分开/"
    "泰国子公司项目人工费台账-2026年9月汇总.xlsx"
)
TARGET = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/工资分开/人工台账展开.xlsx"
)
SOURCE_SHEET = "基础数据"
TARGET_SHEET = "财务数据"

# (摘要里的工种, 应付金额列, 泰铢列, 应付是否为人民币)
# 泰铢工种没有「元」列，应付金额列写 None，应付、实付、付款都用泰铢列。
CATEGORIES = (
    ("中方正式工", "中方正式工元", "中方正式工铢", True),
    ("中方劳务", "中方劳务工元", "中方劳务工铢", True),
    ("子公司属地工", None, "子公司属地工铢", False),
    ("劳务公司", None, "劳务公司铢", False),
)

RMB = "RMB"
THB = "泰铢"

# 目标表要写入的列。没有出现在某一行字典里的列保持空白。
WRITE_COLUMNS = (
    "项目编码",
    "项目名称",
    "摘要",
    "应付金额",
    "实付金额",
    "应付币种",
    "实付币种",
    "付款金额",
    "汇率",
    "汇率缺失",
    "日期",
)


def _num(value):
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    return float(text)


def _empty_amount(value) -> bool:
    number = value if isinstance(value, float) else _num(value)
    return number is None or abs(number) <= 1e-9


def _month_end(year, month) -> datetime:
    year_n = int(year)
    month_n = int(month)
    return datetime(year_n, month_n, monthrange(year_n, month_n)[1])


def _as_date(value, year, month) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, str) and value.strip():
        text = value.strip().replace("/", "-").replace(".", "-")
        return datetime.strptime(text[:10], "%Y-%m-%d")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value)
        if isinstance(converted, datetime):
            return converted
        return datetime(converted.year, converted.month, converted.day)
    if year not in (None, "") and month not in (None, ""):
        return _month_end(year, month)
    raise ValueError(f"缺少年月日：年={year!r} 月={month!r} 值={value!r}")


def _year_month(value, year, month) -> str:
    if value not in (None, ""):
        return str(value).strip()
    return f"{int(year):04d}-{int(month):02d}"


def _cell(row, index, key):
    return row[index[key]]


def load_source(path: Path, sheet_name: str):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise SystemExit(f"源表没有工作表「{sheet_name}」，现有：{wb.sheetnames}")
        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        header = next(rows)
        index = {}
        for pos, name in enumerate(header):
            if name is not None and str(name).strip():
                index[str(name).strip()] = pos
        needed = ["项目编码", "项目名称", "年", "月", "年月", "年月日", "汇率"]
        for _label, payable_col, thb_col, _is_rmb in CATEGORIES:
            if payable_col:
                needed.append(payable_col)
            needed.append(thb_col)
        missing = [name for name in needed if name not in index]
        if missing:
            raise SystemExit(f"源表缺少列：{missing}")
        records = []
        for row in rows:
            if row is None:
                continue
            code = row[index["项目编码"]] if index["项目编码"] < len(row) else None
            name = row[index["项目名称"]] if index["项目名称"] < len(row) else None
            code_blank = code is None or str(code).strip() == ""
            name_blank = name is None or str(name).strip() == ""
            if code_blank and name_blank:
                continue
            records.append(row)
        return index, records
    finally:
        wb.close()


def expand_rows(index, records):
    """返回 (写出的行, 元与铢只有一边有数的提示)。"""
    written = []
    warnings = []
    for row in records:
        code = _cell(row, index, "项目编码")
        name = _cell(row, index, "项目名称")
        year = _cell(row, index, "年")
        month = _cell(row, index, "月")
        year_month = _year_month(_cell(row, index, "年月"), year, month)
        when = _as_date(_cell(row, index, "年月日"), year, month)
        rate = _num(_cell(row, index, "汇率"))
        for label, payable_col, thb_col, is_rmb in CATEGORIES:
            thb = _num(_cell(row, index, thb_col))
            payable = _num(_cell(row, index, payable_col)) if payable_col else thb
            if _empty_amount(payable) and _empty_amount(thb):
                continue
            if is_rmb and _empty_amount(payable) != _empty_amount(thb):
                warnings.append(
                    f"{name} {year_month} {label} 元与铢只有一边有数："
                    f"元={payable} 铢={thb}"
                )
            item = {
                "项目编码": code,
                "项目名称": name,
                "摘要": f"{label} {year_month}",
                "应付金额": payable,
                "实付金额": thb,
                "付款金额": thb,
                "汇率缺失": 0,
                "日期": when,
            }
            if is_rmb:
                item["应付币种"] = RMB
                item["实付币种"] = THB
                item["汇率"] = rate
            else:
                item["应付币种"] = THB
                item["实付币种"] = THB
            written.append(item)
    return written, warnings


def write_target(path: Path, sheet_name: str, rows) -> None:
    wb = openpyxl.load_workbook(path)
    try:
        if sheet_name not in wb.sheetnames:
            raise SystemExit(f"目标表没有工作表「{sheet_name}」，现有：{wb.sheetnames}")
        ws = wb[sheet_name]
        header_cells = next(ws.iter_rows(min_row=1, max_row=1, max_col=ws.max_column))
        columns = {}
        for cell in header_cells:
            if cell.value is not None and str(cell.value).strip():
                columns[str(cell.value).strip()] = cell.column
        missing = [name for name in WRITE_COLUMNS if name not in columns]
        if missing:
            raise SystemExit(f"目标表缺少列：{missing}")

        scan = [columns[name] for name in ("项目编码", "项目名称", "付款金额") if name in columns]
        last_data = 1
        for row_idx in range(2, ws.max_row + 1):
            if any(ws.cell(row_idx, col).value not in (None, "") for col in scan):
                last_data = row_idx
        clear_until = max(last_data, 1 + len(rows))
        clear_cols = [columns[name] for name in WRITE_COLUMNS]
        for row_idx in range(2, clear_until + 1):
            for col in clear_cols:
                ws.cell(row_idx, col).value = None
        for offset, item in enumerate(rows):
            row_idx = 2 + offset
            for key, value in item.items():
                if value is not None:
                    ws.cell(row_idx, columns[key]).value = value
        wb.save(path)
    finally:
        wb.close()


def _print_summary(rows, warnings) -> None:
    print(f"写出 {len(rows)} 行")
    totals = {}
    for item in rows:
        label = item["摘要"].split()[0]
        count, payable, paid = totals.get(label, (0, 0.0, 0.0))
        totals[label] = (
            count + 1,
            payable + float(item["应付金额"] or 0),
            paid + float(item["付款金额"] or 0),
        )
    for label, (count, payable, paid) in totals.items():
        print(f"  {label}: {count} 行，应付 {payable:,.2f}，付款 {paid:,.2f}")
    if warnings:
        print(f"需看一眼的行 {len(warnings)}：")
        for line in warnings[:20]:
            print(f"  {line}")


def main(argv: list[str]) -> None:
    source = SOURCE
    target = TARGET
    if len(argv) == 3:
        source = Path(argv[1]).expanduser()
        target = Path(argv[2]).expanduser()
    elif len(argv) != 1:
        raise SystemExit(
            "用法：python3 展开人工台账.py\n"
            "或：  python3 展开人工台账.py 源.xlsx 目标.xlsx"
        )
    if not source.is_file():
        raise SystemExit(f"找不到源文件：{source}")
    if not target.is_file():
        raise SystemExit(f"找不到目标文件：{target}")
    index, records = load_source(source, SOURCE_SHEET)
    rows, warnings = expand_rows(index, records)
    write_target(target, TARGET_SHEET, rows)
    print(f"源：{source}")
    print(f"目标：{target}")
    print(f"源数据 {len(records)} 行")
    _print_summary(rows, warnings)


if __name__ == "__main__":
    main(sys.argv)
