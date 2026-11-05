# -*- coding: utf-8 -*-
"""把「最新银行账」和「流水」按项目、按日期左右对照。

同一项目、同一天里，实付金额相同的行先配到同一行，多出来的行插入空行并把后面的记录下移。
对不上金额的剩余行仍按当天原顺序补齐。实付差值 = 银行账实付金额 − 流水实付金额；缺一边时缺的金额按 0。
同一天里差额不为 0 的实付差值作为一组：这组求和为 0 时标绿色，否则标红色。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BANK_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/B_财务付款_std/最新银行账"
)
FLOW_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/B_财务付款_std/流水"
)
OUT_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/流水与银行账按日对照"
)
SUMMARY_NAME = "00_项目汇总.xlsx"

BANK_FIELDS = [
    ("项目编码", "银行账_项目编码"),
    ("摘要", "银行账_摘要"),
    ("实付金额", "银行账_实付金额"),
    ("日期", "银行账_日期"),
]
FLOW_FIELDS = [
    ("摘要", "流水_摘要"),
    ("应付金额", "流水_应付金额"),
    ("预扣税额", "流水_预扣税额"),
    ("实付金额", "流水_实付金额"),
    ("增值税额", "流水_增值税额"),
    ("不含税额", "流水_不含税额"),
    ("付款金额", "流水_付款金额"),
    ("日期", "流水_日期"),
]
AMOUNT_HEADERS = {
    "银行账_实付金额",
    "实付差值",
    "流水_应付金额",
    "流水_预扣税额",
    "流水_实付金额",
    "流水_增值税额",
    "流水_不含税额",
    "流水_付款金额",
}
DATE_HEADERS = {"银行账_日期", "流水_日期", "对照日期"}


def project_key(path: Path, folder_name: str) -> str:
    stem = path.stem
    suffix = f"_{folder_name}"
    if stem.endswith(suffix):
        return stem[: -len(suffix)]
    return stem


def list_books(folder: Path) -> dict[str, Path]:
    books: dict[str, Path] = {}
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$") or path.name.startswith("."):
            continue
        books[project_key(path, folder.name)] = path
    return books


def read_detail(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    sheet = "明细数据" if "明细数据" in xl.sheet_names else xl.sheet_names[0]
    frame = pd.read_excel(path, sheet_name=sheet)
    if "日期" in frame.columns:
        frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce")
    return frame


def paid_amount(row: pd.Series | None) -> float:
    if row is None:
        return 0.0
    for column in ("实付金额", "付款金额"):
        if column not in row.index:
            continue
        value = row[column]
        if pd.notna(value):
            return float(value)
    return 0.0


def day_key(value) -> str:
    if value is None or pd.isna(value):
        return ""
    stamp = pd.Timestamp(value)
    if pd.isna(stamp):
        return ""
    return stamp.strftime("%Y-%m-%d")


def side_values(row: pd.Series | None, fields: list[tuple[str, str]]) -> dict:
    out = {header: None for _source, header in fields}
    if row is None:
        return out
    for source, header in fields:
        if source not in row.index:
            continue
        value = row[source]
        if pd.isna(value):
            continue
        out[header] = value
    return out


def align_day(left_rows: list, right_rows: list) -> list[tuple]:
    """同一天内按实付金额对齐。多出的行插空，使后面的相同金额仍落在同一行。"""
    right_used = [False] * len(right_rows)
    assigned: list[int | None] = [None] * len(left_rows)

    def amount(row) -> float:
        return round(paid_amount(row), 2)

    for i, left in enumerate(left_rows):
        target = amount(left)
        for j, right in enumerate(right_rows):
            if right_used[j]:
                continue
            if abs(amount(right) - target) <= 0.01:
                right_used[j] = True
                assigned[i] = j
                break

    leftover_left = [i for i, j in enumerate(assigned) if j is None]
    leftover_right = [j for j, used in enumerate(right_used) if not used]
    for i, j in zip(leftover_left, leftover_right):
        assigned[i] = j
        right_used[j] = True

    output: list[tuple] = []
    placed_right: dict[int, int] = {}
    for i, left in enumerate(left_rows):
        j = assigned[i]
        right = right_rows[j] if j is not None else None
        output.append((left, right))
        if j is not None:
            placed_right[j] = len(output) - 1

    for j, right in enumerate(right_rows):
        if j in placed_right:
            continue
        insert_at = 0
        for prev_j, pos in placed_right.items():
            if prev_j < j:
                insert_at = max(insert_at, pos + 1)
        output.insert(insert_at, (None, right))
        for prev_j, pos in list(placed_right.items()):
            if pos >= insert_at:
                placed_right[prev_j] = pos + 1
        placed_right[j] = insert_at
    return output


def pair_project(bank: pd.DataFrame | None, flow: pd.DataFrame | None, key: str) -> list[dict]:
    bank_rows = [] if bank is None else [row for _, row in bank.iterrows()]
    flow_rows = [] if flow is None else [row for _, row in flow.iterrows()]
    bank_rows.sort(key=lambda row: (day_key(row.get("日期")) or "9999-99-99",))
    flow_rows.sort(key=lambda row: (day_key(row.get("日期")) or "9999-99-99",))

    days: list[str] = []
    seen: set[str] = set()
    for row in bank_rows + flow_rows:
        day = day_key(row.get("日期")) or "9999-99-99"
        if day not in seen:
            seen.add(day)
            days.append(day)
    days.sort()

    paired: list[dict] = []
    for day in days:
        left = [row for row in bank_rows if (day_key(row.get("日期")) or "9999-99-99") == day]
        right = [row for row in flow_rows if (day_key(row.get("日期")) or "9999-99-99") == day]
        code, name = split_project_key(key)
        for bank_row, flow_row in align_day(left, right):
            record = {
                "项目编码": code,
                "项目名称": name,
                "对照日期": None if day == "9999-99-99" else day,
            }
            record.update(side_values(bank_row, BANK_FIELDS))
            record.update(side_values(flow_row, FLOW_FIELDS))
            record["实付差值"] = round(paid_amount(bank_row) - paid_amount(flow_row), 2)
            paired.append(record)
    return paired


def split_project_key(key: str) -> tuple[str, str]:
    """项目键是「项目编码_项目名称」，名称里可以再带下划线或连字符。"""
    code, sep, name = key.partition("_")
    if not sep:
        return "", key
    return code, name


def detail_columns() -> list[str]:
    return (
        ["项目编码", "项目名称", "对照日期"]
        + [header for _source, header in BANK_FIELDS]
        + ["实付差值"]
        + [header for _source, header in FLOW_FIELDS]
    )


def safe_workbook_name(key: str) -> str:
    text = key.strip() or "未命名项目"
    for char in '\\/:*?"<>|':
        text = text.replace(char, "_")
    return text[:150] + ".xlsx"


def one_project_summary(key: str, detail: pd.DataFrame) -> dict:
    bank_paid = pd.to_numeric(detail["银行账_实付金额"], errors="coerce").sum(min_count=1)
    flow_paid = pd.to_numeric(detail["流水_实付金额"], errors="coerce").sum(min_count=1)
    diff = pd.to_numeric(detail["实付差值"], errors="coerce").sum()
    code, name = split_project_key(key)
    return {
        "项目编码": code,
        "项目名称": name,
        "文件名": safe_workbook_name(key),
        "银行账行数": int((detail["银行账_摘要"].notna() | detail["银行账_实付金额"].notna()).sum()),
        "流水行数": int((detail["流水_摘要"].notna() | detail["流水_实付金额"].notna()).sum()),
        "银行账实付合计": None if pd.isna(bank_paid) else round(float(bank_paid), 2),
        "流水实付合计": None if pd.isna(flow_paid) else round(float(flow_paid), 2),
        "实付差值合计": round(float(diff), 2),
        "差额行数": int((detail["实付差值"].abs() > 0.01).sum()),
    }


def write_project_book(detail: pd.DataFrame, path: Path) -> None:
    diff_only = detail[detail["实付差值"].abs() > 0.01].copy()
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="按日对照", index=False)
        diff_only.to_excel(writer, sheet_name="仅差额", index=False)
        for sheet_name in ("按日对照", "仅差额"):
            style_sheet(writer.book[sheet_name])


def write_summary_book(rows: list[dict], path: Path) -> None:
    frame = pd.DataFrame(rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="项目汇总", index=False)
        style_sheet(writer.book["项目汇总"], highlight_diff=False)


def style_sheet(ws, highlight_diff: bool = True) -> None:
    header_fill = {
        "银行": PatternFill("solid", fgColor="D6EAF8"),
        "流水": PatternFill("solid", fgColor="D5F5E3"),
        "mid": PatternFill("solid", fgColor="FDEBD0"),
    }
    diff_fill = PatternFill("solid", fgColor="F5B7B1")
    offset_fill = PatternFill("solid", fgColor="A9DFBF")
    header_font = Font(bold=True)
    center = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30
    headers = [cell.value for cell in ws[1]]
    diff_col = headers.index("实付差值") + 1 if "实付差值" in headers else None
    for cell in ws[1]:
        text = str(cell.value or "")
        if text.startswith("银行账"):
            cell.fill = header_fill["银行"]
        elif text.startswith("流水"):
            cell.fill = header_fill["流水"]
        else:
            cell.fill = header_fill["mid"]
        cell.font = header_font
        cell.alignment = center
    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        if header in AMOUNT_HEADERS:
            ws.column_dimensions[letter].width = 16
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col_idx).number_format = '#,##0.00'
        elif header in DATE_HEADERS:
            ws.column_dimensions[letter].width = 14
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col_idx).number_format = "YYYY-MM-DD"
        elif header and "摘要" in str(header):
            ws.column_dimensions[letter].width = 42
        elif header == "项目名称":
            ws.column_dimensions[letter].width = 28
        else:
            ws.column_dimensions[letter].width = 18
    if highlight_diff and diff_col:
        date_col = headers.index("对照日期") + 1 if "对照日期" in headers else None
        groups: dict[object, list[tuple[int, float]]] = {}
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row, diff_col)
            try:
                value = float(cell.value)
            except (TypeError, ValueError):
                continue
            if abs(value) <= 0.01:
                continue
            day = ws.cell(row, date_col).value if date_col else row
            groups.setdefault(day, []).append((row, value))
        for items in groups.values():
            total = round(sum(value for _row, value in items), 2)
            fill = offset_fill if abs(total) <= 0.01 else diff_fill
            for row, _value in items:
                ws.cell(row, diff_col).fill = fill


def prepare_output_dir() -> None:
    legacy = OUT_DIR.with_suffix(".xlsx")
    if legacy.is_file():
        legacy.unlink()
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def main() -> int:
    if not BANK_DIR.is_dir() or not FLOW_DIR.is_dir():
        print("找不到银行账或流水目录", file=sys.stderr)
        return 1
    print(f"银行账: {BANK_DIR}")
    print(f"流水:   {FLOW_DIR}")
    bank_books = list_books(BANK_DIR)
    flow_books = list_books(FLOW_DIR)
    keys = sorted(set(bank_books) | set(flow_books))
    prepare_output_dir()
    columns = detail_columns()
    summaries: list[dict] = []
    written: set[str] = {SUMMARY_NAME}
    total_rows = 0
    total_diff = 0
    for key in keys:
        bank = read_detail(bank_books[key]) if key in bank_books else None
        flow = read_detail(flow_books[key]) if key in flow_books else None
        detail = pd.DataFrame(pair_project(bank, flow, key), columns=columns)
        name = safe_workbook_name(key)
        write_project_book(detail, OUT_DIR / name)
        summaries.append(one_project_summary(key, detail))
        written.add(name)
        total_rows += len(detail)
        total_diff += int((detail["实付差值"].abs() > 0.01).sum()) if not detail.empty else 0
        print(f"  {name}: 对照 {len(detail)} 行")
    write_summary_book(summaries, OUT_DIR / SUMMARY_NAME)
    for old in OUT_DIR.glob("*.xlsx"):
        if old.name not in written and not old.name.startswith("~$"):
            old.unlink()
    print(f"写出 {len(keys)} 个项目文件 → {OUT_DIR}")
    print(f"对照 {total_rows} 行，其中差额 {total_diff} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
