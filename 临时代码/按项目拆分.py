# -*- coding: utf-8 -*-
"""按「项目编码_项目名称」拆分 Excel，并保留原表格式。

运行后输入要拆分的文件路径。结果写在原文件同一个文件夹里，
文件名为「项目编码_项目名称.xlsx」。原文件不会被修改。

也可以直接带上路径：

    python3 按项目拆分.py /路径/财务付款.xlsx

同一编码、两个名称会分成两个文件，名称按原文字保留。
没有项目编码的，文件名以「_项目名称」开头。
金额为空的整行会跳过；有内容的行全部写出。
"""

from __future__ import annotations

import sys
from collections import defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_SCAN_ROWS = 30
CODE_HEADER = "项目编码"
NAME_HEADER = "项目名称"
_ILLEGAL = '\\/:*?"<>|'


def prompt_path() -> Path | None:
    print("按「项目编码_项目名称」拆分 Excel，原表的字体、底色、数字和日期格式会保留。")
    print("结果保存在原文件同一个文件夹。直接回车退出。")
    while True:
        try:
            raw = input("请输入 Excel 路径：").strip()
        except EOFError:
            print()
            return None
        if not raw:
            return None
        if (raw.startswith('"') and raw.endswith('"')) or (
            raw.startswith("'") and raw.endswith("'")
        ):
            raw = raw[1:-1]
        path = Path(raw).expanduser()
        if path.is_file():
            return path
        print(f"找不到文件：{path}")


def find_header(ws):
    """返回 (表头行, 项目编码列, 项目名称列)。前若干行里没有这两列则返回 None。"""
    last = min(HEADER_SCAN_ROWS, ws.max_row or 1)
    max_col = ws.max_column or 1
    for row_idx in range(1, last + 1):
        found = {}
        for col_idx in range(1, max_col + 1):
            value = ws.cell(row_idx, col_idx).value
            if isinstance(value, str):
                label = value.strip()
                if label in (CODE_HEADER, NAME_HEADER) and label not in found:
                    found[label] = col_idx
        if CODE_HEADER in found and NAME_HEADER in found:
            return row_idx, found[CODE_HEADER], found[NAME_HEADER]
    return None


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value).strip()


def row_has_value(ws, row_idx, max_col) -> bool:
    for col_idx in range(1, max_col + 1):
        value = ws.cell(row_idx, col_idx).value
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return True
    return False


def file_stem(code: str, name: str) -> str:
    stem = f"{code}_{name}"
    cleaned = "".join("_" if ch in _ILLEGAL or ord(ch) < 32 else ch for ch in stem)
    cleaned = cleaned.rstrip(" .")
    if cleaned in ("", "_"):
        cleaned = "未填写项目"
    if len(cleaned) > 120:
        cleaned = cleaned[:120].rstrip(" .")
    return cleaned


def apply_default_font(src_book, dst_book) -> None:
    """原表没单独设字体的格子，沿用工作簿默认字体。新工作簿默认是 Calibri。"""
    if not src_book._fonts:
        return
    src_font = src_book._fonts[0]
    font = Font(
        name=src_font.name,
        sz=src_font.size,
        family=src_font.family,
        charset=src_font.charset,
        scheme=src_font.scheme,
        bold=src_font.bold,
        italic=src_font.italic,
        color=copy(src_font.color) if src_font.color is not None else None,
    )
    dst_book._fonts[0] = font
    for named in dst_book._named_styles:
        if named.name in ("Normal", "常规"):
            named.font = font


def copy_cell(src, dst) -> None:
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format
    if getattr(src, "quotePrefix", False):
        dst.quotePrefix = True
    if src.hyperlink is not None:
        dst.hyperlink = copy(src.hyperlink)
    if src.comment is not None:
        dst.comment = copy(src.comment)


def copy_row(src_ws, dst_ws, src_row, dst_row, max_col) -> None:
    for col_idx in range(1, max_col + 1):
        copy_cell(src_ws.cell(src_row, col_idx), dst_ws.cell(dst_row, col_idx))
    src_dim = src_ws.row_dimensions[src_row]
    dst_dim = dst_ws.row_dimensions[dst_row]
    if src_dim.height is not None:
        dst_dim.height = src_dim.height
    if src_dim.hidden:
        dst_dim.hidden = True


def copy_columns(src_ws, dst_ws) -> None:
    for letter, dim in src_ws.column_dimensions.items():
        if not letter or not str(letter)[:1].isalpha():
            continue
        target = dst_ws.column_dimensions[letter]
        target.width = dim.width
        target.hidden = dim.hidden
        target.bestFit = dim.bestFit
        target.outlineLevel = dim.outlineLevel


def copy_sheet_view(src_ws, dst_ws, header_row) -> None:
    view = src_ws.sheet_view
    dst_ws.sheet_view.zoomScale = view.zoomScale
    dst_ws.sheet_view.zoomScaleNormal = view.zoomScaleNormal
    dst_ws.sheet_view.view = view.view
    dst_ws.sheet_view.showGridLines = view.showGridLines
    dst_ws.sheet_view.showRowColHeaders = view.showRowColHeaders
    freeze = src_ws.freeze_panes
    if not freeze:
        return
    letters, row_text = "", ""
    for ch in str(freeze):
        if ch.isalpha():
            letters += ch
        else:
            row_text += ch
    try:
        freeze_row = int(row_text)
    except ValueError:
        return
    if freeze_row <= header_row + 1:
        dst_ws.freeze_panes = freeze
    else:
        dst_ws.freeze_panes = f"{letters}{header_row + 1}"


def copy_page_setup(src_ws, dst_ws) -> None:
    setup = src_ws.page_setup
    dst_setup = dst_ws.page_setup
    dst_setup.orientation = setup.orientation
    dst_setup.paperSize = setup.paperSize
    dst_setup.scale = setup.scale
    dst_setup.fitToWidth = setup.fitToWidth
    dst_setup.fitToHeight = setup.fitToHeight
    dst_setup.fitToPage = setup.fitToPage
    dst_setup.pageOrder = setup.pageOrder
    dst_setup.horizontalDpi = setup.horizontalDpi
    dst_setup.verticalDpi = setup.verticalDpi
    margins = src_ws.page_margins
    dst_ws.page_margins.left = margins.left
    dst_ws.page_margins.right = margins.right
    dst_ws.page_margins.top = margins.top
    dst_ws.page_margins.bottom = margins.bottom
    dst_ws.page_margins.header = margins.header
    dst_ws.page_margins.footer = margins.footer
    dst_ws.print_options.horizontalCentered = src_ws.print_options.horizontalCentered
    dst_ws.print_options.verticalCentered = src_ws.print_options.verticalCentered
    dst_ws.sheet_format.defaultRowHeight = src_ws.sheet_format.defaultRowHeight
    dst_ws.sheet_format.defaultColWidth = src_ws.sheet_format.defaultColWidth
    dst_ws.sheet_format.customHeight = src_ws.sheet_format.customHeight
    dst_ws.sheet_properties.tabColor = src_ws.sheet_properties.tabColor
    dst_ws.sheet_properties.filterMode = src_ws.sheet_properties.filterMode
    dst_ws.oddHeader.left.text = src_ws.oddHeader.left.text
    dst_ws.oddHeader.center.text = src_ws.oddHeader.center.text
    dst_ws.oddHeader.right.text = src_ws.oddHeader.right.text
    dst_ws.oddFooter.left.text = src_ws.oddFooter.left.text
    dst_ws.oddFooter.center.text = src_ws.oddFooter.center.text
    dst_ws.oddFooter.right.text = src_ws.oddFooter.right.text
    dst_ws.print_title_rows = src_ws.print_title_rows
    dst_ws.print_title_cols = src_ws.print_title_cols
    dst_ws.page_setup.fitToHeight = setup.fitToHeight
    dst_ws.sheet_properties.pageSetUpPr.fitToPage = (
        src_ws.sheet_properties.pageSetUpPr.fitToPage
    )


def apply_filter(src_ws, dst_ws, header_row, last_row) -> None:
    ref = src_ws.auto_filter.ref if src_ws.auto_filter is not None else None
    if not ref:
        return
    min_col, min_row, max_col, _max_row = range_boundaries(ref)
    start_row = min_row if min_row <= header_row else header_row
    end_row = max(last_row, start_row)
    dst_ws.auto_filter.ref = (
        f"{get_column_letter(min_col)}{start_row}:{get_column_letter(max_col)}{end_row}"
    )


def _remap_sqref(sqref, row_map, header_row, new_last) -> str:
    pieces = []
    for part in str(sqref).split():
        bounds = CellRange(part)
        if bounds.max_row <= header_row:
            pieces.append(str(bounds))
            continue
        mapped = [row_map[row] for row in range(bounds.min_row, bounds.max_row + 1) if row in row_map]
        covers_data_block = bounds.max_row - bounds.min_row >= 1 and bounds.min_row <= header_row + 1
        if covers_data_block and new_last > header_row:
            updated = CellRange(str(bounds))
            updated.min_row = min(bounds.min_row, header_row + 1)
            if updated.min_row > header_row + 1:
                updated.min_row = header_row + 1
            updated.max_row = new_last
            pieces.append(str(updated))
            continue
        if mapped and mapped[-1] - mapped[0] + 1 == len(mapped):
            updated = CellRange(str(bounds))
            updated.min_row = mapped[0]
            updated.max_row = mapped[-1]
            pieces.append(str(updated))
    return " ".join(pieces)


def copy_merges(src_ws, dst_ws, row_map, header_row, new_last) -> None:
    for merged in src_ws.merged_cells.ranges:
        mapped = _remap_sqref(str(merged), row_map, header_row, new_last)
        if not mapped:
            continue
        for part in mapped.split():
            dst_ws.merge_cells(part)


def copy_validations(src_ws, dst_ws, row_map, header_row, new_last) -> None:
    if not src_ws.data_validations:
        return
    for item in src_ws.data_validations.dataValidation:
        sqref = _remap_sqref(item.sqref, row_map, header_row, new_last)
        if not sqref:
            continue
        validation = DataValidation(
            type=item.type,
            formula1=item.formula1,
            formula2=item.formula2,
            allow_blank=item.allow_blank,
            showErrorMessage=item.showErrorMessage,
            showInputMessage=item.showInputMessage,
            showDropDown=item.showDropDown,
            error=item.error,
            errorTitle=item.errorTitle,
            prompt=item.prompt,
            promptTitle=item.promptTitle,
            operator=item.operator,
            sqref=sqref,
        )
        dst_ws.add_data_validation(validation)


def copy_conditional(src_ws, dst_ws, row_map, header_row, new_last) -> None:
    for sqref, rules in src_ws.conditional_formatting._cf_rules.items():
        mapped = _remap_sqref(str(sqref), row_map, header_row, new_last)
        if not mapped:
            continue
        for rule in rules:
            dst_ws.conditional_formatting.add(mapped, copy(rule))


def prepare_sheet(src_ws, dst_ws, header_row, kept_rows):
    """复制表头以上的行和本项目的数据行。返回新表最后一行。"""
    max_col = src_ws.max_column or 1
    row_map = {}
    new_row = 1
    for src_row in range(1, header_row + 1):
        copy_row(src_ws, dst_ws, src_row, new_row, max_col)
        row_map[src_row] = new_row
        new_row += 1
    for src_row in kept_rows:
        copy_row(src_ws, dst_ws, src_row, new_row, max_col)
        row_map[src_row] = new_row
        new_row += 1
    last_row = new_row - 1
    copy_columns(src_ws, dst_ws)
    copy_sheet_view(src_ws, dst_ws, header_row)
    copy_page_setup(src_ws, dst_ws)
    apply_filter(src_ws, dst_ws, header_row, last_row)
    try:
        copy_merges(src_ws, dst_ws, row_map, header_row, last_row)
        copy_validations(src_ws, dst_ws, row_map, header_row, last_row)
        copy_conditional(src_ws, dst_ws, row_map, header_row, last_row)
    except Exception as exc:
        print(f"  「{src_ws.title}」的合并单元格或条件下拉未完整复制：{exc}")
    return last_row


def copy_whole_sheet(src_ws, dst_ws) -> None:
    """附带工作表没有项目列，整表复制到每个项目文件。"""
    max_col = src_ws.max_column or 1
    last_value = 1
    for row_idx in range(1, (src_ws.max_row or 1) + 1):
        if row_has_value(src_ws, row_idx, max_col):
            last_value = row_idx
    prepare_sheet(src_ws, dst_ws, last_value, [])


def collect_groups(ws):
    header = find_header(ws)
    if header is None:
        return None
    header_row, code_col, name_col = header
    max_col = ws.max_column or 1
    groups = defaultdict(list)
    order = []
    for row_idx in range(header_row + 1, (ws.max_row or header_row) + 1):
        if not row_has_value(ws, row_idx, max_col):
            continue
        key = (
            cell_text(ws.cell(row_idx, code_col).value),
            cell_text(ws.cell(row_idx, name_col).value),
        )
        if key not in groups:
            order.append(key)
        groups[key].append(row_idx)
    return {
        "sheet": ws,
        "header_row": header_row,
        "groups": groups,
        "order": order,
    }


def unique_stems(keys):
    used = {}
    stems = {}
    for key in keys:
        base = file_stem(*key)
        stem = base
        seq = 2
        while stem.lower() in used:
            stem = f"{base}_{seq}"
            seq += 1
        used[stem.lower()] = key
        stems[key] = stem
    return stems


def canon(value):
    if value is None or value == "":
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    if isinstance(value, datetime):
        return ("d", value.year, value.month, value.day, value.hour, value.minute, value.second)
    if isinstance(value, date):
        return ("d", value.year, value.month, value.day, 0, 0, 0)
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, (int, float)):
        return ("n", float(value))
    return ("s", str(value))


def verify(source_path: Path, plans, auxiliary, written) -> None:
    from collections import Counter

    source_counts = Counter()
    book = openpyxl.load_workbook(source_path, read_only=True, data_only=False)
    try:
        for plan in plans:
            ws = book[plan["sheet"].title]
            header = find_header(ws)
            if header is None:
                raise SystemExit(f"回读时找不到表头：{ws.title}")
            header_row = header[0]
            for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                if row_idx <= header_row:
                    continue
                values = tuple(canon(value) for value in row)
                if any(value is not None for value in values):
                    source_counts[(ws.title, values)] += 1
    finally:
        book.close()

    back_counts = Counter()
    for _key, path, _rows in written:
        book = openpyxl.load_workbook(path, read_only=True, data_only=False)
        try:
            for plan in plans:
                title = plan["sheet"].title
                if title not in book.sheetnames:
                    continue
                ws = book[title]
                header = find_header(ws)
                if header is None:
                    raise SystemExit(f"{path.name} 的「{title}」缺少项目编码或项目名称")
                for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                    if row_idx <= header[0]:
                        continue
                    values = tuple(canon(value) for value in row)
                    if any(value is not None for value in values):
                        back_counts[(title, values)] += 1
        finally:
            book.close()
    if source_counts != back_counts:
        missing = sum((source_counts - back_counts).values())
        extra = sum((back_counts - source_counts).values())
        raise SystemExit(f"拆分后数据对不上：缺 {missing} 行，多 {extra} 行")
    print(f"已核对 {sum(source_counts.values())} 行，与原表一致。")
    if auxiliary:
        print("附带工作表已复制到每个项目文件：" + "、".join(auxiliary))


def split_workbook(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise SystemExit("这个文件是旧版 .xls，请先在 Excel 里另存为 .xlsx 再拆分。")
    if suffix not in {".xlsx", ".xlsm"}:
        raise SystemExit(f"只拆分 .xlsx 或 .xlsm：{path.name}")

    book = openpyxl.load_workbook(path, data_only=False, keep_vba=suffix == ".xlsm")
    try:
        plans = []
        auxiliary = []
        for ws in book.worksheets:
            plan = collect_groups(ws)
            if plan is None:
                auxiliary.append(ws.title)
            elif plan["order"]:
                plans.append(plan)
            else:
                print(f"「{ws.title}」有项目编码和项目名称，但没有数据行，已跳过。")
        if not plans:
            raise SystemExit("没有找到同时含「项目编码」和「项目名称」的工作表。")

        keys = []
        seen = set()
        for plan in plans:
            for key in plan["order"]:
                if key not in seen:
                    seen.add(key)
                    keys.append(key)
        stems = unique_stems(keys)
        out_dir = path.parent
        written = []
        total = len(keys)
        for index, key in enumerate(keys, 1):
            stem = stems[key]
            out_path = out_dir / f"{stem}.xlsx"
            if out_path.resolve() == path.resolve():
                out_path = out_dir / f"{stem}_拆分.xlsx"
            output = Workbook()
            apply_default_font(book, output)
            default_sheet = output.active
            default_sheet.title = "_待删除"
            created = 0
            for plan in plans:
                kept = plan["groups"].get(key, [])
                if not kept:
                    continue
                sheet = output.create_sheet(plan["sheet"].title)
                prepare_sheet(plan["sheet"], sheet, plan["header_row"], kept)
                created += 1
            for title in auxiliary:
                sheet = output.create_sheet(title)
                copy_whole_sheet(book[title], sheet)
                created += 1
            if created == 0:
                output.close()
                continue
            output.remove(default_sheet)
            try:
                output.save(out_path)
            except PermissionError:
                output.close()
                raise SystemExit(f"无法保存，请先关闭已打开的文件：{out_path}")
            output.close()
            row_count = sum(len(plan["groups"].get(key, [])) for plan in plans)
            written.append((key, out_path, row_count))
            label = f"{key[0]}_{key[1]}" if key[0] else f"(无编码)_{key[1]}"
            print(f"{index}/{total}  {label}.xlsx  {row_count} 行", flush=True)
    finally:
        book.close()

    data_rows = sum(item[2] for item in written)
    print(f"原文件未改：{path}")
    print(f"写出 {len(written)} 个文件、{data_rows} 行 → {out_dir}")
    verify(path, plans, auxiliary, written)


def main(argv: list[str]) -> None:
    if len(argv) > 2:
        raise SystemExit("用法：python3 按项目拆分.py [Excel路径]")
    if len(argv) == 2:
        path = Path(argv[1]).expanduser()
        if not path.is_file():
            raise SystemExit(f"找不到文件：{path}")
    else:
        path = prompt_path()
        if path is None:
            return
    split_workbook(path)


if __name__ == "__main__":
    main(sys.argv)
