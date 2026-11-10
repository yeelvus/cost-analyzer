#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按月份文件夹对比，把后一个文件夹相对前一个文件夹新增的采购行拆成独立 Excel。

例如 2026年09月_各项目采购台账 对比 2026年10月_各项目采购台账：
同一项目只保留 10 月文件里多出来的行。第一个月份文件夹没有上一文件夹，不作为增量。
同一笔采购按物品名称、规格、单位、数量、单价、货币种类、签订日期、供货商识别。
每个有新增的源文件单独另存一份，表头、列顺序和单元格格式与原台账一致，只保留新增行。
"""

from __future__ import annotations

import shutil
import stat
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

_PKG = Path(__file__).resolve().parents[1]
_CODE = _PKG.parent
for _p in (str(_PKG), str(_CODE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from _采购路径 import (  # noqa: E402
    RAW_DIR,
    SUM_DIR,
    extract_month_iso,
    extract_project_name,
    is_excel,
    month_sort_key,
)

OUT_DIR_NAME = "月度增量"
OLD_BOOK_NAME = "采购台账_月度增量.xlsx"

CONTENT_COLS = [
    "材料归类",
    "采购部门",
    "物品名称",
    "规格",
    "单位",
    "数量",
    "单价",
    "不含税金额",
    "7%增值税",
    "含税金额",
    "货币种类",
    "汇率",
    "签订日期",
    "供货商名称",
    "备注",
]
NUM_COLS = {"数量", "单价", "不含税金额", "7%增值税", "含税金额", "汇率"}
# 后补的分类和税额不参与同一笔采购的识别。
KEY_COLS = ["物品名称", "规格", "单位", "数量", "单价", "货币种类", "签订日期", "供货商名称"]
HEADER_ALIASES = {
    "材料归类": ["材料归类"],
    "采购部门": ["采购部门"],
    "物品名称": ["物品名称", "采购物资名称", "物资名称"],
    "规格": ["规格", "规格型号", "型号", "规格型号/参数"],
    "单位": ["标准单位", "单位"],
    "数量": ["数量"],
    "单价": ["单价", "不含税单价", "含税单价"],
    "不含税金额": ["不含税金额", "金额"],
    "7%增值税": ["7%增值税", "税额", "增值税", "税金"],
    "含税金额": ["含税金额", "总金额", "价税合计"],
    "货币种类": ["货币种类", "币种"],
    "汇率": ["汇率"],
    "签订日期": ["签订日期", "日期"],
    "供货商名称": ["供货商名称", "供货单位", "供应商", "供应商名称"],
    "备注": ["备注"],
}
TOTAL_NAMES = {"合计", "总计", "小计", "合 计"}


def _blank(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    return str(value).strip() == ""


def _norm_cell(column: str, value):
    if _blank(value):
        return ""
    if column in NUM_COLS:
        number = pd.to_numeric(value, errors="coerce")
        if pd.isna(number):
            return str(value).replace("\n", " ").strip()
        return round(float(number), 2)
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%Y-%m-%d")
        except (TypeError, ValueError):
            pass
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def _header_map(cells: list) -> dict[str, int] | None:
    labels = ["" if _blank(v) else str(v).replace("\n", "").strip() for v in cells]
    if "序号" not in labels and "编号" not in labels and "No." not in labels:
        return None
    found: dict[str, int] = {}
    for name, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias in labels:
                found[name] = labels.index(alias)
                break
    if "物品名称" not in found or "含税金额" not in found:
        return None
    return found


def read_ledger(path: Path) -> tuple[list[tuple[str, int, list[tuple[int, dict]]]], str]:
    """读出每张已识别工作表的表头行号，以及数据行的 Excel 行号。失败时返回空列表和原因。"""
    try:
        book = pd.ExcelFile(path)
    except Exception as exc:
        return [], f"无法打开: {exc}"

    sheets: list[tuple[str, int, list[tuple[int, dict]]]] = []
    recognized = False
    for sheet in book.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception:
            continue
        if raw.empty:
            continue
        header_at = None
        mapping = None
        for idx in range(min(12, len(raw))):
            mapping = _header_map(raw.iloc[idx].tolist())
            if mapping:
                header_at = idx
                break
        if mapping is None or header_at is None:
            continue
        recognized = True
        body: list[tuple[int, dict]] = []
        for index, record in raw.iloc[header_at + 1 :].iterrows():
            values = record.tolist()
            item = {}
            for name, col in mapping.items():
                item[name] = values[col] if col < len(values) else ""
            name = _norm_cell("物品名称", item.get("物品名称", ""))
            if name in TOTAL_NAMES or "合计" in name:
                continue
            if not any(_norm_cell(col, item.get(col, "")) for col in CONTENT_COLS):
                continue
            body.append((int(index) + 1, {col: item.get(col, "") for col in CONTENT_COLS}))
        sheets.append((sheet, header_at + 1, body))
    if not recognized:
        return [], "未识别到表头"
    return sheets, ""


def _row_key(row: dict) -> tuple:
    return tuple(_norm_cell(col, row.get(col, "")) for col in KEY_COLS)


def _counts(files: list[tuple[Path, list]]) -> Counter:
    counts: Counter = Counter()
    for _path, sheets in files:
        for _sheet, _header, body in sheets:
            for _excel_row, row in body:
                counts[_row_key(row)] += 1
    return counts


def _keep_last_extras(prev_counts: Counter, files: list[tuple[Path, list]]) -> dict[Path, dict[str, set[int]]]:
    """同一键多出来的行，保留文件中靠后的那几行。新增采购通常追加在表尾。"""
    grouped: dict[tuple, list[tuple[Path, str, int]]] = defaultdict(list)
    for path, sheets in files:
        for sheet, _header, body in sheets:
            for excel_row, row in body:
                grouped[_row_key(row)].append((path, sheet, excel_row))
    keep: dict[Path, dict[str, set[int]]] = defaultdict(lambda: defaultdict(set))
    for key, locs in grouped.items():
        extra = len(locs) - prev_counts.get(key, 0)
        if extra > 0:
            for path, sheet, excel_row in locs[-extra:]:
                keep[path][sheet].add(excel_row)
    return keep


def _month_folders(raw_root: Path) -> list[tuple[str, Path]]:
    folders = []
    for path in sorted(raw_root.iterdir()):
        if not path.is_dir() or path.name.startswith("."):
            continue
        month = extract_month_iso(path.name)
        if month:
            folders.append((month, path))
    folders.sort(key=lambda item: month_sort_key(item[0]))
    return folders


def _projects_in(folder: Path) -> dict[str, list[Path]]:
    grouped: dict[str, list[Path]] = defaultdict(list)
    for path in sorted(folder.iterdir()):
        if is_excel(path):
            grouped[extract_project_name(path.name)].append(path)
    return grouped


def _load_folder(folder: Path, month: str, failures: list[dict]) -> dict[str, list[tuple[Path, list]]]:
    loaded: dict[str, list[tuple[Path, list]]] = {}
    for project, paths in _projects_in(folder).items():
        files: list[tuple[Path, list]] = []
        for path in paths:
            sheets, error = read_ledger(path)
            if error:
                failures.append({"台账月份": month, "项目名称": project, "来源文件": path.name, "原因": error})
                continue
            files.append((path, sheets))
        loaded[project] = files
    return loaded


def _unmerge_below(ws, header_row: int) -> None:
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row > header_row:
            ws.unmerge_cells(str(merged))


def _delete_other_rows(ws, header_row: int, keep: set[int]) -> None:
    max_row = ws.max_row or header_row
    ranges: list[tuple[int, int]] = []
    start = None
    for row in range(header_row + 1, max_row + 1):
        if row in keep:
            if start is not None:
                ranges.append((start, row - start))
                start = None
        elif start is None:
            start = row
    if start is not None:
        ranges.append((start, max_row - start + 1))
    for row, amount in reversed(ranges):
        if amount > 0:
            ws.delete_rows(row, amount)


def write_increment(source: Path, dest: Path, sheets: list, keep_by_sheet: dict[str, set[int]]) -> None:
    """复制原文件，只删除表头以下、不属于本次增量的行。"""
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, dest)
    # 源文件若是只读，copy2 会把只读位一并带上，随后无法写回删行结果。
    dest.chmod(dest.stat().st_mode | stat.S_IWUSR)
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        dest.unlink(missing_ok=True)
        raise ValueError(f"只保留 xlsx/xlsm，当前是 {source.suffix}")
    wb = load_workbook(dest)
    try:
        for sheet_name, header_row, _body in sheets:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            keep = keep_by_sheet.get(sheet_name, set())
            last_col = get_column_letter(ws.max_column or 1)
            had_filter = bool(ws.auto_filter.ref)
            _unmerge_below(ws, header_row)
            _delete_other_rows(ws, header_row, keep)
            if had_filter:
                ws.auto_filter.ref = f"A{header_row}:{last_col}{header_row + len(keep)}"
        wb.save(dest)
    except Exception:
        dest.unlink(missing_ok=True)
        raise
    finally:
        wb.close()


def _clear_stale(dest_dir: Path, written: set[Path]) -> None:
    if not dest_dir.is_dir():
        return
    for path in dest_dir.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        if path.resolve() not in written:
            path.unlink()


def build(raw_root: Path, out_root: Path) -> tuple[int, int, list[dict]]:
    folders = _month_folders(raw_root)
    failures: list[dict] = []
    if len(folders) < 2:
        print("月份文件夹不足两个，没有可对比的增量")
        return 0, 0, failures

    cache: dict[Path, dict[str, list]] = {}

    def load(folder: Path, month: str) -> dict[str, list]:
        if folder not in cache:
            cache[folder] = _load_folder(folder, month, failures)
        return cache[folder]

    print(f"起始文件夹不提取: {folders[0][1].name}")
    # 上一文件夹只作为对照，读一次后留给下一个月复用。
    load(folders[0][1], folders[0][0])
    total_files = 0
    total_rows = 0
    for (prev_month, prev_dir), (month, curr_dir) in zip(folders, folders[1:]):
        print(f"{prev_dir.name} → {curr_dir.name}")
        previous = load(prev_dir, prev_month)
        current = load(curr_dir, month)
        dest_dir = out_root / curr_dir.name
        written: set[Path] = set()
        folder_rows = 0
        folder_files = 0
        for project in sorted(current):
            prev_files = previous.get(project, [])
            keep = _keep_last_extras(_counts(prev_files), current[project])
            for path, sheets in current[project]:
                kept = sum(len(rows) for rows in keep.get(path, {}).values())
                if kept == 0:
                    continue
                dest = dest_dir / path.name
                note = "上一文件夹没有该项目" if project not in previous else "相对上一文件夹多出的条目"
                try:
                    write_increment(path, dest, sheets, keep.get(path, {}))
                except Exception as exc:
                    failures.append({"台账月份": month, "项目名称": project, "来源文件": path.name, "原因": f"写出失败: {exc}"})
                    print(f"  {project}: 写出失败 {exc}")
                    continue
                written.add(dest.resolve())
                folder_files += 1
                folder_rows += kept
                print(f"  {project}: 新增 {kept} 行（{note}）")
        _clear_stale(dest_dir, written)
        total_files += folder_files
        total_rows += folder_rows
        print(f"  本文件夹 {folder_files} 个文件，新增 {folder_rows} 行")
        # 再往前的月份不再参与对照，释放读取结果。
        cache.pop(prev_dir, None)
    return total_files, total_rows, failures


def main() -> int:
    if not RAW_DIR.is_dir():
        print(f"找不到采购台账目录: {RAW_DIR}", file=sys.stderr)
        return 1
    out_root = SUM_DIR / OUT_DIR_NAME
    print(f"读取: {RAW_DIR}")
    print(f"写出: {out_root}")
    files, rows, failures = build(RAW_DIR, out_root)
    old_book = SUM_DIR / OLD_BOOK_NAME
    if old_book.exists():
        old_book.unlink()
        print(f"已删除旧的汇总表: {old_book.name}")
    print(f"完成: {files} 个增量文件，{rows} 行")
    if failures:
        print(f"失败 {len(failures)} 个文件:")
        for item in failures:
            print(f"  {item['台账月份']} {item['来源文件']}: {item['原因']}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
