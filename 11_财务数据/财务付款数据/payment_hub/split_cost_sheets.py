# -*- coding: utf-8 -*-
"""一次性把采购、机具、人工费汇总、飞签机票写进「拆分」项目工作簿。

月结和科目宽拆分不调用这里。需要再写时单独运行 attach_split_cost_sheets。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook

from .config import Paths
from .utils import safe_text

TARGET_SUFFIX = "_以发生付款成本.xlsx"
SHEET_NAMES = (
    "采购数据",
    "机具费",
    "飞签机票",
    "机票月度汇总",
    "人工费汇总",
    "人工费按月",
    "人工费科目",
    "人工费科目宽",
)

# 文件名对不上正式编码时，按项目名归到已有的拆分工作簿。
# 两份欣强采购是同一张 4 行表，只保留无编码那份。
_FILE_ALIASES = {
    "无编码_INSEE项目": ("未识别编码_INSEE",),
    "INSEE": ("未识别编码_INSEE",),
    "子公司": ("ZGS_子公司",),
    "无编码_欣强项目": ("未识别编码_欣强电子", "ZGS_欣强电子"),
    "需手动核对_欣强项目": (),
}
# 人工费汇总里没有编码、名称又和文件名不完全一样的行。
_LABOR_NAME_ALIASES = {
    "insee": "未识别编码_INSEE",
    "子公司-现场费用": "未识别编码_子公司-现场费用",
}


def _std_root() -> Path:
    from settings import DATA_ROOT

    return Path(DATA_ROOT)


def _source_dirs() -> Dict[str, Path]:
    root = _std_root()
    return {
        "采购": root / "02_Std/09_采购台账/F_采购部采购台账_std_split",
        "机具": root / "02_Std/10_机具费/E_机具费_std_split",
        "机票": root / "02_Std/11_人工费/B_飞签机票_std",
    }


def _labor_summary_path(paths: Paths) -> Path:
    return paths.group_sum_xlsx("人工费")


def _norm_name(name: str) -> str:
    text = safe_text(name).casefold().replace(" ", "").replace("　", "")
    if text.endswith("项目") and len(text) > 2:
        text = text[: -len("项目")]
    return text


def _stem(path: Path, suffix: str) -> str:
    name = path.name
    if name.endswith(suffix):
        return name[: -len(suffix)]
    return path.stem


def _targets(group_dir: Path) -> Dict[str, Path]:
    found: Dict[str, Path] = {}
    for path in group_dir.glob(f"*{TARGET_SUFFIX}"):
        if path.name.startswith("~$"):
            continue
        found[_stem(path, TARGET_SUFFIX)] = path
    return found


def _target_parts(stem: str) -> Tuple[str, str]:
    code, _, name = stem.partition("_")
    return code, name


def _match_stem(stem: str, targets: Dict[str, Path]) -> List[str]:
    if stem in _FILE_ALIASES:
        return [item for item in _FILE_ALIASES[stem] if item in targets]
    if stem in targets:
        return [stem]
    code, name = _target_parts(stem)
    norm = _norm_name(name)
    by_code = [
        key
        for key, (kcode, kname) in ((k, _target_parts(k)) for k in targets)
        if code and kcode == code and _norm_name(kname) == norm
    ]
    if by_code:
        return by_code
    by_code_only = [key for key in targets if code and _target_parts(key)[0] == code]
    if len(by_code_only) == 1:
        return by_code_only
    by_name = [key for key in targets if norm and _norm_name(_target_parts(key)[1]) == norm]
    if len(by_name) == 1:
        return by_name
    return []


def _clean(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item") and not isinstance(value, (bytes, str)):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    return value


def _write_frame(wb: Workbook, title: str, frame: pd.DataFrame) -> None:
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = None
    header_fill = PatternFill("solid", fgColor="DDEFE3")
    bold = Font(bold=True, color="1F3A2E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B7C7BC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = [str(col) for col in frame.columns]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = header_fill
        cell.font = bold
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = min(max(len(header) + 4, 12), 36)
    amount_cols = {
        idx
        for idx, header in enumerate(headers, start=1)
        if any(word in header for word in ("金额", "费用", "单价", "付款", "税", "泰铢"))
    }
    date_cols = {
        idx
        for idx, header in enumerate(headers, start=1)
        if "日期" in header or header in {"标准日期", "订票时间"}
    }
    for row_idx, values in enumerate(frame.itertuples(index=False, name=None), start=2):
        for col, raw in enumerate(values, start=1):
            value = _clean(raw)
            cell = ws.cell(row_idx, col, value)
            if col in date_cols and hasattr(value, "year"):
                cell.number_format = "yyyy-mm-dd"
            elif col in amount_cols and isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.number_format = "#,##0.00"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
    ws.row_dimensions[1].height = 24


def _collapse_frames(
    frames: Sequence[Tuple[str, pd.DataFrame]],
) -> List[Tuple[str, pd.DataFrame]]:
    """同一工作表被两份来源写到时，保留包含另一方的那份，避免短表盖掉长表。"""
    order: List[str] = []
    grouped: Dict[str, pd.DataFrame] = {}
    for title, frame in frames:
        if title not in grouped:
            order.append(title)
            grouped[title] = frame
            continue
        grouped[title] = _merge_same_sheet(title, grouped[title], frame)
    return [(title, grouped[title]) for title in order]


# 同一笔业务在「已填编码」和「未填编码」两份文件里会各出现一次。
_DEDUP_IGNORE = ("项目编码", "原始项目名称", "_来源文件")


def _business_key(frame: pd.DataFrame) -> pd.Series:
    columns = [col for col in frame.columns if col not in _DEDUP_IGNORE]
    if not columns:
        columns = list(frame.columns)
    return frame[columns].fillna("").astype(str).agg("\t".join, axis=1)


def _merge_same_sheet(title: str, left: pd.DataFrame, right: pd.DataFrame) -> pd.DataFrame:
    columns = list(dict.fromkeys([*left.columns, *right.columns]))
    first = left.reindex(columns=columns)
    second = right.reindex(columns=columns)
    first_keys = set(_business_key(first))
    second_keys = set(_business_key(second))
    if second_keys <= first_keys:
        print(f"    · {title} 后来源 {len(second)} 行已包含在先写入的 {len(first)} 行中")
        return first
    if first_keys <= second_keys:
        print(f"    · {title} 改用更全的来源 {len(second)} 行（含先前 {len(first)} 行）")
        return second
    merged = pd.concat([first, second], ignore_index=True)
    merged["_key"] = _business_key(merged)
    if "项目编码" in merged.columns:
        merged["_coded"] = merged["项目编码"].map(lambda value: 0 if safe_text(value) else 1)
        merged = merged.sort_values(["_key", "_coded"], kind="stable")
    merged = merged.drop_duplicates("_key", keep="first").drop(columns=["_key", "_coded"], errors="ignore")
    print(f"    ⚠️ {title} 两份来源有差异，按业务列合并 {len(first)}+{len(second)} → {len(merged)}")
    return merged.reset_index(drop=True)


def _copy_workbook_sheets(
    source: Path,
    renames: Dict[str, str],
    bucket: Dict[str, List[Tuple[str, pd.DataFrame]]],
    target_stems: Sequence[str],
) -> None:
    book = pd.ExcelFile(source)
    frames = []
    for sheet, title in renames.items():
        if sheet not in book.sheet_names:
            continue
        frames.append((title, pd.read_excel(book, sheet_name=sheet)))
    book.close()
    for stem in target_stems:
        bucket.setdefault(stem, []).extend(frames)


def _nonzero_columns(row: pd.Series, keep: Sequence[str]) -> pd.DataFrame:
    columns = list(keep)
    for col in row.index:
        if col in columns:
            continue
        number = pd.to_numeric(row[col], errors="coerce")
        if pd.notna(number) and float(number) != 0:
            columns.append(col)
    return pd.DataFrame([row.reindex(columns)])


def _labor_tables(path: Path) -> Dict[str, pd.DataFrame]:
    book = pd.ExcelFile(path)
    tables = {
        "总体汇总": pd.read_excel(book, sheet_name="总体汇总"),
        "项目按月宽表": pd.read_excel(book, sheet_name="项目按月宽表"),
        "项目一级科目汇总": pd.read_excel(book, sheet_name="项目一级科目汇总"),
        "宽表": pd.read_excel(book, sheet_name="一级二级科目宽表", header=None),
    }
    book.close()
    for key in ("总体汇总", "项目按月宽表", "项目一级科目汇总"):
        frame = tables[key]
        frame["项目编码"] = frame["项目编码"].map(safe_text)
        frame["项目名称"] = frame["项目名称"].map(safe_text)
    return tables


def _pick_labor_row(frame: pd.DataFrame, code: str, name: str) -> Optional[pd.Series]:
    """编码和名称一起对。ZGS 同时对应多个项目，不能只凭编码取唯一行。"""
    norm = _norm_name(name)
    if code:
        same = frame[(frame["项目编码"] == code) & (frame["项目名称"].map(_norm_name) == norm)]
        if len(same) == 1:
            return same.iloc[0]
        if len(same) > 1:
            return None
    named = frame[frame["项目名称"].map(_norm_name) == norm]
    if code:
        named = named[named["项目编码"].map(safe_text).isin({"", code})]
    if len(named) == 1:
        return named.iloc[0]
    return None


def _labor_frames_for_target(
    tables: Dict[str, pd.DataFrame],
    stem: str,
) -> Tuple[List[Tuple[str, pd.DataFrame]], Optional[object]]:
    code, name = _target_parts(stem)
    if code in {"未识别编码", "需手动核对", "无编码"}:
        code = ""
    frames: List[Tuple[str, pd.DataFrame]] = []
    summary = _match_labor_summary(tables["总体汇总"], stem, code, name)
    if summary is None:
        return frames, None
    ident = [col for col in ("业务分组", "项目编码", "项目名称") if col in summary.index]
    frames.append(("人工费汇总", pd.DataFrame([summary])))
    monthly = _match_labor_summary(tables["项目按月宽表"], stem, code, name)
    if monthly is not None:
        frames.append(("人工费按月", _nonzero_columns(monthly, ident)))
    subject = _match_labor_summary(tables["项目一级科目汇总"], stem, code, name)
    if subject is not None:
        frames.append(("人工费科目", _nonzero_columns(subject, ident)))
    wide = _labor_wide_frame(tables["宽表"], code, safe_text(summary.get("项目名称")) or name)
    if wide is not None and not wide.empty:
        frames.append(("人工费科目宽", wide))
    return frames, summary.name


def _match_labor_summary(
    frame: pd.DataFrame, stem: str, code: str, name: str
) -> Optional[pd.Series]:
    alias_stem = _LABOR_NAME_ALIASES.get(_norm_name(name))
    if alias_stem and alias_stem != stem:
        return None
    # 别名目标用原表里的名称再找一次。无编码的 Insee项目、子公司-现场费用走这里。
    if stem in _LABOR_NAME_ALIASES.values() and not code:
        wanted = [key for key, value in _LABOR_NAME_ALIASES.items() if value == stem]
        for norm in wanted:
            named = frame[frame["项目名称"].map(_norm_name) == norm]
            named = named[named["项目编码"].map(safe_text) == ""]
            if len(named) == 1:
                return named.iloc[0]
    return _pick_labor_row(frame, code, name)


def _labor_wide_frame(wide: pd.DataFrame, code: str, name: str) -> Optional[pd.DataFrame]:
    if wide.empty or wide.shape[1] < 3:
        return None
    codes = [safe_text(value) for value in wide.iloc[0].tolist()]
    names = [safe_text(value) for value in wide.iloc[1].tolist()]
    norm = _norm_name(name)
    chosen: Optional[int] = None
    if code:
        exact = [
            idx
            for idx, (col_code, col_name) in enumerate(zip(codes, names))
            if idx >= 2 and col_code == code and _norm_name(col_name) == norm
        ]
        if len(exact) == 1:
            chosen = exact[0]
    if chosen is None and norm:
        named = [idx for idx, col_name in enumerate(names) if idx >= 2 and _norm_name(col_name) == norm]
        if code:
            named = [idx for idx in named if codes[idx] in {"", code}]
        if len(named) == 1:
            chosen = named[0]
    if chosen is None:
        return None
    body = wide.iloc[2:, [0, 1, chosen]].copy()
    body.columns = ["一级科目", "二级科目", "付款金额"]
    body["付款金额"] = pd.to_numeric(body["付款金额"], errors="coerce").fillna(0.0)
    return body.reset_index(drop=True)


def attach_split_cost_sheets(group_dir: Path, paths: Optional[Paths] = None) -> List[str]:
    """把四类来源写入拆分目录里已有的项目工作簿。"""
    group_dir = Path(group_dir)
    targets = _targets(group_dir)
    if not targets:
        print(f"未找到拆分项目文件: {group_dir}")
        return []

    pending: Dict[str, List[Tuple[str, pd.DataFrame]]] = {}
    unmatched: List[str] = []
    skipped: List[str] = []
    dirs = _source_dirs()
    file_sources = (
        (dirs["采购"], "_采购数据.xlsx", {"采购数据": "采购数据"}),
        (dirs["机具"], "_机具费.xlsx", {"汇总": "机具费"}),
        (dirs["机票"], "_机票.xlsx", {"明细数据": "飞签机票", "月度汇总": "机票月度汇总"}),
    )
    for folder, suffix, renames in file_sources:
        if not folder.is_dir():
            print(f"  ⚠️ 找不到目录: {folder}")
            continue
        for path in sorted(folder.glob("*.xlsx")):
            if path.name.startswith("~$"):
                continue
            stem = _stem(path, suffix)
            if stem in _FILE_ALIASES and not _FILE_ALIASES[stem]:
                skipped.append(path.name)
                continue
            matched = _match_stem(stem, targets)
            if not matched:
                unmatched.append(path.name)
                continue
            _copy_workbook_sheets(path, renames, pending, matched)

    labor_unused: List[str] = []
    labor_path = _labor_summary_path(paths or Paths())
    if labor_path.is_file():
        tables = _labor_tables(labor_path)
        used_rows = set()
        for stem in targets:
            frames, row_id = _labor_frames_for_target(tables, stem)
            if row_id is not None:
                used_rows.add(row_id)
            if frames:
                pending.setdefault(stem, []).extend(frames)
        summary = tables["总体汇总"]
        for idx, row in summary.iterrows():
            if idx in used_rows:
                continue
            labor_unused.append(
                f"{safe_text(row.get('项目编码')) or '无编码'}_{safe_text(row.get('项目名称'))}"
            )
    else:
        print(f"  ⚠️ 找不到人工费汇总: {labor_path}")

    written: List[str] = []
    from openpyxl import load_workbook

    for stem, path in sorted(targets.items()):
        frames = _collapse_frames(pending.get(stem) or [])
        if not frames:
            continue
        wb = load_workbook(path)
        for title, _frame in frames:
            if title in wb.sheetnames:
                del wb[title]
        for title, frame in frames:
            _write_frame(wb, title, frame)
        wb.save(path)
        wb.close()
        written.append(str(path))
        titles = "、".join(title for title, _ in frames)
        print(f"  ✅ {path.name} ← {titles}")

    print(f"\n✅ 已写入 {len(written)} 个拆分项目文件")
    if skipped:
        print("重复来源已跳过:")
        for name in skipped:
            print(f"  · {name}")
    if unmatched:
        print("未对应到拆分项目、未写入:")
        for name in unmatched:
            print(f"  · {name}")
    if labor_unused:
        print("人工费汇总中没有对应拆分项目、未写入:")
        for name in labor_unused:
            print(f"  · {name}")
    return written
