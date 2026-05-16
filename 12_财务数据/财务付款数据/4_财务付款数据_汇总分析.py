import os
import re
import sys
import shutil
import importlib.util
from collections import OrderedDict, defaultdict
from pathlib import Path
from typing import Dict, List, Optional, cast

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ====================== 配置区 ======================
class Config:
    INPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/A_财务付款_raw"
    
    # 输出路径（同时支持两个位置）
    OUTPUT_Z = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report/财务付款数据汇总.xlsx"
    OUTPUT_ORIGINAL = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/C_财务付款_sum/财务付款数据汇总.xlsx"
    
    AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]
    SUBJECT_WIDE_SHEET = "一级二级科目宽表"
    SUBJECT_SUMMARY_ROW = "一级科目汇总"
    PROJECT_LEVEL1_SHEET = "项目一级科目汇总"


def iter_excel_files(input_root: Path) -> List[Path]:
    excel_files: List[Path] = []
    seen: set[Path] = set()
    for pattern in ("*.xlsx", "*.xls", "*.xlsm"):
        for path in input_root.rglob(pattern):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            excel_files.append(path)
    return sorted(excel_files)


def infer_time_tag_from_path(file_path: Path, input_root: Path) -> str:
    candidates: List[str] = []

    try:
        rel_parts = file_path.relative_to(input_root).parts
    except ValueError:
        rel_parts = file_path.parts

    for part in rel_parts[:-1]:
        candidates.append(part)
    candidates.append(file_path.stem)

    for text in reversed(candidates):
        tag = parse_folder_tag(text)
        digits = "".join(ch for ch in safe_text(tag) if ch.isdigit())
        if len(digits) >= 8 and digits.startswith("20"):
            return digits[:8]
        if len(digits) >= 6 and digits.startswith("20"):
            return digits[:6]

    return ""


CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve_project_info_dir() -> Optional[str]:
    """探测项目基本信息目录"""
    candidates = [
        os.path.normpath(os.path.join(CURRENT_DIR, "..", "..", "0_项目基本信息")),
        os.path.normpath(os.path.join(CURRENT_DIR, "..", "0_项目基本信息")),
        "/Users/wyattsun/Desktop/code/0_项目基本信息",
    ]
    for p in candidates:
        if os.path.isdir(p):
            return p
    return None


PROJECT_INFO_DIR = _resolve_project_info_dir()
if PROJECT_INFO_DIR and PROJECT_INFO_DIR not in sys.path:
    sys.path.append(PROJECT_INFO_DIR)


def load_formatter():
    """加载 Excel 样式模块"""
    if not PROJECT_INFO_DIR:
        print("⚠️ 未找到项目信息目录，跳过样式设置。")
        return lambda *a, **k: None

    module_path = Path(PROJECT_INFO_DIR) / "excel_formatting_utils.py"
    if not module_path.exists():
        print(f"⚠️ 未找到格式模块: {module_path}")
        return lambda *a, **k: None

    try:
        spec = importlib.util.spec_from_file_location("excel_formatting_utils", module_path)
        if spec is None or spec.loader is None:
            print(f"⚠️ 无法创建模块加载器: {module_path}")
            return lambda *a, **k: None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module.apply_common_excel_layout
    except Exception as e:
        print(f"⚠️ 加载格式模块失败: {e}")
        return lambda *a, **k: None


APPLY_COMMON_EXCEL_LAYOUT = load_formatter()


# ====================== 核心函数 ======================
def parse_folder_tag(folder_name: str) -> str:
    match = re.search(r"(20\d{6})", folder_name)
    if match:
        return match.group(1)
    digits = "".join(ch for ch in folder_name if ch.isdigit())
    return digits[:8] if len(digits) >= 8 else folder_name


def ymd_sort_key(value: str) -> int:
    text = safe_text(value)
    if re.fullmatch(r"\d{8}", text):
        return int(text)
    if re.fullmatch(r"\d{6}", text):
        return int(f"{text}01")
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 8:
        return int(digits[:8])
    if len(digits) >= 6:
        return int(f"{digits[:6]}01")
    return -1


def ymd_to_month(value: str) -> str:
    text = safe_text(value)
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 6:
        return f"{digits[:4]}-{digits[4:6]}"
    return text


def normalize_month_value(value, fallback_year: str = "") -> str:
    text = safe_text(value)
    if not text:
        return ""

    date_match = re.search(r"(20\d{2})[^\d]?(\d{1,2})(?:[^\d]?(\d{1,2}))?", text)
    if date_match:
        year = date_match.group(1)
        month = int(date_match.group(2))
        if 1 <= month <= 12:
            return f"{year}{month:02d}"

    pure_digits = "".join(ch for ch in text if ch.isdigit())
    if len(pure_digits) >= 8:
        y = pure_digits[:4]
        m = pure_digits[4:6]
        if y.startswith("20") and m.isdigit() and 1 <= int(m) <= 12:
            return f"{y}{int(m):02d}"
    if len(pure_digits) >= 6:
        y = pure_digits[:4]
        m = pure_digits[4:6]
        if y.startswith("20") and m.isdigit() and 1 <= int(m) <= 12:
            return f"{y}{int(m):02d}"

    if pure_digits.isdigit() and 1 <= int(pure_digits) <= 12 and fallback_year:
        return f"{fallback_year}{int(pure_digits):02d}"

    return ""


def extract_month_from_dataframe(df: pd.DataFrame, fallback_ymd: str) -> str:
    fallback_year = ""
    fallback_digits = "".join(ch for ch in safe_text(fallback_ymd) if ch.isdigit())
    if len(fallback_digits) >= 4:
        fallback_year = fallback_digits[:4]

    for col in ["年月", "日期"]:
        if col not in df.columns:
            continue
        for value in df[col].dropna().tolist():
            ym = normalize_month_value(value, fallback_year=fallback_year)
            if ym:
                return ym

    if "年" in df.columns and "月" in df.columns:
        for _, row in df[["年", "月"]].dropna(how="all").iterrows():
            year_val = safe_text(row.get("年"))
            month_val = safe_text(row.get("月"))
            year_digits = "".join(ch for ch in year_val if ch.isdigit())
            if len(year_digits) >= 4:
                year = year_digits[:4]
            else:
                year = fallback_year
            ym = normalize_month_value(month_val, fallback_year=year)
            if ym:
                return ym

    if "月" in df.columns:
        for value in df["月"].dropna().tolist():
            ym = normalize_month_value(value, fallback_year=fallback_year)
            if ym:
                return ym

    ym_fallback = normalize_month_value(fallback_ymd, fallback_year=fallback_year)
    if ym_fallback:
        return ym_fallback
    return fallback_ymd


def extract_month_from_row(row: pd.Series, fallback_ymd: str) -> str:
    fallback_year = ""
    fallback_digits = "".join(ch for ch in safe_text(fallback_ymd) if ch.isdigit())
    if len(fallback_digits) >= 4:
        fallback_year = fallback_digits[:4]

    ym = normalize_month_value(row.get("年月"), fallback_year=fallback_year)
    if ym:
        return ym

    ym = normalize_month_value(row.get("日期"), fallback_year=fallback_year)
    if ym:
        return ym

    if "年" in row.index and "月" in row.index:
        year_val = safe_text(row.get("年"))
        month_val = safe_text(row.get("月"))
        year_digits = "".join(ch for ch in year_val if ch.isdigit())
        year = year_digits[:4] if len(year_digits) >= 4 else fallback_year
        ym = normalize_month_value(month_val, fallback_year=year)
        if ym:
            return ym

    ym = normalize_month_value(row.get("月"), fallback_year=fallback_year)
    if ym:
        return ym

    ym_fallback = normalize_month_value(fallback_ymd, fallback_year=fallback_year)
    if ym_fallback:
        return ym_fallback
    return fallback_ymd


def safe_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def subject_sort_key(value: str):
    text = safe_text(value)
    match = re.match(r"^(\d+)[_\-\.、\s]*(.*)$", text)
    if match:
        return (0, int(match.group(1)), match.group(2))
    return (1, text)


def extract_subject_rows(file_path: Path) -> List[Dict[str, object]]:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"⚠️ 读取失败 {file_path.name}: {e}")
        return []

    if df.empty:
        return []

    if "实付金额" not in df.columns and "增值税额" not in df.columns:
        print(f"⚠️ 缺少金额列: {file_path.name}")
        return []

    amount_df = pd.DataFrame()
    for col in ["实付金额", "增值税额"]:
        if col in df.columns:
            amount_df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else:
            amount_df[col] = 0.0

    df["金额汇总"] = amount_df.sum(axis=1)
    df = df[df["金额汇总"] != 0].copy()
    if df.empty:
        return []

    project_code = ""
    if "项目编码" in df.columns:
        codes = df["项目编码"].dropna().astype(str).str.strip()
        codes = codes[~codes.str.lower().isin(["", "nan", "none"])]
        if not codes.empty:
            project_code = codes.iloc[0]

    project_name = ""
    if "项目名称" in df.columns:
        names = df["项目名称"].dropna().astype(str).str.strip()
        names = names[~names.str.lower().isin(["", "nan", "none"])]
        if not names.empty:
            project_name = names.iloc[0]
    if not project_name:
        project_name = file_path.stem

    rows: List[Dict[str, object]] = []
    for _, row in df.iterrows():
        level1 = safe_text(row.get("一级科目")) or safe_text(row.get("科目")) or "未分类一级科目"
        level2 = safe_text(row.get("二级科目")) or safe_text(row.get("科目")) or "未分类二级科目"
        rows.append(
            {
                "项目编码": safe_text(project_code),
                "项目名称": safe_text(project_name),
                "一级科目": level1,
                "二级科目": level2,
                "金额汇总": float(row.get("金额汇总", 0) or 0),
            }
        )
    return rows


def build_subject_wide_payload() -> Dict[str, object]:
    input_root = Path(Config.INPUT_ROOT)
    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    rows: List[Dict[str, object]] = []
    excel_files = iter_excel_files(input_root)
    print(f"\n检测到Excel文件数: {len(excel_files)}")
    for file_path in excel_files:
        rows.extend(extract_subject_rows(file_path))

    if not rows:
        return {
            "project_order": [],
            "layout": OrderedDict(),
            "project_map": {},
            "grand_map": {},
        }

    df = pd.DataFrame(rows)
    grouped = (
        df.groupby(["项目名称", "一级科目", "二级科目"], dropna=False, as_index=False)["金额汇总"]
        .sum()
    )

    level1_order = sorted(grouped["一级科目"].dropna().unique().tolist(), key=subject_sort_key)
    layout: "OrderedDict[str, List[str]]" = OrderedDict()
    for level1 in level1_order:
        level2_values = grouped.loc[grouped["一级科目"] == level1, "二级科目"].dropna().unique().tolist()
        layout[level1] = sorted(level2_values, key=subject_sort_key)

    project_order = sorted(grouped["项目名称"].dropna().unique().tolist(), key=subject_sort_key)

    project_map: Dict[str, Dict[str, Dict[str, float]]] = defaultdict(lambda: defaultdict(dict))
    grand_map: Dict[str, Dict[str, float]] = defaultdict(dict)
    for _, row in grouped.iterrows():
        project_name = safe_text(row.get("项目名称"))
        level1 = safe_text(row.get("一级科目"))
        level2 = safe_text(row.get("二级科目"))
        amount = float(row.get("金额汇总", 0) or 0)
        project_map[project_name][level1][level2] = amount
        grand_map[level1][level2] = grand_map[level1].get(level2, 0.0) + amount

    return {
        "project_order": project_order,
        "layout": layout,
        "project_map": project_map,
        "grand_map": grand_map,
    }


def build_project_level1_summary_dataframe() -> pd.DataFrame:
    """生成项目维度的一级科目横向汇总表。"""
    input_root = Path(Config.INPUT_ROOT)
    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    rows: List[Dict[str, object]] = []
    excel_files = iter_excel_files(input_root)
    for file_path in excel_files:
        rows.extend(extract_subject_rows(file_path))

    if not rows:
        return pd.DataFrame(columns=["项目编码", "项目名称"])

    df = pd.DataFrame(rows)
    grouped = (
        df.groupby(["项目编码", "项目名称", "一级科目"], dropna=False, as_index=False)["金额汇总"]
        .sum()
    )

    level1_order = sorted(grouped["一级科目"].dropna().unique().tolist(), key=subject_sort_key)
    wide_df = grouped.pivot_table(
        index=["项目编码", "项目名称"],
        columns="一级科目",
        values="金额汇总",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for col in level1_order:
        if col not in wide_df.columns:
            wide_df[col] = 0

    wide_df = wide_df[["项目编码", "项目名称"] + level1_order]
    return wide_df.sort_values(["项目编码", "项目名称"], ascending=[True, True]).reset_index(drop=True)


def extract_project_monthly_rows(file_path: Path, ymd: str) -> List[Dict]:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"⚠️ 读取失败 {file_path.name}: {e}")
        return []

    if df.empty or "实付金额" not in df.columns:
        if "实付金额" not in df.columns and not df.empty:
            print(f"⚠️ 缺少实付金额列: {file_path.name}")
        return []

    # 过滤全零金额行
    amt_cols = [c for c in Config.AMOUNT_COLUMNS if c in df.columns]
    if amt_cols:
        numeric = df[amt_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
        df = df[numeric.sum(axis=1) > 0]

    if df.empty:
        return []

    source_ym = extract_month_from_dataframe(df, ymd)

    # 项目编码 & 名称
    project_code = ""
    if "项目编码" in df.columns:
        codes = df["项目编码"].dropna().astype(str).str.strip()
        codes = codes[~codes.str.lower().isin(["", "nan", "none"])]
        if not codes.empty:
            project_code = codes.iloc[0]

    project_name = ""
    if "项目名称" in df.columns:
        names = df["项目名称"].dropna().astype(str).str.strip()
        names = names[~names.str.lower().isin(["", "nan", "none"])]
        if not names.empty:
            project_name = names.iloc[0]

    if not project_name:
        project_name = file_path.stem

    for col in Config.AMOUNT_COLUMNS:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["_年月日"] = df.apply(lambda r: extract_month_from_row(r, ymd), axis=1)
    fallback_month = ymd_to_month(source_ym)
    df["_月份"] = df["_年月日"].apply(lambda v: ymd_to_month(safe_text(v)) or fallback_month)

    grouped = df.groupby("_月份", dropna=False, as_index=False)[Config.AMOUNT_COLUMNS].sum()

    result_rows: List[Dict] = []
    for _, g_row in grouped.iterrows():
        row: Dict[str, object] = {
            "项目编码": safe_text(project_code),
            "项目名称": safe_text(project_name),
            "年月日": safe_text(g_row.get("_月份")),
        }
        for col in Config.AMOUNT_COLUMNS:
            row[col] = float(g_row.get(col, 0) or 0)
        result_rows.append(row)

    return result_rows


def build_summary_dataframe() -> pd.DataFrame:
    input_root = Path(Config.INPUT_ROOT)
    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    rows = []
    excel_files = iter_excel_files(input_root)
    print(f"\n检测到Excel文件数: {len(excel_files)}")
    for file_path in excel_files:
        ymd = infer_time_tag_from_path(file_path, input_root)
        file_rows = extract_project_monthly_rows(file_path, ymd)
        if file_rows:
            rows.extend(file_rows)

    if not rows:
        return pd.DataFrame(columns=["项目编码", "项目名称"] + Config.AMOUNT_COLUMNS)

    df = pd.DataFrame(rows)

    # 按项目聚合全部月份数据（总体汇总）
    agg_map = {col: (col, "sum") for col in Config.AMOUNT_COLUMNS}
    overall = df.groupby(["项目名称"], dropna=False).agg(
        项目编码=("项目编码", lambda s: next((v for v in s if safe_text(v)), "")),
        **agg_map,
    ).reset_index()

    return overall[["项目编码", "项目名称"] + Config.AMOUNT_COLUMNS].reset_index(drop=True)


def build_monthly_wide_dataframe() -> pd.DataFrame:
    input_root = Path(Config.INPUT_ROOT)
    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    rows = []
    excel_files = iter_excel_files(input_root)
    for file_path in excel_files:
        ymd = infer_time_tag_from_path(file_path, input_root)
        file_rows = extract_project_monthly_rows(file_path, ymd)
        for row in file_rows:
            month_key = ymd_to_month(safe_text(row.get("年月日")))
            rows.append(
                {
                    "项目编码": safe_text(row.get("项目编码")),
                    "项目名称": safe_text(row.get("项目名称")),
                    "月份": month_key,
                    "实付金额": float(row.get("实付金额", 0) or 0),
                    "增值税额": float(row.get("增值税额", 0) or 0),
                    "月度金额汇总": float(row.get("实付金额", 0) or 0) + float(row.get("增值税额", 0) or 0),
                }
            )

    if not rows:
        return pd.DataFrame(columns=["项目编码", "项目名称"])

    monthly_df = pd.DataFrame(rows)
    monthly_grouped = (
        monthly_df.groupby(["项目编码", "项目名称", "月份"], dropna=False, as_index=False)[["月度金额汇总"]]
        .sum()
    )

    month_order = sorted(monthly_grouped["月份"].dropna().unique(), key=ymd_sort_key)

    summary_wide = monthly_grouped.pivot_table(
        index=["项目编码", "项目名称"],
        columns="月份",
        values="月度金额汇总",
        aggfunc="sum",
        fill_value=0,
    )
    summary_wide = summary_wide.reindex(columns=month_order, fill_value=0)
    summary_wide.columns = [f"金额汇总_{m}" for m in summary_wide.columns]

    wide_df = summary_wide.reset_index()
    return wide_df


def write_subject_wide_sheet(wb, payload: Dict[str, object]):
    layout = cast(OrderedDict[str, List[str]], payload.get("layout", OrderedDict()))
    project_order = cast(List[str], payload.get("project_order", []))
    project_map = cast(Dict[str, Dict[str, Dict[str, float]]], payload.get("project_map", {}))
    grand_map = cast(Dict[str, Dict[str, float]], payload.get("grand_map", {}))

    ws = wb.create_sheet(title=Config.SUBJECT_WIDE_SHEET)
    ws.freeze_panes = "C2"
    ws.sheet_view.showGridLines = False
    # 让“小计”作为分组汇总行（位于明细行上方）
    ws.sheet_properties.outlinePr.summaryBelow = False

    title_fill = PatternFill("solid", fgColor="DDEFE3")
    header_fill = PatternFill("solid", fgColor="C9E1D1")
    summary_fill = PatternFill("solid", fgColor="EAF4EC")
    bold_font = Font(bold=True, color="1F3A2E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B7C7BC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.cell(1, 1, "一级科目")
    ws.cell(1, 2, "二级科目")
    for c in [1, 2]:
        cell = ws.cell(1, c)
        cell.fill = title_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border

    start_project_col = 3
    for idx, project_name in enumerate(project_order):
        cell = ws.cell(1, start_project_col + idx, project_name)
        cell.fill = title_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border

    summary_col = start_project_col + len(project_order)
    summary_head = ws.cell(1, summary_col, Config.SUBJECT_SUMMARY_ROW)
    summary_head.fill = title_fill
    summary_head.font = bold_font
    summary_head.alignment = center
    summary_head.border = border

    row_idx = 2
    for level1, level2_list in layout.items():
        subtotal_row = row_idx
        subtotal_first = ws.cell(subtotal_row, 1, level1)
        subtotal_second = ws.cell(subtotal_row, 2, "小计")
        for c in [subtotal_first, subtotal_second]:
            c.fill = header_fill
            c.font = bold_font
            c.alignment = center
            c.border = border

        level1_total_all_projects = 0.0
        for p_idx, project_name in enumerate(project_order):
            project_level1_map = project_map.get(project_name, {}).get(level1, {})
            value = sum(float(project_level1_map.get(level2, 0.0) or 0.0) for level2 in level2_list)
            cell = ws.cell(subtotal_row, start_project_col + p_idx, value)
            cell.number_format = '#,##0.00'
            cell.fill = header_fill
            cell.font = bold_font
            cell.alignment = center
            cell.border = border
            level1_total_all_projects += value

        total_cell = ws.cell(subtotal_row, summary_col, level1_total_all_projects)
        total_cell.number_format = '#,##0.00'
        total_cell.fill = header_fill
        total_cell.font = bold_font
        total_cell.alignment = center
        total_cell.border = border
        row_idx += 1

        detail_start_row = row_idx
        for level2 in level2_list:
            l1_cell = ws.cell(row_idx, 1, level1)
            l2_cell = ws.cell(row_idx, 2, level2)
            l1_cell.alignment = center
            l2_cell.alignment = center
            l1_cell.border = border
            l2_cell.border = border

            level2_total_all_projects = 0.0
            for p_idx, project_name in enumerate(project_order):
                value = float(project_map.get(project_name, {}).get(level1, {}).get(level2, 0.0) or 0.0)
                cell = ws.cell(row_idx, start_project_col + p_idx, value)
                cell.number_format = '#,##0.00'
                cell.alignment = center
                cell.border = border
                level2_total_all_projects += value

            total_cell = ws.cell(row_idx, summary_col, level2_total_all_projects)
            total_cell.number_format = '#,##0.00'
            total_cell.alignment = center
            total_cell.border = border
            row_idx += 1

        detail_end_row = row_idx - 1
        if detail_end_row >= detail_start_row:
            # 使用Excel组合功能：将二级科目明细行归入“小计”行下
            ws.row_dimensions.group(
                detail_start_row,
                detail_end_row,
                outline_level=1,
                hidden=False,
            )

    final_row = row_idx
    ws.cell(final_row, 1, "全部汇总")
    ws.cell(final_row, 2, "-")
    for c in [1, 2]:
        cell = ws.cell(final_row, c)
        cell.fill = summary_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border

    grand_total = 0.0
    for p_idx, project_name in enumerate(project_order):
        project_total = 0.0
        for level1, level2_list in layout.items():
            for level2 in level2_list:
                project_total += float(project_map.get(project_name, {}).get(level1, {}).get(level2, 0.0) or 0.0)
        cell = ws.cell(final_row, start_project_col + p_idx, project_total)
        cell.number_format = '#,##0.00'
        cell.fill = summary_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border
        grand_total += project_total

    total_cell = ws.cell(final_row, summary_col, grand_total)
    total_cell.number_format = '#,##0.00'
    total_cell.fill = summary_fill
    total_cell.font = bold_font
    total_cell.alignment = center
    total_cell.border = border

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 22
    for col_idx in range(start_project_col, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

    ws.row_dimensions[1].height = 30
    return ws


def apply_formatting(output_path: str):
    try:
        wb = load_workbook(output_path)
        for sheet in ["总体汇总", "项目按月宽表", Config.PROJECT_LEVEL1_SHEET]:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            headers = [cell.value for cell in ws[1]]
            amount_headers = [h for h in headers if h in Config.AMOUNT_COLUMNS]
            if sheet == "项目按月宽表":
                amount_headers = [
                    h for h in headers
                    if isinstance(h, str) and h.startswith("金额汇总_")
                ]
            elif sheet == Config.PROJECT_LEVEL1_SHEET:
                amount_headers = [
                    h for h in headers
                    if isinstance(h, str) and h not in {"项目编码", "项目名称"}
                ]

            APPLY_COMMON_EXCEL_LAYOUT(
                ws,
                numeric_start_col=ws.max_column + 1,
                amount_headers=amount_headers,
                header_row=1,
                header_height=40,
                freeze_panes="A2",
            )
        wb.save(output_path)
        print(f"✅ 样式应用完成: {Path(output_path).name}")
    except Exception as e:
        print(f"⚠️ 样式应用失败: {e}")


def main():
    print("🚀 开始处理财务付款数据汇总...")

    summary_df = build_summary_dataframe()
    monthly_wide_df = build_monthly_wide_dataframe()
    subject_wide_payload = build_subject_wide_payload()
    project_level1_df = build_project_level1_summary_dataframe()

    # 输出策略：先生成主文件，再复制
    output_paths = [Config.OUTPUT_Z, Config.OUTPUT_ORIGINAL]

    for path_str in output_paths:
        path = Path(path_str)
        path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="总体汇总", index=False)
            monthly_wide_df.to_excel(writer, sheet_name="项目按月宽表", index=False)
            project_level1_df.to_excel(writer, sheet_name=Config.PROJECT_LEVEL1_SHEET, index=False)
            write_subject_wide_sheet(writer.book, subject_wide_payload)

        apply_formatting(str(path))
        print(f"✅ 输出完成 → {path}")

    print("\n🎉 全部处理完成！")
    print(f"项目数量: {len(summary_df)}")
    print(f"输出列: {', '.join(summary_df.columns)}")
    print(f"按月宽表列数: {len(monthly_wide_df.columns)}")


if __name__ == "__main__":
    main()