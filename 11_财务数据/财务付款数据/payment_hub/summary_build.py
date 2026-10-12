# -*- coding: utf-8 -*-
"""
步骤 summary：B_std → 汇总分析 Excel（原 4_）
"""

from __future__ import annotations

import re
from collections import OrderedDict, defaultdict
from pathlib import Path
from typing import Dict, List, Tuple, cast

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import Paths
from .deps import get_excel_layout_fn
from .payment_amount import ensure_payment_amount, payment_total
from .utils import (
    infer_time_tag_from_path,
    iter_group_std_files,
    level1_sort_key,
    parse_std_project_filename,
    resolve_business_groups,
    safe_text,
    ymd_sort_key,
    ymd_to_month,
)

AMOUNT_COLUMNS = ["付款金额"]
# 不计入成本：一级科目名（taxonomy L1_exclude）
EXCLUDE_COST_L1_NAMES = {"不计入成本"}
# 总体汇总中追加的不计入/计入成本金额列
EXCLUDE_COST_AMOUNT_COLUMNS = [f"不计入成本_{c}" for c in AMOUNT_COLUMNS]
INCLUDE_COST_AMOUNT_COLUMNS = [f"计入成本_{c}" for c in AMOUNT_COLUMNS]
SUBJECT_WIDE_SHEET = "一级二级科目宽表"
SUBJECT_SUMMARY_ROW = "一级科目汇总"
PROJECT_LEVEL1_SHEET = "项目一级科目汇总"
EXCLUDED_SCAN_DIRS = {
    "C_财务付款_sum",
    "D_财务付款_学习池",
    "03_Sum",
    "0_Sum",
    "12_财务付款",
}


def is_exclude_cost_l1(level1) -> bool:
    """是否一级科目「不计入成本」。"""
    name = safe_text(level1)
    if not name:
        return False
    if name in EXCLUDE_COST_L1_NAMES:
        return True
    # 兼容带编号前缀，如「9_不计入成本」
    return name.endswith("不计入成本") or "不计入成本" in name


def _l2_detail_header(level1: str, level2: str, used_headers: set) -> str:
    """
    生成二级明细列名，保证不与一级合计列或其它列冲突。
    - 二级与一级同名 →「一级·其他」
    - 仍冲突 → 加「一级·」前缀或序号后缀
    """
    l1 = safe_text(level1) or "未分类一级科目"
    l2 = safe_text(level2) or "未分类二级科目"
    if l2 == l1:
        candidates = [f"{l1}·其他", f"{l1}·{l2}"]
    else:
        candidates = [l2, f"{l1}·{l2}"]
    for base in candidates:
        header = base
        n = 2
        while header in used_headers:
            header = f"{base}_{n}"
            n += 1
        if header not in used_headers:
            return header
    # 兜底
    header = f"{l1}·{l2}"
    n = 2
    while header in used_headers:
        header = f"{l1}·{l2}_{n}"
        n += 1
    return header


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
    pure = "".join(ch for ch in text if ch.isdigit())
    if len(pure) >= 6:
        y, m = pure[:4], pure[4:6]
        if y.startswith("20") and m.isdigit() and 1 <= int(m) <= 12:
            return f"{y}{int(m):02d}"
    if pure.isdigit() and 1 <= int(pure) <= 12 and fallback_year:
        return f"{fallback_year}{int(pure):02d}"
    return ""


def extract_month_from_dataframe(df: pd.DataFrame, fallback_ymd: str) -> str:
    fb_digits = "".join(ch for ch in safe_text(fallback_ymd) if ch.isdigit())
    fallback_year = fb_digits[:4] if len(fb_digits) >= 4 else ""
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
            year_digits = "".join(ch for ch in year_val if ch.isdigit())
            year = year_digits[:4] if len(year_digits) >= 4 else fallback_year
            ym = normalize_month_value(safe_text(row.get("月")), fallback_year=year)
            if ym:
                return ym
    ym = normalize_month_value(fallback_ymd, fallback_year=fallback_year)
    return ym or fallback_ymd


def extract_month_from_row(row: pd.Series, fallback_ymd: str) -> str:
    fb_digits = "".join(ch for ch in safe_text(fallback_ymd) if ch.isdigit())
    fallback_year = fb_digits[:4] if len(fb_digits) >= 4 else ""
    for key in ("年月", "日期"):
        ym = normalize_month_value(row.get(key), fallback_year=fallback_year)
        if ym:
            return ym
    if "年" in row.index and "月" in row.index:
        year_val = safe_text(row.get("年"))
        year_digits = "".join(ch for ch in year_val if ch.isdigit())
        year = year_digits[:4] if len(year_digits) >= 4 else fallback_year
        ym = normalize_month_value(safe_text(row.get("月")), fallback_year=year)
        if ym:
            return ym
    ym = normalize_month_value(row.get("月"), fallback_year=fallback_year)
    if ym:
        return ym
    return normalize_month_value(fallback_ymd, fallback_year=fallback_year) or fallback_ymd


def parse_project_from_filename(file_path: Path) -> tuple[str, str]:
    """兼容 编码_名称_财务付款数据 / 旧 编码_名称_(日期)。"""
    return parse_std_project_filename(file_path)


def pick_first_valid_text(series: pd.Series) -> str:
    values = series.dropna().astype(str).str.strip()
    values = values[~values.str.lower().isin(["", "nan", "none"])]
    return safe_text(values.iloc[0]) if not values.empty else ""


def _read_std_df(file_path: Path) -> pd.DataFrame:
    """优先读「明细数据」，否则第一张表。"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet = "明细数据" if "明细数据" in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(file_path, sheet_name=sheet)
    except Exception as e:
        print(f"⚠️ 读取失败 {file_path.name}: {e}")
        return pd.DataFrame()


def extract_subject_rows(file_path: Path) -> List[Dict[str, object]]:
    df = ensure_payment_amount(_read_std_df(file_path))
    if df.empty or "付款金额" not in df.columns:
        if not df.empty:
            print(f"⚠️ 缺少付款金额列: {file_path.name}")
        return []

    df["金额汇总"] = pd.to_numeric(df["付款金额"], errors="coerce").fillna(0.0)
    df = df[df["金额汇总"] != 0]
    if df.empty:
        return []

    file_code, file_name = parse_project_from_filename(file_path)
    project_code = pick_first_valid_text(df["项目编码"]) if "项目编码" in df.columns else file_code
    project_name = pick_first_valid_text(df["项目名称"]) if "项目名称" in df.columns else (file_name or file_path.stem)
    if not project_code:
        project_code = file_code
    if not project_name:
        project_name = file_name or file_path.stem

    rows = []
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


def extract_project_monthly_rows(file_path: Path, ymd: str) -> List[Dict]:
    df = ensure_payment_amount(_read_std_df(file_path))
    if df.empty or "付款金额" not in df.columns:
        if not df.empty:
            print(f"⚠️ 缺少付款金额列: {file_path.name}")
        return []

    df["付款金额"] = pd.to_numeric(df["付款金额"], errors="coerce").fillna(0.0)
    df = df[df["付款金额"] != 0]
    if df.empty:
        return []

    source_ym = extract_month_from_dataframe(df, ymd)
    file_code, file_name = parse_project_from_filename(file_path)
    project_code = pick_first_valid_text(df["项目编码"]) if "项目编码" in df.columns else file_code
    project_name = pick_first_valid_text(df["项目名称"]) if "项目名称" in df.columns else (file_name or file_path.stem)
    if not project_code:
        project_code = file_code
    if not project_name:
        project_name = file_name or file_path.stem

    for col in AMOUNT_COLUMNS:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df = df.copy()
    df["_年月日"] = df.apply(lambda r: extract_month_from_row(r, ymd), axis=1)
    fallback_month = ymd_to_month(source_ym)
    df["_月份"] = df["_年月日"].apply(lambda v: ymd_to_month(safe_text(v)) or fallback_month)
    grouped = df.groupby("_月份", dropna=False, as_index=False)[AMOUNT_COLUMNS].sum()

    result = []
    for _, g_row in grouped.iterrows():
        row = {
            "项目编码": safe_text(project_code),
            "项目名称": safe_text(project_name),
            "年月日": safe_text(g_row.get("_月份")),
        }
        for col in AMOUNT_COLUMNS:
            row[col] = float(g_row.get(col, 0) or 0)
        result.append(row)
    return result


def _empty_subject_wide_payload() -> Dict[str, object]:
    return {"project_order": [], "layout": OrderedDict(), "project_map": {}, "grand_map": {}}


def _order_project_columns(pairs: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """项目列按项目编码倒序。没有编码的列排在后面，仍按名称排序。"""
    unique = list(dict.fromkeys(pairs))
    coded = [item for item in unique if item[0]]
    blank = [item for item in unique if not item[0]]
    coded.sort(key=lambda item: (item[0], item[1]), reverse=True)
    blank.sort(key=lambda item: level1_sort_key(item[1]))
    return coded + blank


def _subject_wide_payload_from_frame(df: pd.DataFrame) -> Dict[str, object]:
    """明细 → 一级二级科目宽表。project_order 为 (项目编码, 项目名称)。"""
    if df is None or df.empty or "金额汇总" not in df.columns:
        return _empty_subject_wide_payload()

    work = df.copy()
    if "项目编码" not in work.columns:
        work["项目编码"] = ""
    if "项目名称" not in work.columns:
        work["项目名称"] = ""
    work["项目编码"] = work["项目编码"].map(safe_text)
    work["项目名称"] = work["项目名称"].map(safe_text)
    if "一级科目" in work.columns:
        work["一级科目"] = work["一级科目"].map(safe_text).replace("", "未分类一级科目")
    else:
        work["一级科目"] = "未分类一级科目"
    if "二级科目" in work.columns:
        work["二级科目"] = work["二级科目"].map(safe_text).replace("", "未分类二级科目")
    else:
        work["二级科目"] = "未分类二级科目"
    work["金额汇总"] = pd.to_numeric(work["金额汇总"], errors="coerce").fillna(0.0)
    work = work[work["金额汇总"] != 0]
    if work.empty:
        return _empty_subject_wide_payload()

    grouped = work.groupby(
        ["项目编码", "项目名称", "一级科目", "二级科目"], dropna=False, as_index=False
    )["金额汇总"].sum()
    level1_order = sorted(grouped["一级科目"].dropna().unique().tolist(), key=level1_sort_key)
    layout: OrderedDict[str, List[str]] = OrderedDict()
    for level1 in level1_order:
        l2s = grouped.loc[grouped["一级科目"] == level1, "二级科目"].dropna().unique().tolist()
        layout[level1] = sorted(l2s, key=level1_sort_key)

    pairs = [
        (safe_text(code), safe_text(name))
        for code, name in zip(grouped["项目编码"], grouped["项目名称"])
    ]
    project_order = _order_project_columns(pairs)
    project_map: Dict[Tuple[str, str], Dict[str, Dict[str, float]]] = defaultdict(
        lambda: defaultdict(dict)
    )
    grand_map: Dict[str, Dict[str, float]] = defaultdict(dict)
    for _, row in grouped.iterrows():
        key = (safe_text(row["项目编码"]), safe_text(row["项目名称"]))
        l1, l2 = safe_text(row["一级科目"]), safe_text(row["二级科目"])
        amount = float(row["金额汇总"] or 0)
        project_map[key][l1][l2] = amount
        grand_map[l1][l2] = grand_map[l1].get(l2, 0.0) + amount
    return {
        "project_order": project_order,
        "layout": layout,
        "project_map": project_map,
        "grand_map": grand_map,
    }


def build_subject_wide_payload(std_dir: Path) -> Dict[str, object]:
    rows: List[Dict[str, object]] = []
    files = iter_group_std_files(std_dir)
    print(f"\n检测到本组 std 文件数: {len(files)}  （目录: {std_dir}）")
    for fp in files:
        rows.extend(extract_subject_rows(fp))
    if not rows:
        return _empty_subject_wide_payload()
    return _subject_wide_payload_from_frame(pd.DataFrame(rows))


def build_project_level1_summary_dataframe(
    std_dir: Path,
    *,
    group_label: str = "",
) -> Tuple[pd.DataFrame, List[List[str]]]:
    """
    项目 × 一级科目宽表，每个一级科目合计后展开其二级类目。

    返回值第二项为各组二级明细列名（每组对应一个一级，列连续），
    供写入 Excel 时做列组合（大纲折叠）。
    """
    label = group_label or Path(std_dir).name
    rows = []
    for fp in iter_group_std_files(std_dir):
        rows.extend(extract_subject_rows(fp))
    if not rows:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"]), []

    df = pd.DataFrame(rows)
    df["一级科目"] = df["一级科目"].map(safe_text).replace("", "未分类一级科目")
    df["二级科目"] = df["二级科目"].map(safe_text).replace("", "未分类二级科目")

    grouped_l1 = df.groupby(
        ["项目编码", "项目名称", "一级科目"], dropna=False, as_index=False
    )["金额汇总"].sum()
    level1_order = sorted(
        grouped_l1["一级科目"].dropna().unique().tolist(), key=level1_sort_key
    )
    wide = grouped_l1.pivot_table(
        index=["项目编码", "项目名称"],
        columns="一级科目",
        values="金额汇总",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    for col in level1_order:
        if col not in wide.columns:
            wide[col] = 0

    # 各一级下的二级明细
    used_headers: set = set(level1_order) | {"业务分组", "项目编码", "项目名称"}
    # l1 -> [(l2_raw, header), ...]
    l1_l2_headers: Dict[str, List[Tuple[str, str]]] = {}
    detail_groups: List[List[str]] = []

    for l1 in level1_order:
        sub = df.loc[df["一级科目"] == l1]
        if sub.empty:
            l1_l2_headers[l1] = []
            continue
        l2_order = sorted(sub["二级科目"].dropna().unique().tolist(), key=level1_sort_key)
        pairs: List[Tuple[str, str]] = []
        for l2 in l2_order:
            header = _l2_detail_header(l1, l2, used_headers)
            used_headers.add(header)
            pairs.append((l2, header))
        l1_l2_headers[l1] = pairs
        if pairs:
            detail_groups.append([h for _, h in pairs])

    # 透视为 项目×(一级,二级)，再按 header 填列
    grouped_l2 = df.groupby(
        ["项目编码", "项目名称", "一级科目", "二级科目"], dropna=False, as_index=False
    )["金额汇总"].sum()
    amount_lookup: Dict[Tuple[str, str, str, str], float] = {}
    for _, row in grouped_l2.iterrows():
        key = (
            safe_text(row["项目编码"]),
            safe_text(row["项目名称"]),
            safe_text(row["一级科目"]),
            safe_text(row["二级科目"]),
        )
        amount_lookup[key] = float(row["金额汇总"] or 0)

    # 列序：一级合计 + 紧跟其二级明细
    col_order = ["业务分组", "项目编码", "项目名称"]
    all_detail_headers: List[str] = []
    for l1 in level1_order:
        col_order.append(l1)
        for _l2, header in l1_l2_headers.get(l1, []):
            col_order.append(header)
            all_detail_headers.append(header)

    # 填充二级明细列
    for header in all_detail_headers:
        wide[header] = 0.0

    # 建立 header -> (l1, l2) 反查
    header_to_l1_l2: Dict[str, Tuple[str, str]] = {}
    for l1, pairs in l1_l2_headers.items():
        for l2, header in pairs:
            header_to_l1_l2[header] = (l1, l2)

    if all_detail_headers and not wide.empty:
        detail_values = {h: [] for h in all_detail_headers}
        for _, prow in wide.iterrows():
            pcode = safe_text(prow["项目编码"])
            pname = safe_text(prow["项目名称"])
            for header in all_detail_headers:
                l1, l2 = header_to_l1_l2[header]
                detail_values[header].append(
                    amount_lookup.get((pcode, pname, l1, l2), 0.0)
                )
        for header in all_detail_headers:
            wide[header] = detail_values[header]

    wide.insert(0, "业务分组", label)
    wide = wide[col_order]
    return (
        wide.sort_values(["业务分组", "项目编码", "项目名称"]).reset_index(drop=True),
        detail_groups,
    )


def apply_level2_column_groups(ws, detail_groups: List[List[str]]) -> None:
    """对各一级科目下的二级明细列启用 Excel 列组合（大纲，汇总列在左侧）。"""
    if not detail_groups:
        return
    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i for i, h in enumerate(headers, start=1) if h}

    # 合计列在明细左侧
    ws.sheet_properties.outlinePr.summaryRight = False

    detail_fill = PatternFill("solid", fgColor="F3F0E6")
    detail_font = Font(bold=True, color="5C4E2A")

    for group in detail_groups:
        cols = [header_to_col[h] for h in group if h in header_to_col]
        if not cols:
            continue
        start_col, end_col = min(cols), max(cols)
        start_letter = get_column_letter(start_col)
        end_letter = get_column_letter(end_col)
        ws.column_dimensions.group(
            start_letter, end_letter, outline_level=1, hidden=False
        )
        for col_idx in cols:
            cell = ws.cell(1, col_idx)
            cell.fill = detail_fill
            cell.font = detail_font


def build_summary_from_detail_df(
    detail_df: pd.DataFrame, *, group_label: str = ""
) -> pd.DataFrame:
    """
    从已分类明细 DataFrame 直接汇总（快路径）。

    金额列是付款金额（预扣税额 + 实付金额）：
    - 付款金额：全部合计
    - 不计入成本_付款金额：一级科目=不计入成本 的合计
    - 计入成本_付款金额：全部 − 不计入成本
    """
    label = group_label or ""
    base_cols = (
        ["业务分组", "项目编码", "项目名称"]
        + AMOUNT_COLUMNS
        + EXCLUDE_COST_AMOUNT_COLUMNS
        + INCLUDE_COST_AMOUNT_COLUMNS
    )
    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=base_cols)

    df = ensure_payment_amount(detail_df)
    for col in AMOUNT_COLUMNS:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    # 去掉金额全 0
    df = df[df[AMOUNT_COLUMNS].sum(axis=1) != 0]
    if df.empty:
        return pd.DataFrame(columns=base_cols)

    if "项目编码" not in df.columns:
        df["项目编码"] = ""
    if "项目名称" not in df.columns:
        df["项目名称"] = ""
    df["项目编码"] = df["项目编码"].map(safe_text)
    df["项目名称"] = df["项目名称"].map(safe_text)

    l1_col = "一级科目" if "一级科目" in df.columns else None
    if l1_col:
        excl_mask = df[l1_col].map(is_exclude_cost_l1)
    else:
        excl_mask = pd.Series(False, index=df.index)

    for col in AMOUNT_COLUMNS:
        df[f"_excl_{col}"] = df[col].where(excl_mask, 0.0)
        df[f"_incl_{col}"] = df[col].where(~excl_mask, 0.0)

    agg_cols = (
        AMOUNT_COLUMNS
        + [f"_excl_{c}" for c in AMOUNT_COLUMNS]
        + [f"_incl_{c}" for c in AMOUNT_COLUMNS]
    )
    overall = df.groupby(["项目编码", "项目名称"], dropna=False, as_index=False)[
        agg_cols
    ].sum()

    rename = {}
    for c in AMOUNT_COLUMNS:
        rename[f"_excl_{c}"] = f"不计入成本_{c}"
        rename[f"_incl_{c}"] = f"计入成本_{c}"
    overall = overall.rename(columns=rename)
    overall.insert(0, "业务分组", label)

    # 金额合计与付款金额同一口径
    overall["不计入成本_金额合计"] = overall["不计入成本_付款金额"]
    overall["计入成本_金额合计"] = overall["计入成本_付款金额"]

    out_cols = (
        ["业务分组", "项目编码", "项目名称"]
        + AMOUNT_COLUMNS
        + EXCLUDE_COST_AMOUNT_COLUMNS
        + ["不计入成本_金额合计"]
        + INCLUDE_COST_AMOUNT_COLUMNS
        + ["计入成本_金额合计"]
    )
    for c in out_cols:
        if c not in overall.columns:
            overall[c] = 0.0
    return overall[out_cols]


def build_monthly_wide_from_detail_df(
    detail_df: pd.DataFrame, *, group_label: str = ""
) -> pd.DataFrame:
    label = group_label or ""
    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"])

    df = ensure_payment_amount(detail_df)
    for col in AMOUNT_COLUMNS:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if "项目编码" not in df.columns:
        df["项目编码"] = ""
    if "项目名称" not in df.columns:
        df["项目名称"] = ""
    df["项目编码"] = df["项目编码"].map(safe_text)
    df["项目名称"] = df["项目名称"].map(safe_text)

    # 月份键
    months = []
    for _, row in df.iterrows():
        ym = ""
        if "年" in df.columns and "月" in df.columns:
            y = pd.to_numeric(row.get("年"), errors="coerce")
            m = pd.to_numeric(row.get("月"), errors="coerce")
            if pd.notna(y) and pd.notna(m) and 1 <= int(m) <= 12:
                ym = f"{int(y)}-{int(m):02d}"
        if not ym and "日期" in df.columns:
            parsed = pd.to_datetime(row.get("日期"), errors="coerce")
            if pd.notna(parsed):
                ym = parsed.strftime("%Y-%m")
        months.append(ym or "未识别")
    df["_月份"] = months
    df["_月度金额"] = pd.to_numeric(df["付款金额"], errors="coerce").fillna(0.0)

    monthly = (
        df.groupby(["项目编码", "项目名称", "_月份"], dropna=False, as_index=False)["_月度金额"]
        .sum()
        .rename(columns={"_月份": "月份", "_月度金额": "月度金额汇总"})
    )
    monthly.insert(0, "业务分组", label)
    if monthly.empty:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"])

    month_order = sorted(monthly["月份"].dropna().unique(), key=ymd_sort_key)
    wide = monthly.pivot_table(
        index=["业务分组", "项目编码", "项目名称"],
        columns="月份",
        values="月度金额汇总",
        aggfunc="sum",
        fill_value=0,
    )
    wide = wide.reindex(columns=month_order, fill_value=0)
    wide.columns = [f"金额汇总_{m}" for m in wide.columns]
    return wide.reset_index()


def build_subject_wide_from_detail_df(detail_df: pd.DataFrame) -> Dict[str, object]:
    if detail_df is None or detail_df.empty:
        return _empty_subject_wide_payload()

    df = ensure_payment_amount(detail_df).copy()
    if "付款金额" not in df.columns:
        return _empty_subject_wide_payload()
    df["金额汇总"] = pd.to_numeric(df["付款金额"], errors="coerce").fillna(0.0)
    return _subject_wide_payload_from_frame(df)


def build_project_level1_from_detail_df(
    detail_df: pd.DataFrame, *, group_label: str = ""
) -> Tuple[pd.DataFrame, List[List[str]]]:
    """基于已分类明细构建项目×一级科目宽表（复用文件路径版逻辑）。"""
    # 临时写 subject rows 结构再走原逻辑的简化版
    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"]), []

    df = ensure_payment_amount(detail_df)
    df["金额汇总"] = pd.to_numeric(df["付款金额"], errors="coerce").fillna(0.0)
    rows = []
    for _, row in df.iterrows():
        if float(row.get("金额汇总") or 0) == 0:
            continue
        rows.append(
            {
                "项目编码": safe_text(row.get("项目编码")),
                "项目名称": safe_text(row.get("项目名称")),
                "一级科目": safe_text(row.get("一级科目")) or safe_text(row.get("科目")) or "未分类一级科目",
                "二级科目": safe_text(row.get("二级科目")) or safe_text(row.get("科目")) or "未分类二级科目",
                "金额汇总": float(row.get("金额汇总") or 0),
            }
        )
    if not rows:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"]), []

    # 复用 build_project_level1_summary_dataframe 的核心：构造伪 std 不可行，内联调用现有 group 逻辑
    # 用临时 DataFrame 走原函数需要文件；这里复制关键代码路径
    tmp = pd.DataFrame(rows)
    # 借用 build_project_level1 的文件版：写入内存结构 — 直接调用简化实现
    label = group_label or ""
    # 调用现有函数需要 Path；改为：拼一个 minimal frame 用 groupby 逻辑
    # 最简：一级科目透视 + 二级明细（与文件版一致）
    from copy import deepcopy  # noqa: F401

    # 直接用文件版内部逻辑：把 rows 当作 extract_subject_rows 结果
    # 构造 virtual by writing to build_project_level1_summary_dataframe via monkey — 太重
    # 这里调用已有 build_project_level1_summary_dataframe 的核心代码块：
    # 为减少重复，用 tempfile 写 subject-only 不合适
    # 采用：groupby 一级透视（不含二级展开的精简版）若 detail_groups 为空也能用
    grouped_l1 = tmp.groupby(
        ["项目编码", "项目名称", "一级科目"], dropna=False, as_index=False
    )["金额汇总"].sum()
    level1_order = sorted(
        grouped_l1["一级科目"].dropna().unique().tolist(), key=level1_sort_key
    )
    wide = grouped_l1.pivot_table(
        index=["项目编码", "项目名称"],
        columns="一级科目",
        values="金额汇总",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    for col in level1_order:
        if col not in wide.columns:
            wide[col] = 0

    used_headers: set = set(level1_order) | {"业务分组", "项目编码", "项目名称"}
    l1_l2_headers: Dict[str, List[Tuple[str, str]]] = {}
    detail_groups: List[List[str]] = []
    for l1 in level1_order:
        sub = tmp.loc[tmp["一级科目"] == l1]
        l2_order = sorted(sub["二级科目"].dropna().unique().tolist(), key=level1_sort_key)
        pairs: List[Tuple[str, str]] = []
        for l2 in l2_order:
            header = _l2_detail_header(l1, l2, used_headers)
            used_headers.add(header)
            pairs.append((l2, header))
        l1_l2_headers[l1] = pairs
        if pairs:
            detail_groups.append([h for _, h in pairs])

    grouped_l2 = tmp.groupby(
        ["项目编码", "项目名称", "一级科目", "二级科目"], dropna=False, as_index=False
    )["金额汇总"].sum()
    amount_lookup: Dict[Tuple[str, str, str, str], float] = {}
    for _, row in grouped_l2.iterrows():
        key = (
            safe_text(row["项目编码"]),
            safe_text(row["项目名称"]),
            safe_text(row["一级科目"]),
            safe_text(row["二级科目"]),
        )
        amount_lookup[key] = float(row["金额汇总"] or 0)

    col_order = ["业务分组", "项目编码", "项目名称"]
    all_detail_headers: List[str] = []
    for l1 in level1_order:
        col_order.append(l1)
        for _l2, header in l1_l2_headers.get(l1, []):
            col_order.append(header)
            all_detail_headers.append(header)

    for header in all_detail_headers:
        wide[header] = 0.0
    header_to_l1_l2 = {
        h: (l1, l2) for l1, pairs in l1_l2_headers.items() for l2, h in pairs
    }
    if all_detail_headers and not wide.empty:
        detail_values = {h: [] for h in all_detail_headers}
        for _, prow in wide.iterrows():
            pcode = safe_text(prow["项目编码"])
            pname = safe_text(prow["项目名称"])
            for header in all_detail_headers:
                l1, l2 = header_to_l1_l2[header]
                detail_values[header].append(
                    amount_lookup.get((pcode, pname, l1, l2), 0.0)
                )
        for header in all_detail_headers:
            wide[header] = detail_values[header]

    wide.insert(0, "业务分组", label)
    wide = wide[col_order]
    return (
        wide.sort_values(["业务分组", "项目编码", "项目名称"]).reset_index(drop=True),
        detail_groups,
    )


def build_summary_dataframe(std_dir: Path, *, group_label: str = "") -> pd.DataFrame:
    rows = []
    files = iter_group_std_files(std_dir)
    print(f"\n[{group_label or std_dir.name}] 本组 std 文件数: {len(files)}")
    print(f"  扫描目录: {std_dir}")
    for fp in files:
        ymd = infer_time_tag_from_path(fp, std_dir)
        rows.extend(extract_project_monthly_rows(fp, ymd))
    if not rows:
        cols = ["业务分组", "项目编码", "项目名称"] + AMOUNT_COLUMNS
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    overall = df.groupby(["项目编码", "项目名称"], dropna=False, as_index=False)[
        AMOUNT_COLUMNS
    ].sum()
    label = group_label or Path(std_dir).name
    overall.insert(0, "业务分组", label)
    return overall[["业务分组", "项目编码", "项目名称"] + AMOUNT_COLUMNS]


def build_monthly_wide_dataframe(std_dir: Path, *, group_label: str = "") -> pd.DataFrame:
    rows = []
    label = group_label or Path(std_dir).name
    for fp in iter_group_std_files(std_dir):
        ymd = infer_time_tag_from_path(fp, std_dir)
        for row in extract_project_monthly_rows(fp, ymd):
            month_key = ymd_to_month(safe_text(row.get("年月日")))
            rows.append(
                {
                    "业务分组": label,
                    "项目编码": safe_text(row.get("项目编码")),
                    "项目名称": safe_text(row.get("项目名称")),
                    "月份": month_key,
                    "月度金额汇总": float(row.get("付款金额", 0) or 0),
                }
            )
    if not rows:
        return pd.DataFrame(columns=["业务分组", "项目编码", "项目名称"])
    monthly = pd.DataFrame(rows)
    grouped = monthly.groupby(
        ["业务分组", "项目编码", "项目名称", "月份"], dropna=False, as_index=False
    )[["月度金额汇总"]].sum()
    month_order = sorted(grouped["月份"].dropna().unique(), key=ymd_sort_key)
    wide = grouped.pivot_table(
        index=["业务分组", "项目编码", "项目名称"],
        columns="月份",
        values="月度金额汇总",
        aggfunc="sum",
        fill_value=0,
    )
    wide = wide.reindex(columns=month_order, fill_value=0)
    wide.columns = [f"金额汇总_{m}" for m in wide.columns]
    return wide.reset_index()


def _paint(cell, fill, font, alignment, border) -> None:
    cell.fill = fill
    cell.font = font
    cell.alignment = alignment
    cell.border = border


def _project_column_keys(project_order) -> List[Tuple[str, str]]:
    keys: List[Tuple[str, str]] = []
    for item in project_order or []:
        if isinstance(item, (tuple, list)) and len(item) >= 2:
            keys.append((safe_text(item[0]), safe_text(item[1])))
        else:
            keys.append(("", safe_text(item)))
    return keys


def write_subject_wide_sheet(wb, payload: Dict[str, object]):
    layout = cast(OrderedDict, payload.get("layout", OrderedDict()))
    projects = _project_column_keys(payload.get("project_order", []))
    project_map = cast(Dict, payload.get("project_map", {}))

    ws = wb.create_sheet(title=SUBJECT_WIDE_SHEET)
    # 第 1 行项目编码、第 2 行项目名称，左侧两列冻结
    ws.freeze_panes = "C3"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False

    title_fill = PatternFill("solid", fgColor="DDEFE3")
    header_fill = PatternFill("solid", fgColor="C9E1D1")
    summary_fill = PatternFill("solid", fgColor="EAF4EC")
    bold_font = Font(bold=True, color="1F3A2E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B7C7BC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    code_row = 1
    name_row = 2
    ws.merge_cells(start_row=code_row, start_column=1, end_row=code_row, end_column=2)
    code_label = ws.cell(code_row, 1, "项目编码")
    _paint(code_label, title_fill, bold_font, center, border)
    _paint(ws.cell(code_row, 2), title_fill, bold_font, center, border)

    ws.cell(name_row, 1, "一级科目")
    ws.cell(name_row, 2, "二级科目")
    for col in (1, 2):
        _paint(ws.cell(name_row, col), title_fill, bold_font, center, border)

    start_col = 3
    for idx, (project_code, project_name) in enumerate(projects):
        code_cell = ws.cell(code_row, start_col + idx, project_code or None)
        _paint(code_cell, title_fill, bold_font, center, border)
        name_cell = ws.cell(name_row, start_col + idx, project_name)
        _paint(name_cell, title_fill, bold_font, center, border)

    summary_col = start_col + len(projects)
    _paint(ws.cell(code_row, summary_col, None), title_fill, bold_font, center, border)
    head = ws.cell(name_row, summary_col, SUBJECT_SUMMARY_ROW)
    _paint(head, title_fill, bold_font, center, border)

    def _level2_amount(key: Tuple[str, str], level1: str, level2: str) -> float:
        return float(project_map.get(key, {}).get(level1, {}).get(level2, 0) or 0)

    row_idx = 3
    for level1, level2_list in layout.items():
        subtotal_row = row_idx
        for col, val in ((1, level1), (2, "小计")):
            cell = ws.cell(subtotal_row, col, val)
            _paint(cell, header_fill, bold_font, center, border)

        level1_total = 0.0
        for p_idx, key in enumerate(projects):
            value = sum(_level2_amount(key, level1, l2) for l2 in level2_list)
            cell = ws.cell(subtotal_row, start_col + p_idx, value)
            cell.number_format = "#,##0.00"
            _paint(cell, header_fill, bold_font, center, border)
            level1_total += value
        total_cell = ws.cell(subtotal_row, summary_col, level1_total)
        total_cell.number_format = "#,##0.00"
        _paint(total_cell, header_fill, bold_font, center, border)
        row_idx += 1

        detail_start = row_idx
        for level2 in level2_list:
            for col, val in ((1, level1), (2, level2)):
                cell = ws.cell(row_idx, col, val)
                cell.alignment = center
                cell.border = border
            l2_total = 0.0
            for p_idx, key in enumerate(projects):
                value = _level2_amount(key, level1, level2)
                cell = ws.cell(row_idx, start_col + p_idx, value)
                cell.number_format = "#,##0.00"
                cell.alignment = center
                cell.border = border
                l2_total += value
            cell = ws.cell(row_idx, summary_col, l2_total)
            cell.number_format = "#,##0.00"
            cell.alignment = center
            cell.border = border
            row_idx += 1
        detail_end = row_idx - 1
        if detail_end >= detail_start:
            ws.row_dimensions.group(detail_start, detail_end, outline_level=1, hidden=False)

    final_row = row_idx
    for col, val in ((1, "全部汇总"), (2, "-")):
        cell = ws.cell(final_row, col, val)
        _paint(cell, summary_fill, bold_font, center, border)
    grand = 0.0
    for p_idx, key in enumerate(projects):
        project_total = 0.0
        for level1, level2_list in layout.items():
            for level2 in level2_list:
                project_total += _level2_amount(key, level1, level2)
        cell = ws.cell(final_row, start_col + p_idx, project_total)
        cell.number_format = "#,##0.00"
        _paint(cell, summary_fill, bold_font, center, border)
        grand += project_total
    cell = ws.cell(final_row, summary_col, grand)
    cell.number_format = "#,##0.00"
    _paint(cell, summary_fill, bold_font, center, border)

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 22
    for col_idx in range(start_col, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    ws.row_dimensions[code_row].height = 22
    ws.row_dimensions[name_row].height = 30
    return ws


def apply_formatting(
    output_path: Path, detail_groups: List[List[str]] | None = None
) -> None:
    layout_fn = get_excel_layout_fn()
    detail_groups = detail_groups or []
    try:
        wb = load_workbook(output_path)
        for sheet in ["总体汇总", "项目按月宽表", PROJECT_LEVEL1_SHEET]:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            headers = [cell.value for cell in ws[1]]
            if sheet == "项目按月宽表":
                amount_headers = [
                    h for h in headers if isinstance(h, str) and h.startswith("金额汇总_")
                ]
            elif sheet == PROJECT_LEVEL1_SHEET:
                amount_headers = [
                    h for h in headers if isinstance(h, str) and h not in {"项目编码", "项目名称"}
                ]
            else:
                amount_headers = [
                    h
                    for h in headers
                    if isinstance(h, str)
                    and (
                        h in AMOUNT_COLUMNS
                        or h.startswith("不计入成本_")
                        or h.startswith("计入成本_")
                    )
                ]
            layout_fn(
                ws,
                numeric_start_col=ws.max_column + 1,
                amount_headers=amount_headers,
                header_row=1,
                header_height=40,
                freeze_panes="A2",
            )
            if sheet == PROJECT_LEVEL1_SHEET and detail_groups:
                apply_level2_column_groups(ws, detail_groups)
        wb.save(output_path)
        print(f"✅ 样式应用完成: {output_path.name}")
    except Exception as e:
        print(f"⚠️ 样式应用失败: {e}")


def _write_summary_workbook(
    path: Path,
    *,
    summary_df: pd.DataFrame,
    monthly_wide_df: pd.DataFrame,
    subject_payload: Dict[str, object],
    project_level1_df: pd.DataFrame,
    detail_groups,
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="总体汇总", index=False)
        monthly_wide_df.to_excel(writer, sheet_name="项目按月宽表", index=False)
        project_level1_df.to_excel(writer, sheet_name=PROJECT_LEVEL1_SHEET, index=False)
        write_subject_wide_sheet(writer.book, subject_payload)
    apply_formatting(path, detail_groups=detail_groups)


def _cleanup_legacy_sum_files(path: Path, group_label: str = "") -> None:
    """删除旧组名子目录 / 无下划线文件名的汇总，避免和扁平新文件并存。"""
    path = Path(path)
    name = (group_label or "").strip()
    candidates = [path.parent / "财务付款数据汇总.xlsx"]
    if name:
        candidates.extend(
            [
                path.parent / f"{name}_财务付款数据汇总.xlsx",
                path.parent / name / f"{name}_财务付款数据汇总.xlsx",
                path.parent / name / f"{name}_财务付款数据_汇总.xlsx",
            ]
        )
    for legacy in candidates:
        if not legacy.is_file():
            continue
        try:
            if legacy.resolve() == path.resolve():
                continue
            legacy.unlink()
            print(f"   🗑️ 移除旧汇总: {legacy}")
            parent = legacy.parent
            if parent != path.parent and parent.is_dir() and not any(parent.iterdir()):
                parent.rmdir()
        except OSError as exc:
            print(f"  ⚠️ 无法删除旧汇总 {legacy}: {exc}")


def run_summary_for_detail_df(
    detail_df: pd.DataFrame,
    *,
    sum_xlsx: Path,
    group_label: str = "",
) -> List[str]:
    """从已分类明细 DataFrame 写汇总（快路径）。"""
    label = group_label or ""
    print(f"  [{label}] 快路径汇总：明细 {len(detail_df)} 行（已含科目分类）")

    summary_df = build_summary_from_detail_df(detail_df, group_label=label)
    monthly_wide_df = build_monthly_wide_from_detail_df(detail_df, group_label=label)
    subject_payload = build_subject_wide_from_detail_df(detail_df)
    project_level1_df, detail_groups = build_project_level1_from_detail_df(
        detail_df, group_label=label
    )

    path = Path(sum_xlsx)
    _cleanup_legacy_sum_files(path, group_label or "")
    _write_summary_workbook(
        path,
        summary_df=summary_df,
        monthly_wide_df=monthly_wide_df,
        subject_payload=subject_payload,
        project_level1_df=project_level1_df,
        detail_groups=detail_groups,
    )
    written = [str(path)]
    print(f"   → {path}")

    if detail_groups:
        n_detail = sum(len(g) for g in detail_groups)
        print(
            f"   [{label}] 一级科目展开二级明细: {len(detail_groups)} 组 / 共 {n_detail} 列"
        )
    pay = payment_total(summary_df)
    excl = (
        float(pd.to_numeric(summary_df["不计入成本_金额合计"], errors="coerce").sum())
        if "不计入成本_金额合计" in summary_df.columns and not summary_df.empty
        else 0.0
    )
    incl = (
        float(pd.to_numeric(summary_df["计入成本_金额合计"], errors="coerce").sum())
        if "计入成本_金额合计" in summary_df.columns and not summary_df.empty
        else 0.0
    )
    print(
        f"  ✅ [{label}] 汇总完成：项目 {len(summary_df)} 个 · 付款金额 {pay:,.2f}"
        f" · 计入成本 {incl:,.2f} · 不计入成本 {excl:,.2f}"
    )
    return written


def run_summary_for_group(
    group_std_dir: Path,
    *,
    sum_xlsx: Path,
    group_label: str = "",
) -> List[str]:
    """对单个业务分组目录内的 std 生成汇总（绝不扫其他分组）。"""
    group_std_dir = Path(group_std_dir).resolve()
    label = group_label or group_std_dir.name
    if not group_std_dir.is_dir():
        print(f"  [{label}] std 目录不存在，跳过汇总")
        return []

    n_files = len(iter_group_std_files(group_std_dir))
    print(f"  [{label}] 仅汇总目录: {group_std_dir}  （{n_files} 个项目文件）")

    summary_df = build_summary_dataframe(group_std_dir, group_label=label)
    monthly_wide_df = build_monthly_wide_dataframe(group_std_dir, group_label=label)
    subject_payload = build_subject_wide_payload(group_std_dir)
    project_level1_df, detail_groups = build_project_level1_summary_dataframe(
        group_std_dir, group_label=label
    )

    path = Path(sum_xlsx)
    _cleanup_legacy_sum_files(path, label)
    _write_summary_workbook(
        path,
        summary_df=summary_df,
        monthly_wide_df=monthly_wide_df,
        subject_payload=subject_payload,
        project_level1_df=project_level1_df,
        detail_groups=detail_groups,
    )
    written = [str(path)]
    print(f"   → {path}")

    if detail_groups:
        n_detail = sum(len(g) for g in detail_groups)
        print(
            f"   [{label}] 一级科目展开二级明细: {len(detail_groups)} 组 / 共 {n_detail} 列"
        )
    pay = payment_total(summary_df)
    print(f"  ✅ [{label}] 汇总完成：项目 {len(summary_df)} 个 · 付款金额合计 {pay:,.2f}")
    return written


def _archive_ungrouped_root_outputs(paths: Paths, active_groups: List[str]) -> None:
    """
    把 Sum 根目录下旧的「未分组」总表挪到 _历史未分组，
    避免与 {组名}_财务付款数据_汇总.xlsx 混淆。
    """
    import shutil
    from datetime import datetime

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    root_names = {
        paths.sum_xlsx.name,
        paths.cashflow_html.name,
        paths.amount_mismatch_xlsx.name,
    }
    root = Path(paths.sum_dir)
    if not root.is_dir():
        return
    archive = root / "_历史未分组" / stamp
    for name in root_names:
        fp = root / name
        if not fp.is_file():
            continue
        archive.mkdir(parents=True, exist_ok=True)
        dest = archive / name
        try:
            shutil.move(str(fp), str(dest))
            print(f"  📦 已归档旧根目录文件: {fp.name} → {dest}")
        except OSError as exc:
            print(f"  ⚠️ 归档失败 {fp}: {exc}")


def _write_group_index(paths: Paths, group_stats: List[dict]) -> Path:
    """Std/分组汇总索引.xlsx：每组一行，标明路径与合计。"""
    index_path = Path(paths.group_index_xlsx)
    df = pd.DataFrame(group_stats)
    index_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(index_path, index=False)
    print(f"  📑 分组索引: {index_path}")
    return index_path


def run_summary(paths: Paths) -> List[str]:
    """
    优先快路径：读 A2_classified 已分类明细 → 直接汇总
    回退：扫 B_std/{组名}/ 多项目文件

    写出 Sum/{组名}_财务付款数据_汇总.xlsx
    """
    from .classify_raw_std import list_classified_groups, load_classified_detail

    written: List[str] = []
    group_stats: List[dict] = []

    classified_groups = list_classified_groups(paths)
    if classified_groups:
        print(f"汇总快路径（A2_classified）: {', '.join(g for g, _ in classified_groups)}")
        print(f"  源: {paths.classified_dir}/{{组名}}_classified.xlsx")
        print(f"  出: {paths.sum_dir}/{{组名}}_财务付款数据_汇总.xlsx")

        for group_name, class_xlsx in classified_groups:
            print(f"\n{'=' * 64}")
            print(f"  分组 summary · {group_name}（classified）")
            print(f"{'=' * 64}")
            detail = load_classified_detail(class_xlsx)
            sum_path = paths.group_sum_xlsx(group_name)
            outs = run_summary_for_detail_df(
                detail,
                sum_xlsx=sum_path,
                group_label=group_name,
            )
            written.extend(outs)
            try:
                sdf = pd.read_excel(sum_path, sheet_name="总体汇总")
                n_proj = len(sdf)
                pay = payment_total(sdf)
            except Exception:
                n_proj, pay = 0, 0.0
            group_stats.append(
                {
                    "业务分组": group_name,
                    "数据源": str(class_xlsx),
                    "汇总文件": str(sum_path),
                    "项目数": n_proj,
                    "付款金额合计": pay,
                }
            )
    else:
        # 回退旧路径
        std_dir = Path(paths.std_dir)
        if not std_dir.is_dir():
            raise FileNotFoundError(
                f"无 A2_classified 且 std 目录不存在: {std_dir}；请先 classify 或 std"
            )
        groups = resolve_business_groups(paths.raw_dir, paths.std_dir)
        if not groups:
            print("未发现业务分组，跳过 summary")
            return []
        print(f"汇总回退路径（B_std）: {', '.join(g for g, _ in groups)}")
        for group_name, group_std in groups:
            print(f"\n{'=' * 64}")
            print(f"  分组 summary · {group_name}")
            print(f"{'=' * 64}")
            sum_path = paths.group_sum_xlsx(group_name)
            outs = run_summary_for_group(
                group_std,
                sum_xlsx=sum_path,
                group_label=group_name,
            )
            written.extend(outs)
            try:
                sdf = pd.read_excel(sum_path, sheet_name="总体汇总")
                n_proj = len(sdf)
                pay = payment_total(sdf)
            except Exception:
                n_proj, pay = 0, 0.0
            group_stats.append(
                {
                    "业务分组": group_name,
                    "数据源": str(group_std),
                    "汇总文件": str(sum_path),
                    "项目数": n_proj,
                    "付款金额合计": pay,
                }
            )

    _archive_ungrouped_root_outputs(paths, [g["业务分组"] for g in group_stats])
    _write_group_index(paths, group_stats)

    print(f"\n✅ summary 全部完成：{len(group_stats)} 个分组，{len(written)} 个文件")
    for st in group_stats:
        print(f"   · [{st['业务分组']}] 项目 {st['项目数']} → {st['汇总文件']}")
    return written
