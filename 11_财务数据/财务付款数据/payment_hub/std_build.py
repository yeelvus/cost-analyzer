# -*- coding: utf-8 -*-
"""
步骤 std：把 A2 已分类明细按项目拆到 B_std。

不读 A_raw，也不再分类。没有 A2 时直接停止，请先跑 classify。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .config import Paths
from .deps import (
    get_excel_layout_fn,
    get_project_matcher_module,
    get_subject_mapper,
    project_match_db_path,
)
from .utils import (
    has_actual_payment,
    infer_time_tag_from_path,
    iter_excel_files,
    list_raw_groups,
    normalize_text,
    safe_filename,
    to_numeric_series,
)

TARGET_COLUMNS = [
    "项目名称",
    "科目",
    "摘要",
    "供应商名称",
    "应付金额",
    "预扣税额",
    "实付金额",
    "增值税额",
    "不含税额",
    "日期",
    "月",
    "年",
]
MIN_REQUIRED_COLUMNS = ["项目名称", "摘要", "实付金额"]
AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]
DEFAULT_CATEGORY = "6_管理费"

# B_std 子目录：按「全量数据最新月份」是否有支出归类
STD_DIR_HAS_SPEND = "最新月份有支出"
STD_DIR_NO_SPEND = "最新月份未支出"
STD_FILE_TAG = "财务付款数据"

COLUMN_ALIASES = {
    "项目名称": ["项目名称", "项目", "项目名", "工程名称", "工程项目"],
    "科目": ["科目", "会计科目", "财务科目"],
    "摘要": ["摘要", "业务摘要", "凭证摘要", "说明"],
    "供应商名称": ["供应商名称", "供应商", "收款单位", "单位名称", "往来单位"],
    "应付金额": ["应付金额", "应付", "应付款", "应支付金额"],
    "预扣税额": ["预扣税额", "预扣税", "扣税", "税额"],
    "实付金额": ["实付金额", "实付", "实际付款", "支付金额", "付款金额"],
    "增值税额": ["增值税额", "增值税", "进项税额", "销项税额"],
    "不含税额": ["不含税额", "不含税金额", "未税金额", "税前金额"],
    "日期": ["日期", "付款日期", "业务日期", "发生日期"],
    "月": ["月", "月份", "所属月份", "会计期间(月)"],
    "年": ["年", "年份", "所属年份", "会计年度"],
}


def find_column_mapping(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mapping: Dict[str, Optional[str]] = {}
    normalized_to_original = {normalize_text(col): col for col in df.columns}
    for target in TARGET_COLUMNS:
        candidates = COLUMN_ALIASES.get(target, [target])
        found_col = None
        for alias in candidates:
            norm_alias = normalize_text(alias)
            if norm_alias in normalized_to_original:
                found_col = normalized_to_original[norm_alias]
                break
        if found_col is None:
            for alias in candidates:
                norm_alias = normalize_text(alias)
                for norm_col, origin_col in normalized_to_original.items():
                    if norm_alias and (norm_alias in norm_col or norm_col in norm_alias):
                        found_col = origin_col
                        break
                if found_col is not None:
                    break
        mapping[target] = found_col
    return mapping


def normalize_date_month_year(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "日期" in out.columns:
        parsed = pd.to_datetime(out["日期"], errors="coerce")
        out["日期"] = parsed.dt.strftime("%Y-%m-%d").where(parsed.notna(), out["日期"])
        if "年" in out.columns:
            year_num = pd.to_numeric(out["年"], errors="coerce")
            out["年"] = year_num.where(year_num.notna(), parsed.dt.year)
        if "月" in out.columns:
            month_num = pd.to_numeric(out["月"], errors="coerce")
            out["月"] = month_num.where(month_num.notna(), parsed.dt.month)
    if "年" in out.columns:
        out["年"] = pd.to_numeric(out["年"], errors="coerce").astype("Int64")
    if "月" in out.columns:
        out["月"] = pd.to_numeric(out["月"], errors="coerce").astype("Int64")
    return out


def process_single_sheet(file_path: Path, sheet_name: str) -> Optional[pd.DataFrame]:
    try:
        raw_df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as exc:
        print(f"  ⚠️ 读取失败: {file_path.name} | {sheet_name} | {exc}")
        return None
    if raw_df.empty:
        return None

    mapping = find_column_mapping(raw_df)
    missing = [c for c in MIN_REQUIRED_COLUMNS if mapping.get(c) is None]
    if missing:
        print(f"    - 跳过 sheet: {sheet_name}，缺列: {', '.join(missing)}")
        return None

    out = pd.DataFrame()
    for col in TARGET_COLUMNS:
        src = mapping.get(col)
        out[col] = raw_df[src] if src is not None else pd.NA

    for amount_col in AMOUNT_COLUMNS:
        out[amount_col] = to_numeric_series(out[amount_col])

    row_sum = out[AMOUNT_COLUMNS].fillna(0).sum(axis=1)
    out = out[row_sum.ne(0)]
    out = normalize_date_month_year(out)
    out = out.dropna(how="all", subset=TARGET_COLUMNS)
    return None if out.empty else out


def enrich_project_fields(
    df: pd.DataFrame,
    *,
    prompt: bool = False,
    interactive: bool = True,
) -> pd.DataFrame:
    """
    用 ProjectMatcher 标准化项目名称/编码。

    - **不做模糊匹配**，仅精确/规范化别名 + 标准项目名
    - prompt=False（默认）：未命中的保留原名，不弹窗
    - prompt=True：仅对未命中名称弹窗，由用户决定映射；跳过后记入 unmatched_aliases
    """
    mod = get_project_matcher_module()
    ProjectMatcher = getattr(mod, "ProjectMatcher", None)
    db_path = project_match_db_path()

    print(
        f"  项目匹配: 精确别名 only（无模糊）"
        f"{'；弹窗补录=开' if prompt else '；弹窗补录=关（可用 --prompt-project 打开）'}"
    )

    if ProjectMatcher is not None:
        m = ProjectMatcher(db_path=db_path, load_info=True)
        raw_series = df["项目名称"].fillna("").astype(str).str.strip()
        unique = sorted({n for n in raw_series.tolist() if n and n.lower() != "nan"})
        mapped = m.match_batch(
            unique,
            prompt=prompt,
            interactive=interactive,
            window_title="财务付款数据项目匹配",
            fuzzy=False,
        )
        code_map = {k: v.get("project_code", "") for k, v in mapped.items()}
        name_map = {k: v.get("project_name", k) for k, v in mapped.items()}
        enriched = df.copy()
        enriched["项目编码"] = raw_series.map(code_map).fillna("")
        enriched["项目名称"] = raw_series.map(name_map).fillna(raw_series)
        enriched = enriched[["项目编码"] + TARGET_COLUMNS]
        m.save()
        no_code = [k for k in unique if not code_map.get(k)]
        if no_code:
            print(f"  未匹配/无编码: {len(no_code)} 个"
                  f"  例: {no_code[:8]}{'…' if len(no_code) > 8 else ''}")
        return enriched

    # 极旧模块回退：同样仅精确匹配，不模糊
    db = mod.load_or_init_db(db_path, [])
    raw_series = df["项目名称"].fillna("").astype(str).str.strip()
    unique_names = sorted({n for n in raw_series.tolist() if n and n.lower() != "nan"})
    code_map: Dict[str, str] = {}
    name_map: Dict[str, str] = {}
    pending: List[str] = []
    for raw_name in unique_names:
        hit = mod.match_project(raw_name, db, fuzzy=False)
        if hit and hit.get("source") not in {"unmatched_keep", "unmatched"}:
            code_map[raw_name] = hit.get("project_code", "") or ""
            name_map[raw_name] = hit.get("project_name", raw_name)
        else:
            pending.append(raw_name)
            code_map[raw_name] = ""
            name_map[raw_name] = raw_name
    if pending and prompt:
        print(f"  ⚠️ 未匹配项目 {len(pending)} 个，弹出手动匹配…")
        mod.prompt_unmatched_project_mappings(
            pending,
            db,
            window_title="财务付款数据项目匹配",
            db_path=db_path,
            interactive=interactive,
        )
        for raw_name in pending:
            hit = mod.match_project(raw_name, db, fuzzy=False)
            if hit and hit.get("source") not in {"unmatched_keep", "unmatched"}:
                code_map[raw_name] = hit.get("project_code", "") or ""
                name_map[raw_name] = hit.get("project_name", raw_name)
            else:
                mod.mark_unmatched(db, raw_name)
                code_map[raw_name] = ""
                name_map[raw_name] = raw_name
    elif pending:
        # 不弹窗：不强制写入 unmatched_aliases，便于下次 --prompt-project 再确认
        print(f"  未匹配保留原名（未弹窗）: {len(pending)} 个")
    enriched = df.copy()
    enriched["项目编码"] = raw_series.map(code_map).fillna("")
    enriched["项目名称"] = raw_series.map(name_map).fillna(raw_series)
    enriched = enriched[["项目编码"] + TARGET_COLUMNS]
    mod.save_db(db_path, db)
    return enriched


def _cell_str(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def append_subject_levels(df: pd.DataFrame, mapper) -> pd.DataFrame:
    """为明细附加一/二/三级分类。相同 (科目,摘要,供应商) 只分类一次。"""
    enriched = df.copy()
    if "科目" not in enriched.columns and "摘要" not in enriched.columns:
        enriched["一级科目"] = DEFAULT_CATEGORY
        enriched["二级科目"] = "待复核"
        return enriched

    vendor_col = (
        "供应商名称"
        if "供应商名称" in enriched.columns
        else ("供应商" if "供应商" in enriched.columns else None)
    )

    n = len(enriched)
    subjects = (
        enriched["科目"].map(_cell_str).tolist()
        if "科目" in enriched.columns
        else [""] * n
    )
    summaries = (
        enriched["摘要"].map(_cell_str).tolist()
        if "摘要" in enriched.columns
        else [""] * n
    )
    if vendor_col:
        vendors = enriched[vendor_col].map(_cell_str).tolist()
    else:
        vendors = [""] * n

    # 仅有实付金额的行参与分类
    pay_ok = (
        enriched["实付金额"].map(has_actual_payment).tolist()
        if "实付金额" in enriched.columns
        else [False] * n
    )
    triples_all = list(zip(subjects, summaries, vendors))
    triples_pay = [t for t, ok in zip(triples_all, pay_ok) if ok]
    unique_triples = list(dict.fromkeys(triples_pay))
    cache = {}

    engine = getattr(mapper, "engine", None)
    use_full = hasattr(mapper, "classify_full")

    if engine is not None and hasattr(engine, "classify_unique_map"):
        cache = engine.classify_unique_map(unique_triples) if unique_triples else {}
    else:
        for t in unique_triples:
            try:
                if use_full:
                    cache[t] = mapper.classify_full(t[0], t[1], t[2] or None)
                else:
                    l1, l2 = mapper.classify(t[0], t[1])
                    cache[t] = (l1, l2)
            except Exception:
                cache[t] = None

    l1s, l2s, l3s, ids, layers, confs = [], [], [], [], [], []
    n_skip = 0
    for t, ok in zip(triples_all, pay_ok):
        if not ok:
            l1s.append("")
            l2s.append("")
            l3s.append("")
            ids.append("")
            layers.append("skip_no_payment")
            confs.append(None)
            n_skip += 1
            continue
        r = cache.get(t)
        try:
            if r is None:
                l1, l2, l3 = DEFAULT_CATEGORY, "待复核", ""
                leaf, layer, conf = "", "error", 0.0
            elif isinstance(r, tuple):
                l1, l2 = r[0], r[1]
                l3, leaf, layer, conf = "", "", "", None
            else:
                l1, l2 = r.l1_name, r.l2_name
                l3 = getattr(r, "l3_name", "") or ""
                leaf = getattr(r, "l3_id", "") or getattr(r, "l2_id", "") or ""
                layer = getattr(r, "layer", "")
                conf = getattr(r, "confidence", None)
        except Exception:
            l1, l2, l3 = DEFAULT_CATEGORY, "待复核", ""
            leaf, layer, conf = "", "error", 0.0

        l1s.append(str(l1 or "").strip() or DEFAULT_CATEGORY)
        l2s.append(str(l2 or "").strip() or "待复核")
        l3s.append(str(l3 or "").strip())
        ids.append(leaf)
        layers.append(layer)
        confs.append(conf)

    if unique_triples or n_skip:
        print(
            f"  分类缓存: {len(unique_triples)} 种组合 → 有实付 {n - n_skip} 行"
            f"（跳过空实付 {n_skip}）"
            f"（复用率 {(1 - len(unique_triples) / max(n - n_skip, 1)) * 100:.0f}%）"
        )

    enriched["一级科目"] = l1s
    enriched["二级科目"] = l2s
    enriched["三级科目"] = l3s
    enriched["分类ID"] = ids
    enriched["命中层"] = layers
    enriched["置信度"] = confs

    prefer = ["一级科目", "二级科目", "三级科目", "分类ID", "命中层", "置信度"]
    cols = [c for c in enriched.columns if c not in prefer]
    if "摘要" in cols:
        i = cols.index("摘要")
        cols = cols[:i] + prefer + cols[i:]
    else:
        cols = cols + prefer
    return enriched[cols]


def build_detail_wan_sheet(detail_df: pd.DataFrame) -> pd.DataFrame:
    wan = detail_df.copy()
    for col in AMOUNT_COLUMNS:
        if col in wan.columns:
            wan[col] = pd.to_numeric(wan[col], errors="coerce") / 10000.0
    return wan


def format_last_payment_date(df: pd.DataFrame) -> str:
    if df is None or df.empty or "日期" not in df.columns:
        return "未识别日期"
    parsed = pd.to_datetime(df["日期"], errors="coerce")
    if parsed.notna().any():
        return parsed.max().strftime("%Y-%m-%d")
    return "未识别日期"


def year_month_series(df: pd.DataFrame) -> pd.Series:
    """从 日期 / 年+月 推导 YYYY-MM 序列。"""
    if df is None or df.empty:
        return pd.Series(dtype=object)

    result = pd.Series(pd.NA, index=df.index, dtype=object)

    if "日期" in df.columns:
        parsed = pd.to_datetime(df["日期"], errors="coerce")
        result = parsed.dt.strftime("%Y-%m")

    if "年" in df.columns and "月" in df.columns:
        years = pd.to_numeric(df["年"], errors="coerce")
        months = pd.to_numeric(df["月"], errors="coerce")
        valid = years.notna() & months.notna() & (months >= 1) & (months <= 12)
        parts = pd.Series(pd.NA, index=df.index, dtype=object)
        if valid.any():
            parts.loc[valid] = (
                years.loc[valid].astype(int).astype(str)
                + "-"
                + months.loc[valid].astype(int).map(lambda m: f"{int(m):02d}")
            )
        result = result.where(result.notna(), parts)

    return result.fillna("").astype(str)


def row_spend_amount(df: pd.DataFrame) -> pd.Series:
    """行支出口径：预扣税额 + 实付金额（缺列按 0）。"""
    withhold = (
        pd.to_numeric(df["预扣税额"], errors="coerce").fillna(0.0)
        if "预扣税额" in df.columns
        else 0.0
    )
    actual = (
        pd.to_numeric(df["实付金额"], errors="coerce").fillna(0.0)
        if "实付金额" in df.columns
        else 0.0
    )
    return withhold + actual


def detect_global_latest_spend_month(df: pd.DataFrame) -> str:
    """全量数据中有支出的最新 YYYY-MM；无则返回空。"""
    if df is None or df.empty:
        return ""
    ym = year_month_series(df)
    spend = row_spend_amount(df)
    mask = (ym != "") & (spend != 0)
    if not mask.any():
        # 退而求其次：任意有月份的最大月
        valid = ym[ym != ""]
        return str(valid.max()) if not valid.empty else ""
    return str(ym[mask].max())


def project_has_spend_in_month(project_df: pd.DataFrame, latest_ym: str) -> bool:
    if not latest_ym or project_df is None or project_df.empty:
        return False
    ym = year_month_series(project_df)
    spend = row_spend_amount(project_df)
    return bool(((ym == latest_ym) & (spend != 0)).any())


def build_std_project_filename(code: str, name: str) -> str:
    """项目编码_项目名称_财务付款数据.xlsx"""
    code_part = safe_filename(code) if code else "未识别编码"
    name_part = safe_filename(name) if name else "未命名项目"
    return f"{code_part}_{name_part}_{STD_FILE_TAG}.xlsx"


def apply_output_format(output_path: Path, classify_data: Path | None = None) -> None:
    def _to_float(value):
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except Exception:
            return None

    layout_fn = get_excel_layout_fn()
    wb = load_workbook(output_path)
    if not wb.worksheets:
        wb.save(output_path)
        return

    amount_headers = set(AMOUNT_COLUMNS)
    mismatch_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")

    for ws in wb.worksheets:
        if ws.title in {"_分类选项"}:
            continue
        try:
            layout_fn(
                ws,
                numeric_start_col=1,
                amount_headers=amount_headers,
                header_row=1,
                header_height=40,
                freeze_panes="A2",
            )
        except Exception as exc:
            print(f"  ⚠️ 样式失败: {exc}")

        headers = {
            str(cell.value).strip(): idx
            for idx, cell in enumerate(ws[1], start=1)
            if cell.value is not None
        }
        p_col = headers.get("应付金额")
        w_col = headers.get("预扣税额")
        a_col = headers.get("实付金额")
        if p_col and w_col and a_col:
            for row_idx in range(2, ws.max_row + 1):
                payable = _to_float(ws.cell(row=row_idx, column=p_col).value) or 0.0
                withhold = _to_float(ws.cell(row=row_idx, column=w_col).value) or 0.0
                actual = _to_float(ws.cell(row=row_idx, column=a_col).value) or 0.0
                if abs(payable - (withhold + actual)) > 1e-6:
                    for col_idx in range(1, ws.max_column + 1):
                        ws.cell(row=row_idx, column=col_idx).fill = mismatch_fill

    # 一级/二级科目下拉（联动）
    try:
        from .category_dropdowns import apply_category_dropdowns
        from .config import Paths as _Paths

        data_dir = Path(classify_data) if classify_data else _Paths().classify_data
        apply_category_dropdowns(wb, data_dir=data_dir)
    except Exception as exc:
        print(f"  ⚠️ 分类下拉写入失败: {exc}")

    wb.save(output_path)


def _cleanup_stale_project_files(std_dir: Path, keep_relpaths: set[str]) -> None:
    """
    删除本次未产出的项目文件（含子目录）。
    keep_relpaths: 相对 std_dir 的路径，如「最新月份有支出/编码_名称_财务付款数据.xlsx」
    """
    std_dir = Path(std_dir)
    for f in std_dir.rglob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        try:
            rel = f.relative_to(std_dir).as_posix()
        except ValueError:
            continue
        if rel in keep_relpaths:
            continue
        # 典型项目文件 / 新旧命名 / 两个归类子目录下的文件
        parent_name = f.parent.name
        in_bucket = parent_name in {STD_DIR_HAS_SPEND, STD_DIR_NO_SPEND}
        looks_project = bool(
            re.match(
                r"^(THCC|LSBM|XCY|未识别编码|需手动核对).*\.xlsx$",
                f.name,
                re.I,
            )
        ) or f.name.endswith(f"_{STD_FILE_TAG}.xlsx")
        if in_bucket or looks_project or f.parent == std_dir:
            try:
                f.unlink()
                print(f"  🗑️ 移除过期文件: {rel}")
            except OSError:
                pass


def _process_group_std(
    *,
    group_name: str,
    group_raw_dir: Path,
    group_std_dir: Path,
    mapper,
    classify_data: Path,
    cleanup_stale: bool = True,
    prompt_project: bool = False,
) -> List[str]:
    """
    旧路径：组内 Excel 合并后当场分类再拆分。
    月结不再调用。请走 raw_std → recognize → classify → std。
    """
    group_std_dir = Path(group_std_dir)
    group_raw_dir = Path(group_raw_dir)
    group_std_dir.mkdir(parents=True, exist_ok=True)
    has_spend_dir = group_std_dir / STD_DIR_HAS_SPEND
    no_spend_dir = group_std_dir / STD_DIR_NO_SPEND
    has_spend_dir.mkdir(parents=True, exist_ok=True)
    no_spend_dir.mkdir(parents=True, exist_ok=True)

    all_files = iter_excel_files(group_raw_dir)
    if not all_files:
        print(f"  [{group_name}] 未发现 Excel，跳过")
        return []

    # 组内仅扫本组目录；若 raw 即组目录本身，时间标签相对组根
    time_grouped: Dict[str, List[Path]] = {}
    for fp in all_files:
        tag = infer_time_tag_from_path(fp, group_raw_dir)
        time_grouped.setdefault(tag, []).append(fp)

    parts: List[pd.DataFrame] = []
    for tag in sorted(time_grouped.keys()):
        print(f"\n  --- 时间批次: {tag} ---")
        for file_path in time_grouped[tag]:
            print(f"  处理: {file_path.name}")
            try:
                sheets = pd.ExcelFile(file_path).sheet_names
            except Exception as exc:
                print(f"    ⚠️ 打开失败: {exc}")
                continue
            for sheet in sheets:
                part = process_single_sheet(file_path, sheet)
                if part is not None:
                    parts.append(part)

    if not parts:
        print(f"  [{group_name}] 未提取到有效数据")
        return []

    final_df = pd.concat(parts, ignore_index=True)
    final_df = enrich_project_fields(final_df, prompt=prompt_project, interactive=True)
    final_df = append_subject_levels(final_df, mapper)

    latest_ym = detect_global_latest_spend_month(final_df)
    print(f"\n  📅 [{group_name}] 组内最新有支出月份: {latest_ym or '（未识别）'}")

    outputs: List[str] = []
    keep_relpaths: set[str] = set()
    n_has = n_no = 0

    for project_name, project_df in final_df.groupby("项目名称", dropna=False, sort=True):
        name = project_name if not isinstance(project_name, tuple) else project_name[0]
        name = "" if name is None else str(name)
        code = ""
        if "项目编码" in project_df.columns:
            codes = project_df["项目编码"].dropna().astype(str).str.strip()
            codes = codes[codes != ""]
            if not codes.empty:
                code = codes.iloc[0]

        out_name = build_std_project_filename(code, name)
        has_spend = project_has_spend_in_month(project_df, latest_ym)
        bucket = STD_DIR_HAS_SPEND if has_spend else STD_DIR_NO_SPEND
        bucket_dir = has_spend_dir if has_spend else no_spend_dir
        out_path = bucket_dir / out_name
        rel = f"{bucket}/{out_name}"
        keep_relpaths.add(rel)

        wan_df = build_detail_wan_sheet(project_df)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            wan_df.to_excel(writer, sheet_name="明细数据_万泰铢", index=False)
            project_df.to_excel(writer, sheet_name="明细数据", index=False)
        apply_output_format(out_path, classify_data=classify_data)
        outputs.append(str(out_path))
        if has_spend:
            n_has += 1
            tag = "有支出"
        else:
            n_no += 1
            tag = "未支出"
        print(f"  ✅ [{tag}] {group_name}/{bucket}/{out_name}  ({len(project_df)} 行)")

    if cleanup_stale:
        _cleanup_stale_project_files(group_std_dir, keep_relpaths)

    print(
        f"\n  ✅ 分组 [{group_name}] 完成：{len(outputs)} 个项目"
        f"（{STD_DIR_HAS_SPEND} {n_has} / {STD_DIR_NO_SPEND} {n_no}），"
        f"{len(final_df)} 行 → {group_std_dir}"
    )
    if latest_ym:
        print(f"     归类口径：最新月份 = {latest_ym}（预扣税额+实付金额 ≠ 0）")
    return outputs


def _cleanup_legacy_flat_std(std_dir: Path, active_groups: set[str]) -> None:
    """
    迁移后清理旧版「B_std 根下直接有支出/未支出」结构，
    以及已不在 raw 中的分组目录。
    """
    std_dir = Path(std_dir)
    if not std_dir.is_dir():
        return
    for child in list(std_dir.iterdir()):
        if not child.is_dir():
            continue
        if child.name in {STD_DIR_HAS_SPEND, STD_DIR_NO_SPEND}:
            # 旧扁平结构
            try:
                import shutil

                shutil.rmtree(child)
                print(f"  🗑️ 移除旧版扁平目录: {child.name}")
            except OSError as exc:
                print(f"  ⚠️ 无法移除 {child}: {exc}")
            continue
        if child.name.startswith("."):
            continue
        if active_groups and child.name not in active_groups:
            # 仅删除看起来是过期业务组（内含 std 项目文件）
            try:
                has_std = any(child.rglob(f"*_{STD_FILE_TAG}.xlsx"))
            except OSError:
                has_std = False
            if has_std:
                try:
                    import shutil

                    shutil.rmtree(child)
                    print(f"  🗑️ 移除过期分组目录: {child.name}")
                except OSError as exc:
                    print(f"  ⚠️ 无法移除 {child}: {exc}")


def _split_detail_df_to_std(
    final_df: pd.DataFrame,
    *,
    group_name: str,
    group_std_dir: Path,
    classify_data: Path,
    cleanup_stale: bool = True,
    light_format: bool = True,
) -> List[str]:
    """已分类明细 → 按项目拆分到 B_std/{组}/有支出|未支出/。"""
    group_std_dir = Path(group_std_dir)
    group_std_dir.mkdir(parents=True, exist_ok=True)
    has_spend_dir = group_std_dir / STD_DIR_HAS_SPEND
    no_spend_dir = group_std_dir / STD_DIR_NO_SPEND
    has_spend_dir.mkdir(parents=True, exist_ok=True)
    no_spend_dir.mkdir(parents=True, exist_ok=True)

    if final_df is None or final_df.empty:
        print(f"  [{group_name}] 明细为空，跳过拆分")
        return []

    latest_ym = detect_global_latest_spend_month(final_df)
    print(f"  📅 [{group_name}] 组内最新有支出月份: {latest_ym or '（未识别）'}")

    outputs: List[str] = []
    keep_relpaths: set[str] = set()
    n_has = n_no = 0

    for project_name, project_df in final_df.groupby("项目名称", dropna=False, sort=True):
        name = project_name if not isinstance(project_name, tuple) else project_name[0]
        name = "" if name is None else str(name)
        code = ""
        if "项目编码" in project_df.columns:
            codes = project_df["项目编码"].dropna().astype(str).str.strip()
            codes = codes[(codes != "") & (codes.str.lower() != "nan")]
            if not codes.empty:
                code = codes.iloc[0]

        out_name = build_std_project_filename(code, name)
        has_spend = project_has_spend_in_month(project_df, latest_ym)
        bucket = STD_DIR_HAS_SPEND if has_spend else STD_DIR_NO_SPEND
        bucket_dir = has_spend_dir if has_spend else no_spend_dir
        out_path = bucket_dir / out_name
        rel = f"{bucket}/{out_name}"
        keep_relpaths.add(rel)

        wan_df = build_detail_wan_sheet(project_df)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            wan_df.to_excel(writer, sheet_name="明细数据_万泰铢", index=False)
            project_df.to_excel(writer, sheet_name="明细数据", index=False)
        if light_format:
            # 轻量：仅金额格式，避免每个项目文件全量下拉/样式（提速）
            try:
                apply_output_format(out_path, classify_data=None)
            except Exception:
                pass
        else:
            apply_output_format(out_path, classify_data=classify_data)
        outputs.append(str(out_path))
        if has_spend:
            n_has += 1
            tag = "有支出"
        else:
            n_no += 1
            tag = "未支出"
        print(f"  ✅ [{tag}] {group_name}/{bucket}/{out_name}  ({len(project_df)} 行)")

    if cleanup_stale:
        _cleanup_stale_project_files(group_std_dir, keep_relpaths)

    print(
        f"\n  ✅ 分组 [{group_name}] 拆分完成：{len(outputs)} 个项目"
        f"（{STD_DIR_HAS_SPEND} {n_has} / {STD_DIR_NO_SPEND} {n_no}），"
        f"{len(final_df)} 行 → {group_std_dir}"
    )
    return outputs


def run_std_from_classified(
    paths: Paths, *, cleanup_stale: bool = True
) -> List[str]:
    """
    快路径：A2_classified 已分类明细 → B_std 按项目拆分。
    不再重复读 A_raw / 不再重复分类。
    """
    from .classify_raw_std import list_classified_groups, load_classified_detail

    std_dir = Path(paths.std_dir)
    std_dir.mkdir(parents=True, exist_ok=True)
    groups = list_classified_groups(paths)
    if not groups:
        print("未找到 A2_classified，无法快路径拆分 std")
        return []

    print(f"std 快路径（classified → 按项目拆分）: {', '.join(g for g, _ in groups)}")
    all_outputs: List[str] = []
    for group_name, class_xlsx in groups:
        print(f"\n{'=' * 64}")
        print(f"  分组 std · {group_name} ← {class_xlsx.name}")
        print(f"{'=' * 64}")
        detail = load_classified_detail(class_xlsx)
        outs = _split_detail_df_to_std(
            detail,
            group_name=group_name,
            group_std_dir=paths.group_std_dir(group_name),
            classify_data=paths.classify_data,
            cleanup_stale=cleanup_stale,
            light_format=True,
        )
        all_outputs.extend(outs)

    if cleanup_stale:
        _cleanup_legacy_flat_std(std_dir, {g for g, _ in groups})

    print(
        f"\n✅ std 快路径完成：{len(groups)} 个分组，共 {len(all_outputs)} 个项目文件 → {std_dir}"
    )
    return all_outputs


def run_std(
    paths: Paths,
    *,
    cleanup_stale: bool = True,
    prompt_project: bool = False,
) -> List[str]:
    """只拆 A2_classified。分类在 classify 步骤完成。"""
    from .classify_raw_std import list_classified_groups

    if prompt_project:
        print("std 不再匹配项目、不再分类。项目名称在 raw_std 里匹配。")
    if not list_classified_groups(paths):
        raise FileNotFoundError(
            f"未找到 A2 分类结果: {paths.classified_dir}。请先运行 classify。"
        )
    return run_std_from_classified(paths, cleanup_stale=cleanup_stale)
