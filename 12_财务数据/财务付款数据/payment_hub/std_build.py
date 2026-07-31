# -*- coding: utf-8 -*-
"""
步骤 std：原始付款 → 标准项目明细（原 3_）

优化点：
- 直接使用 learn_classify，不依赖旧 3_ 脚本
- 始终以最新 raw 数据重写项目文件（金额变更可刷新）
- 分类附带供应商特征
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
    infer_time_tag_from_path,
    iter_excel_files,
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


def enrich_project_fields(df: pd.DataFrame) -> pd.DataFrame:
    matcher = get_project_matcher_module()
    db_path = project_match_db_path()
    db = matcher.load_or_init_db(db_path, [])
    raw_series = df["项目名称"].fillna("").astype(str).str.strip()
    unique_names = sorted({n for n in raw_series.tolist() if n and n.lower() != "nan"})

    code_map: Dict[str, str] = {}
    name_map: Dict[str, str] = {}
    unmatched: List[str] = []

    for raw_name in unique_names:
        hit = matcher.match_project(raw_name, db)
        if hit:
            code_map[raw_name] = hit.get("project_code", "")
            name_map[raw_name] = hit.get("project_name", raw_name)
        else:
            matcher.mark_unmatched(db, raw_name)
            code_map[raw_name] = ""
            name_map[raw_name] = raw_name
            unmatched.append(raw_name)

    if unmatched:
        print(f"  ⚠️ 未匹配项目 {len(unmatched)} 个，弹出手动匹配…")
        matcher.prompt_unmatched_project_mappings(
            unmatched, db, window_title="财务付款数据项目匹配"
        )
        for raw_name in unmatched:
            hit = matcher.match_project(raw_name, db)
            if hit:
                code_map[raw_name] = hit.get("project_code", "")
                name_map[raw_name] = hit.get("project_name", raw_name)

    enriched = df.copy()
    enriched["项目编码"] = raw_series.map(code_map).fillna("")
    enriched["项目名称"] = raw_series.map(name_map).fillna(raw_series)
    enriched = enriched[["项目编码"] + TARGET_COLUMNS]
    matcher.save_db(db_path, db)
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

    triples = list(zip(subjects, summaries, vendors))
    unique_triples = list(dict.fromkeys(triples))
    cache = {}

    engine = getattr(mapper, "engine", None)
    use_full = hasattr(mapper, "classify_full")

    if engine is not None and hasattr(engine, "classify_unique_map"):
        cache = engine.classify_unique_map(unique_triples)
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
    for t in triples:
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

    if unique_triples:
        print(
            f"  分类缓存: {len(unique_triples)} 种组合 → {n} 行"
            f"（复用率 {(1 - len(unique_triples) / max(n, 1)) * 100:.0f}%）"
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


def _cleanup_stale_project_files(std_dir: Path, keep: set[str]) -> None:
    """删除本次未产出的旧项目文件（避免改名后残留）。"""
    for f in std_dir.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        if f.name not in keep:
            # 仅清理典型命名；保留需手动核对等特殊文件若本次未覆盖
            if re.match(r"^(THCC|LSBM|未识别编码|需手动核对)", f.name):
                try:
                    f.unlink()
                    print(f"  🗑️ 移除过期文件: {f.name}")
                except OSError:
                    pass


def run_std(paths: Paths, *, cleanup_stale: bool = False) -> List[str]:
    raw_dir = Path(paths.raw_dir)
    std_dir = Path(paths.std_dir)
    if not raw_dir.is_dir():
        raise FileNotFoundError(f"原始目录不存在: {raw_dir}")
    std_dir.mkdir(parents=True, exist_ok=True)

    mapper = get_subject_mapper(paths.classify_data)
    all_files = iter_excel_files(raw_dir)
    if not all_files:
        print("未发现 Excel，跳过 std")
        return []

    grouped: Dict[str, List[Path]] = {}
    for fp in all_files:
        tag = infer_time_tag_from_path(fp, raw_dir)
        grouped.setdefault(tag, []).append(fp)

    parts: List[pd.DataFrame] = []
    for tag in sorted(grouped.keys()):
        print(f"\n{'=' * 60}\n处理时间批次: {tag}")
        for file_path in grouped[tag]:
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
        print("未提取到有效数据")
        return []

    final_df = pd.concat(parts, ignore_index=True)
    final_df = enrich_project_fields(final_df)
    final_df = append_subject_levels(final_df, mapper)

    outputs: List[str] = []
    keep_names: set[str] = set()

    for project_name, project_df in final_df.groupby("项目名称", dropna=False, sort=True):
        name = project_name if not isinstance(project_name, tuple) else project_name[0]
        name = "" if name is None else str(name)
        code = ""
        if "项目编码" in project_df.columns:
            codes = project_df["项目编码"].dropna().astype(str).str.strip()
            codes = codes[codes != ""]
            if not codes.empty:
                code = codes.iloc[0]
        code_part = safe_filename(code) if code else "未识别编码"
        name_part = safe_filename(name)
        last_date = format_last_payment_date(project_df)
        out_name = f"{code_part}_{name_part}_({last_date}).xlsx"
        out_path = std_dir / out_name
        keep_names.add(out_name)

        wan_df = build_detail_wan_sheet(project_df)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            wan_df.to_excel(writer, sheet_name="明细数据_万泰铢", index=False)
            project_df.to_excel(writer, sheet_name="明细数据", index=False)
        apply_output_format(out_path, classify_data=paths.classify_data)
        outputs.append(str(out_path))
        print(f"  ✅ {out_name}  ({len(project_df)} 行)")

    if cleanup_stale:
        _cleanup_stale_project_files(std_dir, keep_names)

    print(f"\n✅ std 完成：{len(outputs)} 个项目，{len(final_df)} 行 → {std_dir}")
    return outputs
