# -*- coding: utf-8 -*-
"""待复核导出 / 批量重分类 std 文件。"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook

from .engine import ClassifyEngine, ClassifyResult


REVIEW_COLUMNS = [
    "文件",
    "项目名称",
    "科目",
    "摘要",
    "供应商名称",
    "实付金额",
    "一级科目",
    "二级科目",
    "分类ID",
    "命中层",
    "置信度",
    "原因",
    "正确一级科目",
    "正确二级科目",
    "备注",
]


def classify_dataframe(
    df: pd.DataFrame,
    engine: ClassifyEngine,
    *,
    subject_col: str = "科目",
    summary_col: str = "摘要",
    vendor_col: str = "供应商名称",
) -> pd.DataFrame:
    out = df.copy()
    results: List[ClassifyResult] = []
    for _, row in out.iterrows():
        subj = row.get(subject_col) if subject_col in out.columns else None
        summ = row.get(summary_col) if summary_col in out.columns else None
        vend = row.get(vendor_col) if vendor_col in out.columns else None
        results.append(engine.classify(subj, summ, vend))

    out["一级科目"] = [r.l1_name for r in results]
    out["二级科目"] = [r.l2_name for r in results]
    out["分类ID"] = [r.l2_id for r in results]
    out["命中层"] = [r.layer for r in results]
    out["置信度"] = [r.confidence for r in results]
    out["原因"] = [r.reason for r in results]
    out["一级科目ID"] = [r.l1_id for r in results]
    return out


def export_review_pool(
    std_dir: Path,
    engine: ClassifyEngine,
    out_path: Path,
    *,
    conf_threshold: float = 0.55,
    only_review: bool = True,
    max_rows: Optional[int] = None,
) -> Path:
    """从 B_财务付款_std 扫描，导出待复核表（含空的正确列供填写）。"""
    std_dir = Path(std_dir)
    rows: List[Dict[str, Any]] = []

    for f in sorted(std_dir.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        try:
            df = pd.read_excel(f, sheet_name=0)
        except Exception:
            continue
        cols = {str(c).strip(): c for c in df.columns}
        c_subj = cols.get("科目")
        c_summ = cols.get("摘要")
        c_vend = cols.get("供应商名称") or cols.get("供应商")
        c_amt = cols.get("实付金额")
        c_proj = cols.get("项目名称")

        for _, row in df.iterrows():
            subj = "" if not c_subj or pd.isna(row.get(c_subj)) else str(row.get(c_subj)).strip()
            summ = "" if not c_summ or pd.isna(row.get(c_summ)) else str(row.get(c_summ)).strip()
            vend = "" if not c_vend or pd.isna(row.get(c_vend)) else str(row.get(c_vend)).strip()
            r = engine.classify(subj, summ, vend)
            if only_review and not (r.needs_review or r.confidence < conf_threshold or r.layer == "pending"):
                # 也导出「其他」类便于清洗
                if r.l2_name not in {"其他", "待复核", "(空白)"}:
                    continue
            amt = None
            if c_amt is not None:
                try:
                    amt = float(row.get(c_amt) or 0)
                except Exception:
                    amt = row.get(c_amt)
            proj = "" if not c_proj or pd.isna(row.get(c_proj)) else str(row.get(c_proj)).strip()
            rows.append(
                {
                    "文件": f.name,
                    "项目名称": proj,
                    "科目": subj,
                    "摘要": summ,
                    "供应商名称": vend,
                    "实付金额": amt,
                    "一级科目": r.l1_name,
                    "二级科目": r.l2_name,
                    "分类ID": r.l2_id,
                    "命中层": r.layer,
                    "置信度": r.confidence,
                    "原因": r.reason,
                    "正确一级科目": "",
                    "正确二级科目": "",
                    "备注": "",
                }
            )
            if max_rows and len(rows) >= max_rows:
                break
        if max_rows and len(rows) >= max_rows:
            break

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows, columns=REVIEW_COLUMNS).to_excel(out_path, index=False)
    return out_path


def update_excel_classification_columns(
    excel_path: Path,
    engine: ClassifyEngine,
    sheet_names=None,
    subject_col: str = "科目",
    summary_col: str = "摘要",
    vendor_col: str = "供应商名称",
    level1_col: str = "一级科目",
    level2_col: str = "二级科目",
    extra_cols: bool = True,
) -> int:
    """更新已有 Excel 的分类列（保留其它内容）。"""
    wb = load_workbook(excel_path)
    target_sheets = sheet_names or wb.sheetnames
    updated = 0

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_to_col = {}
        for c in range(1, ws.max_column + 1):
            header = ws.cell(row=1, column=c).value
            if header is None:
                continue
            header_to_col[str(header).strip()] = c

        subj_idx = header_to_col.get(subject_col)
        summ_idx = header_to_col.get(summary_col)
        vend_idx = header_to_col.get(vendor_col)
        if subj_idx is None and summ_idx is None:
            continue

        def ensure_col(name: str) -> int:
            if name in header_to_col:
                return header_to_col[name]
            idx = ws.max_column + 1
            ws.cell(row=1, column=idx, value=name)
            header_to_col[name] = idx
            return idx

        lvl1_idx = ensure_col(level1_col)
        lvl2_idx = ensure_col(level2_col)
        id_idx = conf_idx = layer_idx = None
        if extra_cols:
            id_idx = ensure_col("分类ID")
            layer_idx = ensure_col("命中层")
            conf_idx = ensure_col("置信度")

        for r in range(2, ws.max_row + 1):
            subject_value = ws.cell(row=r, column=subj_idx).value if subj_idx else None
            summary_value = ws.cell(row=r, column=summ_idx).value if summ_idx else None
            vendor_value = ws.cell(row=r, column=vend_idx).value if vend_idx else None
            if subject_value is None and summary_value in (None, ""):
                continue
            result = engine.classify(subject_value, summary_value, vendor_value)
            ws.cell(row=r, column=lvl1_idx, value=result.l1_name)
            ws.cell(row=r, column=lvl2_idx, value=result.l2_name)
            if id_idx:
                ws.cell(row=r, column=id_idx, value=result.l2_id)
            if layer_idx:
                ws.cell(row=r, column=layer_idx, value=result.layer)
            if conf_idx:
                ws.cell(row=r, column=conf_idx, value=result.confidence)
            updated += 1

    # 重分类后刷新一二级下拉（与 taxonomy 同步）
    try:
        hub_root = Path(__file__).resolve().parent.parent
        if str(hub_root) not in sys.path:
            sys.path.insert(0, str(hub_root))
        from payment_hub.category_dropdowns import apply_category_dropdowns

        data_dir = Path(__file__).resolve().parent / "data"
        apply_category_dropdowns(wb, data_dir=data_dir)
    except Exception:
        pass

    wb.save(excel_path)
    return updated


def reclassify_std_dir(
    std_dir: Path,
    engine: ClassifyEngine,
    *,
    sheet_names=None,
    extra_cols: bool = True,
) -> Dict[str, int]:
    std_dir = Path(std_dir)
    stats = {"files": 0, "rows": 0}
    for f in sorted(std_dir.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        n = update_excel_classification_columns(
            f,
            engine,
            sheet_names=sheet_names or ["明细数据_万泰铢", "明细数据"],
            extra_cols=extra_cols,
        )
        stats["files"] += 1
        stats["rows"] += n
    return stats
