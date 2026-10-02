# -*- coding: utf-8 -*-
"""步骤 recognize：给 A1 明细补上项目提示。

不另写业务摘要，不改「摘要」和「项目名称」。
分类时再从「摘要」现场取出业务描述。
"""

from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd

from .classify_raw_std import list_raw_std_groups
from .config import Paths
from .pkg import ensure_code_root
from .payment_amount import YEAR_MONTH_COL, format_year_month
from .raw_std_build import apply_raw_std_format


def annotate_business_columns(df: pd.DataFrame) -> pd.DataFrame:
    """写回项目提示，紧挨在「摘要」后面。已有的业务摘要列去掉。"""
    ensure_code_root()
    from summary_recognize import recognize_summary

    out = df.copy()
    if "业务摘要" in out.columns:
        out = out.drop(columns=["业务摘要"])
    n = len(out)
    summaries = out["摘要"].tolist() if "摘要" in out.columns else [""] * n
    projects = out["项目名称"].tolist() if "项目名称" in out.columns else [""] * n
    hints: List[str] = []
    for summary, project in zip(summaries, projects):
        hints.append(recognize_summary(summary, project).project_hint)
    out["项目提示"] = hints

    cols = [c for c in out.columns if c != "项目提示"]
    if "摘要" in cols:
        i = cols.index("摘要") + 1
        cols = cols[:i] + ["项目提示"] + cols[i:]
    else:
        cols = ["项目提示"] + cols
    return out[cols]


def _recognize_one(path: Path) -> int:
    path = Path(path)
    xl = pd.ExcelFile(path)
    sheet_names = list(xl.sheet_names)
    sheets = {name: pd.read_excel(xl, sheet_name=name) for name in sheet_names}
    xl.close()
    detail = sheets.get("明细")
    if detail is not None and YEAR_MONTH_COL in detail.columns:
        detail[YEAR_MONTH_COL] = detail[YEAR_MONTH_COL].map(format_year_month)
    if detail is None or detail.empty:
        print(f"  ⚠️ 无明细 sheet: {path.name}")
        return 0
    sheets["明细"] = annotate_business_columns(detail)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name in sheet_names:
            sheets[name].to_excel(writer, index=False, sheet_name=name)
    apply_raw_std_format(path)
    hinted = int((sheets["明细"]["项目提示"].astype(str).str.strip() != "").sum())
    print(f"  ✅ {path.name}: {len(sheets['明细'])} 行，项目提示 {hinted} 行")
    return len(sheets["明细"])


def run_recognize(paths: Paths) -> List[str]:
    """A1_raw_std/{组}_财务付款_raw_std.xlsx 的明细补上项目提示，原地写回。"""
    groups = list_raw_std_groups(paths)
    if not groups:
        print(f"未找到 raw_std: {paths.raw_std_dir}")
        return []
    outs: List[str] = []
    for name, path in groups:
        print(f"  [{name}] {path.name}")
        n = _recognize_one(path)
        if n:
            outs.append(str(path))
    print(f"摘要识别完成：{len(outs)} 个分组")
    return outs
