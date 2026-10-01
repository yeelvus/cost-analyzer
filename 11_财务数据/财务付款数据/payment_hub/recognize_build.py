# -*- coding: utf-8 -*-
"""步骤 recognize：给 A1 明细补上业务摘要、项目提示。

不分类、不改项目名称列。分类步骤只读这里写出的业务摘要。
"""

from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd

from .classify_raw_std import list_raw_std_groups
from .config import Paths
from .pkg import ensure_code_root
from .raw_std_build import apply_raw_std_format


def annotate_business_columns(df: pd.DataFrame) -> pd.DataFrame:
    """写回业务摘要、项目提示，并紧挨在「摘要」后面。"""
    ensure_code_root()
    from summary_recognize import recognize_summary

    out = df.copy()
    n = len(out)
    summaries = out["摘要"].tolist() if "摘要" in out.columns else [""] * n
    projects = out["项目名称"].tolist() if "项目名称" in out.columns else [""] * n
    businesses: List[str] = []
    hints: List[str] = []
    for summary, project in zip(summaries, projects):
        rec = recognize_summary(summary, project)
        businesses.append(rec.business_summary)
        hints.append(rec.project_hint)
    out["业务摘要"] = businesses
    out["项目提示"] = hints

    cols = [c for c in out.columns if c not in {"业务摘要", "项目提示"}]
    if "摘要" in cols:
        i = cols.index("摘要") + 1
        cols = cols[:i] + ["业务摘要", "项目提示"] + cols[i:]
    else:
        cols = ["业务摘要", "项目提示"] + cols
    return out[cols]


def _recognize_one(path: Path) -> int:
    path = Path(path)
    xl = pd.ExcelFile(path)
    sheet_names = list(xl.sheet_names)
    sheets = {name: pd.read_excel(xl, sheet_name=name) for name in sheet_names}
    xl.close()
    detail = sheets.get("明细")
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
    """A1_raw_std/{组}_财务付款_raw_std.xlsx 的明细补上两列，原地写回。"""
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
