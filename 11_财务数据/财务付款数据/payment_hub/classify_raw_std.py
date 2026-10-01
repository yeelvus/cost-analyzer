# -*- coding: utf-8 -*-
"""
步骤 classify：读取 A1_财务付款_raw_std 合并表，
用业务摘要（没有则现场识别）+ 科目 + 供应商 → 一级/二级/三级科目，
写出 A2_财务付款_classified/{组名}/{组名}_classified.xlsx

设计目标：一份组内明细只读一次、唯一三元组只分类一次，供后续 sum / std 复用。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .config import Paths
from .pkg import ensure_code_root
from .utils import list_raw_groups, to_numeric_series

AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]


def list_classified_groups(paths: Paths) -> List[Tuple[str, Path]]:
    """已分类分组：(组名, classified xlsx)。"""
    root = Path(paths.classified_dir)
    if not root.is_dir():
        return []
    out: List[Tuple[str, Path]] = []
    for child in sorted(root.iterdir(), key=lambda p: p.name):
        if not child.is_dir() or child.name.startswith("."):
            continue
        xlsx = paths.group_classified_xlsx(child.name)
        if xlsx.is_file():
            out.append((child.name, xlsx))
    return out


def _existing_raw_std_xlsx(paths: Paths, name: str) -> Optional[Path]:
    """扁平路径优先，兼容旧的 A1/{组名}/{组名}_财务付款_raw_std.xlsx。"""
    flat = paths.group_raw_std_xlsx(name)
    if flat.is_file():
        return flat
    nested = Path(paths.raw_std_dir) / name / f"{name}_财务付款_raw_std.xlsx"
    if nested.is_file():
        return nested
    return None


def list_raw_std_groups(paths: Paths) -> List[Tuple[str, Path]]:
    """A1 raw_std 分组：(组名, raw_std xlsx)。优先 A_raw 一级目录名。"""
    groups = list_raw_groups(paths.raw_dir)
    result: List[Tuple[str, Path]] = []
    if groups:
        for name, _ in groups:
            p = _existing_raw_std_xlsx(paths, name)
            if p is not None:
                result.append((name, p))
            else:
                print(f"  ⚠️ 缺少 raw_std: {paths.group_raw_std_xlsx(name)}")
        return result

    # 回退：扫 A1 根目录扁平文件，再扫旧的组名子目录
    root = Path(paths.raw_std_dir)
    if not root.is_dir():
        return []
    seen: set[str] = set()
    for child in sorted(root.iterdir(), key=lambda p: p.name):
        if child.is_file() and child.name.endswith("_财务付款_raw_std.xlsx"):
            name = child.name[: -len("_财务付款_raw_std.xlsx")]
            if name and name not in seen:
                result.append((name, child))
                seen.add(name)
            continue
        if not child.is_dir():
            continue
        p = _existing_raw_std_xlsx(paths, child.name)
        if p is not None and child.name not in seen:
            result.append((child.name, p))
            seen.add(child.name)
    return result


def load_raw_std_detail(path: Path) -> pd.DataFrame:
    path = Path(path)
    df = pd.read_excel(path, sheet_name="明细")
    for col in AMOUNT_COLUMNS:
        if col in df.columns:
            df[col] = to_numeric_series(df[col])
    return df


def classify_dataframe(df: pd.DataFrame, engine) -> pd.DataFrame:
    """
    月结填科目：业务摘要 + 科目 + 供应商 → 一二三级。

    只读科目记忆，不写学习库。仅实付金额非空的行分类；
    实付为空的行科目留空，命中层为 skip_no_payment。
    """
    ensure_code_root()
    from subject_fill import fill_subjects

    return fill_subjects(df, engine=engine, only_paid=True, verbose=True)


def _write_classified(df: pd.DataFrame, out_path: Path) -> None:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # 导出用轻量格式，提速
    export = df.copy()
    for col in ("年", "月"):
        if col in export.columns:
            nums = pd.to_numeric(export[col], errors="coerce")
            export[col] = [None if pd.isna(v) else int(v) for v in nums]
    if "日期" in export.columns:
        parsed = pd.to_datetime(export["日期"], errors="coerce")
        export["日期"] = [
            parsed.iloc[i].to_pydatetime() if pd.notna(parsed.iloc[i]) else export["日期"].iloc[i]
            for i in range(len(export))
        ]

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="明细")
        # 分类概况
        if "一级科目" in df.columns and "二级科目" in df.columns:
            g = (
                df.groupby(["一级科目", "二级科目", "三级科目"], dropna=False)
                .size()
                .reset_index(name="行数")
                .sort_values("行数", ascending=False)
            )
            g.to_excel(writer, index=False, sheet_name="分类统计")
        if "命中层" in df.columns:
            layer = (
                df.groupby("命中层", dropna=False)
                .size()
                .reset_index(name="行数")
                .sort_values("行数", ascending=False)
            )
            layer.to_excel(writer, index=False, sheet_name="命中层统计")


def classify_one_group(
    group_name: str,
    raw_std_path: Path,
    out_path: Path,
    engine,
) -> Optional[Path]:
    print(f"  读取: {raw_std_path}")
    df = load_raw_std_detail(raw_std_path)
    if df.empty:
        print(f"  [{group_name}] 明细为空，跳过")
        return None
    print(f"  [{group_name}] {len(df)} 行，开始摘要分类…")
    classified = classify_dataframe(df, engine)
    _write_classified(classified, out_path)
    review_n = int(
        (classified["命中层"].astype(str).isin(["pending", "low_confidence", "conflict"])).sum()
        if "命中层" in classified.columns
        else 0
    )
    print(
        f"  ✅ [{group_name}] 分类完成 {len(classified)} 行 → {out_path.name}"
        f"（待复核倾向 {review_n} 行）"
    )
    return out_path


def run_classify_raw_std(paths: Paths) -> List[str]:
    """
    A1_raw_std/{组}_财务付款_raw_std.xlsx
      → A2_classified/{组}/*_classified.xlsx
    """
    ensure_code_root()
    from learn_classify.engine import ClassifyEngine

    groups = list_raw_std_groups(paths)
    if not groups:
        print("未找到 A1_raw_std 合并文件，请先运行 raw-std")
        return []

    print(f"classify 分组: {', '.join(g for g, _ in groups)}")
    print(f"输出: {paths.classified_dir}/{{组名}}/")
    engine = ClassifyEngine.from_data_dir(paths.classify_data)

    written: List[str] = []
    for group_name, raw_path in groups:
        print(f"\n{'=' * 64}")
        print(f"  classify · {group_name}")
        print(f"{'=' * 64}")
        out = paths.group_classified_xlsx(group_name)
        result = classify_one_group(group_name, raw_path, out, engine)
        if result is not None:
            written.append(str(result))

    print(f"\n✅ classify 完成：{len(written)} 个文件 → {paths.classified_dir}")
    for p in written:
        print(f"   · {p}")
    return written


def load_classified_detail(path: Path) -> pd.DataFrame:
    path = Path(path)
    df = pd.read_excel(path, sheet_name="明细")
    for col in AMOUNT_COLUMNS:
        if col in df.columns:
            df[col] = to_numeric_series(df[col])
    return df
