# -*- coding: utf-8 -*-
"""
步骤 classify：读取 A1_财务付款_raw_std 合并表，
按 摘要(+科目/供应商) 学习分类 → 一级/二级/三级科目，
写出 A2_财务付款_classified/{组名}/{组名}_classified.xlsx

设计目标：一份组内明细只读一次、唯一三元组只分类一次，供后续 sum / std 复用。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .config import Paths
from .pkg import ensure_code_root
from .utils import has_actual_payment, list_raw_groups, safe_text, to_numeric_series

AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]
DEFAULT_L1 = "6_管理费"
DEFAULT_L2 = "待复核"

CLASSIFY_COLS = [
    "一级科目",
    "二级科目",
    "三级科目",
    "分类ID",
    "命中层",
    "置信度",
]


def _cell(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() in {"nan", "none", "nat"} else s


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
    对明细附加一级/二级/三级科目（唯一三元组缓存）。

    仅对「实付金额非空」的行做科目分类；实付为空的行不调用学习引擎，
    科目列留空，命中层标记为 skip_no_payment。
    """
    out = df.copy()
    n = len(out)
    pay_mask = (
        out["实付金额"].map(has_actual_payment)
        if "实付金额" in out.columns
        else pd.Series([False] * n, index=out.index)
    )
    n_pay = int(pay_mask.sum())
    n_skip = n - n_pay

    subjects = (
        out["科目"].map(_cell).tolist() if "科目" in out.columns else [""] * n
    )
    summaries = (
        out["摘要"].map(_cell).tolist() if "摘要" in out.columns else [""] * n
    )
    if "供应商名称" in out.columns:
        vendors = out["供应商名称"].map(_cell).tolist()
    elif "供应商" in out.columns:
        vendors = out["供应商"].map(_cell).tolist()
    else:
        vendors = [""] * n

    # 只对有实付金额的行建唯一三元组并分类
    pay_indices = [i for i, ok in enumerate(pay_mask.tolist()) if ok]
    triples_pay = [(subjects[i], summaries[i], vendors[i]) for i in pay_indices]
    unique = list(dict.fromkeys(triples_pay))
    print(
        f"  摘要分类: 总 {n} 行，有实付 {n_pay} 行（跳过空实付 {n_skip}），"
        f"唯一 (科目,摘要,供应商) = {len(unique)}"
    )

    cache = engine.classify_unique_map(unique) if unique else {}

    l1s = [""] * n
    l2s = [""] * n
    l3s = [""] * n
    ids = [""] * n
    layers = [""] * n
    confs: List = [None] * n

    for i in range(n):
        if not pay_mask.iloc[i]:
            layers[i] = "skip_no_payment"
            confs[i] = None
            continue
        t = (subjects[i], summaries[i], vendors[i])
        r = cache.get(t)
        if r is None:
            l1s[i] = DEFAULT_L1
            l2s[i] = DEFAULT_L2
            l3s[i] = ""
            ids[i] = ""
            layers[i] = "error"
            confs[i] = 0.0
            continue
        l1s[i] = str(getattr(r, "l1_name", "") or DEFAULT_L1)
        l2s[i] = str(getattr(r, "l2_name", "") or DEFAULT_L2)
        l3s[i] = str(getattr(r, "l3_name", "") or "")
        leaf = getattr(r, "l3_id", "") or getattr(r, "l2_id", "") or ""
        ids[i] = str(leaf or "")
        layers[i] = str(getattr(r, "layer", "") or "")
        confs[i] = getattr(r, "confidence", None)

    out["一级科目"] = l1s
    out["二级科目"] = l2s
    out["三级科目"] = l3s
    out["分类ID"] = ids
    out["命中层"] = layers
    out["置信度"] = confs

    # 科目列插到摘要前，便于阅读
    prefer = CLASSIFY_COLS
    cols = [c for c in out.columns if c not in prefer]
    if "摘要" in cols:
        i = cols.index("摘要")
        cols = cols[:i] + prefer + cols[i:]
    else:
        cols = cols + prefer
    return out[cols]


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
