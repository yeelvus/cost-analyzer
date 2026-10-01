# -*- coding: utf-8 -*-
"""把摘要列对应成一级 / 二级 / 三级科目。

只读科目记忆，返回带科目列的表，或另存一份工作簿。
不写 memory、taxonomy，也不写月结用的 A1 / A2 / B_std / 汇总。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, List, Optional, Sequence

import pandas as pd

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

RECOGNIZE_COLS = ["业务摘要", "项目提示"]
CLASSIFY_COLS = ["一级科目", "二级科目", "三级科目", "分类ID", "命中层", "置信度"]

DEFAULT_L1 = "6_管理费"
DEFAULT_L2 = "待复核"
_STAT_SHEETS = ("分类统计", "命中层统计")


def default_data_dir() -> Path:
    return _ROOT / "learn_classify" / "data"


def default_output_path(src: Path) -> Path:
    src = Path(src)
    return src.with_name(f"{src.stem}_科目{src.suffix}")


def _cell(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() in {"nan", "none", "nat"} else s


def _has_actual_payment(value: Any) -> bool:
    """非空、能当成数字、且绝对值大于 0。与月结里「有实付才分类」同一规则。"""
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return abs(float(value)) > 1e-12
    text = str(value).strip().replace(",", "").replace("，", "").replace(" ", "")
    if not text or text.lower() in {"nan", "none", "nat", "-"}:
        return False
    try:
        return abs(float(text)) > 1e-12
    except (TypeError, ValueError):
        return False


def _column(df: pd.DataFrame, name: Optional[str], n: int) -> List[str]:
    if name and name in df.columns:
        return df[name].map(_cell).tolist()
    return [""] * n


def _pick_vendor_col(df: pd.DataFrame, vendor_col: Optional[str]) -> Optional[str]:
    if vendor_col is not None:
        return vendor_col or None
    if "供应商名称" in df.columns:
        return "供应商名称"
    if "供应商" in df.columns:
        return "供应商"
    return None


def load_engine(engine: Any = None, data_dir: Optional[Path] = None):
    """只打开科目记忆来分类。调用方要写记忆时不要走这里。"""
    if engine is not None:
        return engine
    from learn_classify.engine import ClassifyEngine

    return ClassifyEngine.from_data_dir(Path(data_dir) if data_dir else default_data_dir())


def fill_subjects(
    df: pd.DataFrame,
    *,
    summary_col: str = "摘要",
    project_col: str = "项目名称",
    subject_col: str = "科目",
    vendor_col: Optional[str] = None,
    business_col: str = "业务摘要",
    amount_col: str = "实付金额",
    only_paid: bool = False,
    engine: Any = None,
    data_dir: Optional[Path] = None,
    verbose: bool = False,
) -> pd.DataFrame:
    """
    给每一行填上业务摘要、项目提示和一二三级科目。返回新表。

    only_paid=True 时只分类实付金额非空的行（月结用）；
    其余行科目留空，命中层为 skip_no_payment。
    不读不写任何文件；engine 为空时只读 data_dir 里的科目记忆。
    """
    from summary_recognize import resolve_business

    out = df.copy()
    n = len(out)
    if only_paid and amount_col in out.columns:
        pay_flags = out[amount_col].map(_has_actual_payment).tolist()
    elif only_paid:
        pay_flags = [False] * n
    else:
        pay_flags = [True] * n

    subjects = _column(out, subject_col, n)
    summaries = _column(out, summary_col, n)
    projects = _column(out, project_col, n)
    given = _column(out, business_col, n)
    vendors = _column(out, _pick_vendor_col(out, vendor_col), n)

    businesses: List[str] = []
    hints: List[str] = []
    for summary, project, business in zip(summaries, projects, given):
        rec = resolve_business(summary, project, business)
        businesses.append(rec.business_summary)
        hints.append(rec.project_hint)

    out["业务摘要"] = businesses
    out["项目提示"] = hints

    triples = [
        (subjects[i], businesses[i], vendors[i])
        for i, ok in enumerate(pay_flags)
        if ok
    ]
    unique = list(dict.fromkeys(triples))
    if verbose:
        if only_paid:
            n_pay = sum(1 for ok in pay_flags if ok)
            print(
                f"  摘要分类: 总 {n} 行，有实付 {n_pay} 行（跳过空实付 {n - n_pay}），"
                f"唯一 (科目,业务摘要,供应商) = {len(unique)}"
            )
        else:
            print(f"  摘要分类: {n} 行，唯一 (科目,业务摘要,供应商) = {len(unique)}")

    eng = load_engine(engine, data_dir) if unique else engine
    cache = eng.classify_unique_map(unique) if unique and eng is not None else {}

    l1s = [""] * n
    l2s = [""] * n
    l3s = [""] * n
    ids = [""] * n
    layers = [""] * n
    confs: List[Any] = [None] * n

    for i in range(n):
        if not pay_flags[i]:
            layers[i] = "skip_no_payment"
            continue
        key = (subjects[i], businesses[i], vendors[i])
        result = cache.get(key)
        if result is None and eng is not None and hasattr(eng, "_triple_key"):
            result = cache.get(eng._triple_key(*key))
        if result is None:
            l1s[i] = DEFAULT_L1
            l2s[i] = DEFAULT_L2
            layers[i] = "error"
            confs[i] = 0.0
            continue
        l1s[i] = str(getattr(result, "l1_name", "") or DEFAULT_L1)
        l2s[i] = str(getattr(result, "l2_name", "") or DEFAULT_L2)
        l3s[i] = str(getattr(result, "l3_name", "") or "")
        leaf = getattr(result, "l3_id", "") or getattr(result, "l2_id", "") or ""
        ids[i] = str(leaf or "")
        layers[i] = str(getattr(result, "layer", "") or "")
        confs[i] = getattr(result, "confidence", None)

    out["一级科目"] = l1s
    out["二级科目"] = l2s
    out["三级科目"] = l3s
    out["分类ID"] = ids
    out["命中层"] = layers
    out["置信度"] = confs

    block = RECOGNIZE_COLS + CLASSIFY_COLS
    cols = [c for c in out.columns if c not in block]
    anchor = summary_col if summary_col in cols else None
    if anchor is None and "摘要" in cols:
        anchor = "摘要"
    if anchor is not None:
        i = cols.index(anchor) + 1
        cols = cols[:i] + block + cols[i:]
    else:
        cols = block + cols
    return out[cols]


def _write_stats(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    if {"一级科目", "二级科目", "三级科目"} <= set(df.columns):
        grouped = (
            df.groupby(["一级科目", "二级科目", "三级科目"], dropna=False)
            .size()
            .reset_index(name="行数")
            .sort_values("行数", ascending=False)
        )
        grouped.to_excel(writer, index=False, sheet_name="分类统计")
    if "命中层" in df.columns:
        layers = (
            df.groupby("命中层", dropna=False)
            .size()
            .reset_index(name="行数")
            .sort_values("行数", ascending=False)
        )
        layers.to_excel(writer, index=False, sheet_name="命中层统计")


def fill_workbook(
    src: Path,
    dest: Optional[Path] = None,
    *,
    sheet: Optional[str] = None,
    header_row: int = 1,
    summary_col: str = "摘要",
    project_col: str = "项目名称",
    subject_col: str = "科目",
    vendor_col: Optional[str] = None,
    business_col: str = "业务摘要",
    amount_col: str = "实付金额",
    only_paid: bool = False,
    inplace: bool = False,
    engine: Any = None,
    data_dir: Optional[Path] = None,
) -> Path:
    """
    读一份 Excel，把指定表的摘要列填成一二三级科目，写到另一个文件。

    默认文件名是「原名_科目.xlsx」。不覆盖原文件，除非 inplace=True。
    """
    src = Path(src).expanduser()
    if not src.is_file():
        raise FileNotFoundError(src)
    if inplace:
        if dest is not None and Path(dest).expanduser().resolve() != src.resolve():
            raise ValueError("inplace 时不要再指定另一个 -o")
        dest_path = src
    else:
        dest_path = Path(dest).expanduser() if dest else default_output_path(src)
        if dest_path.resolve() == src.resolve():
            raise ValueError(
                "输出和原文件是同一个路径。不覆盖原文件时请换一个 -o，"
                f"或省略 -o 以写到 {default_output_path(src).name}"
            )

    header = max(0, int(header_row) - 1)
    book = pd.ExcelFile(src)
    names: Sequence[str] = list(book.sheet_names)
    if not names:
        raise ValueError(f"没有工作表: {src}")
    target = sheet or ("明细" if "明细" in names else names[0])
    if target not in names:
        raise KeyError(f"没有工作表 {target!r}，现有: {list(names)}")

    frames: dict[str, pd.DataFrame] = {}
    for name in names:
        if name in _STAT_SHEETS:
            continue
        frames[name] = pd.read_excel(
            book, sheet_name=name, header=header if name == target else 0
        )
    book.close()
    if summary_col not in frames[target].columns:
        raise KeyError(
            f"工作表 {target!r} 没有列 {summary_col!r}，现有: {list(frames[target].columns)}"
        )

    filled = fill_subjects(
        frames[target],
        summary_col=summary_col,
        project_col=project_col,
        subject_col=subject_col,
        vendor_col=vendor_col,
        business_col=business_col,
        amount_col=amount_col,
        only_paid=only_paid,
        engine=engine,
        data_dir=data_dir,
        verbose=True,
    )
    frames[target] = filled

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest_path.with_name(f".~{dest_path.stem}.tmp{dest_path.suffix}")
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for name, frame in frames.items():
                frame.to_excel(writer, index=False, sheet_name=name[:31])
            _write_stats(writer, filled)
        os.replace(tmp, dest_path)
    except Exception:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        raise
    print(f"已写出科目: {dest_path}")
    return dest_path
