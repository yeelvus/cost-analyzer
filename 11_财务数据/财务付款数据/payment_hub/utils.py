# -*- coding: utf-8 -*-
"""共用工具。"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, List, Optional, Set

import pandas as pd


def normalize_text(text) -> str:
    if text is None:
        return ""
    s = str(text).strip()
    s = s.replace("\n", "").replace("\r", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[()（）\[\]【】:：_\-]", "", s)
    return s.lower()


def safe_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def has_actual_payment(value) -> bool:
    """
    是否「有实付金额」：非空且可解析为数字，且绝对值 > 0。

    空 / None / NaN / 空白 / 0 / 0.0 → False
    （不进学习池、不做科目分类）
    """
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


def filter_df_has_actual_payment(
    df: pd.DataFrame, amount_col: str = "实付金额"
) -> pd.DataFrame:
    """保留实付金额非空的行；无该列时返回空表（避免误分类）。"""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    if amount_col not in df.columns:
        return df.iloc[0:0].copy()
    mask = df[amount_col].map(has_actual_payment)
    return df.loc[mask].copy()


def safe_filename(name, max_len: int = 80) -> str:
    if name is None:
        return "未命名项目"
    text = str(name).strip()
    if not text:
        return "未命名项目"
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len]


def extract_time_tag(folder_name: str) -> str:
    match = re.search(r"(20\d{6})", str(folder_name))
    if match:
        return match.group(1)
    match = re.search(r"(\d{6,8})", str(folder_name))
    if match:
        return match.group(1)
    return str(folder_name)


def infer_time_tag_from_path(file_path: Path | str, input_root: Path | str) -> str:
    file_path = Path(file_path)
    input_root = Path(input_root)
    try:
        rel_parts = list(file_path.relative_to(input_root).parts)
    except ValueError:
        rel_parts = list(file_path.parts)
    candidates = rel_parts[:-1] + [file_path.stem]
    for text in reversed(candidates):
        tag = str(extract_time_tag(text))
        digits = "".join(ch for ch in tag if ch.isdigit())
        if len(digits) >= 8 and digits.startswith("20"):
            return digits[:8]
        if len(digits) >= 6 and digits.startswith("20"):
            return digits[:6]
    return "未识别时间"


def to_numeric_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace({"": pd.NA, "None": pd.NA, "nan": pd.NA, "NaN": pd.NA})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def iter_excel_files(
    root: Path | str,
    *,
    recursive: bool = True,
    exclude_dir_names: Optional[Set[str]] = None,
) -> List[Path]:
    root = Path(root)
    exclude_dir_names = exclude_dir_names or set()
    files: List[Path] = []
    seen: set = set()
    patterns = ("*.xlsx", "*.xls", "*.xlsm")
    iterator: Iterable[Path]
    if recursive:
        iterator = []
        for pat in patterns:
            iterator = list(iterator) + list(root.rglob(pat))
    else:
        iterator = []
        for pat in patterns:
            iterator = list(iterator) + list(root.glob(pat))

    for path in iterator:
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if any(part in exclude_dir_names for part in path.parts):
            continue
        key = path.resolve()
        if key in seen:
            continue
        seen.add(key)
        files.append(path)
    return sorted(files)


# B_std 内「有/无支出」归类目录名（不是业务分组）
STD_SPEND_BUCKETS: Set[str] = {"最新月份有支出", "最新月份未支出"}


def list_raw_groups(raw_dir: Path | str) -> List[tuple[str, Path]]:
    """
    发现 A_raw 下一级业务分组目录。

    规则：只认 raw 下的一级子文件夹；组内多个 Excel 合并处理。
    若一级子目录均无 Excel，但 raw 根目录有 Excel，则回退为单组「未分组」。
    """
    from common.paths import strip_order_prefix as _strip_pfx

    raw_dir = Path(raw_dir)
    if not raw_dir.is_dir():
        return []

    groups: List[tuple[str, Path]] = []
    for child in sorted(raw_dir.iterdir(), key=lambda p: p.name):
        if not child.is_dir():
            continue
        if child.name.startswith(".") or child.name.startswith("~$"):
            continue
        if not iter_excel_files(child):
            continue
        groups.append((_strip_pfx(child.name) or child.name, child))

    if groups:
        return groups

    # 兼容：Excel 直接放在 A_raw 根下
    if iter_excel_files(raw_dir, recursive=False) or iter_excel_files(raw_dir):
        return [("未分组", raw_dir)]
    return []


def list_std_groups(std_dir: Path | str) -> List[tuple[str, Path]]:
    """
    发现 B_std 下业务分组目录。

    - 新结构：B_std/{流水|银行账}/最新月份有支出|未支出/*.xlsx
    - 旧结构：B_std/最新月份有支出|未支出/*.xlsx → 单组「未分组」
    """
    std_dir = Path(std_dir)
    if not std_dir.is_dir():
        return []

    groups: List[tuple[str, Path]] = []
    legacy_buckets = False
    for child in sorted(std_dir.iterdir(), key=lambda p: p.name):
        if not child.is_dir():
            continue
        if child.name.startswith(".") or child.name.startswith("~$"):
            continue
        if child.name in STD_SPEND_BUCKETS:
            legacy_buckets = True
            continue
        if not iter_excel_files(child):
            continue
        groups.append((child.name, child))

    if groups:
        return groups

    if legacy_buckets or iter_excel_files(std_dir):
        return [("未分组", std_dir)]
    return []


def resolve_business_groups(
    raw_dir: Path | str,
    std_dir: Path | str,
) -> List[tuple[str, Path]]:
    """
    业务分组唯一口径：优先 A_raw 一级子目录；每组对应 B_std/{组名}/。

    汇总 / 现金流 / 不平明细均应使用本函数，避免扫整个 B_std 根目录混组。
    """
    raw_dir = Path(raw_dir)
    std_dir = Path(std_dir)

    raw_groups = list_raw_groups(raw_dir)
    if raw_groups:
        result: List[tuple[str, Path]] = []
        for name, _raw_path in raw_groups:
            group_std = std_dir / name
            if not group_std.is_dir():
                print(f"  ⚠️ 分组 [{name}] 尚无 B_std 目录，跳过: {group_std}")
                continue
            n = len(iter_excel_files(group_std))
            if n == 0:
                print(f"  ⚠️ 分组 [{name}] B_std 下无 Excel，跳过: {group_std}")
                continue
            result.append((name, group_std))
        return result

    # raw 无一级分组时回退 std
    return list_std_groups(std_dir)


def iter_group_std_files(group_std_dir: Path | str) -> List[Path]:
    """
    仅枚举某一业务分组目录内的 std Excel（不会越界到兄弟分组）。
    """
    root = Path(group_std_dir).resolve()
    if not root.is_dir():
        return []
    out: List[Path] = []
    for fp in iter_excel_files(root):
        try:
            fp.resolve().relative_to(root)
        except ValueError:
            continue
        out.append(fp)
    return sorted(out)


def parse_std_project_filename(file_path: Path | str) -> tuple[str, str]:
    """
    从 B_std 项目文件名解析 (项目编码, 项目名称)。

    支持：
    - 编码_名称_财务付款数据
    - 编码_名称_(YYYY-MM-DD)  （旧命名）
    - 编码_名称
    """
    stem = safe_text(Path(file_path).stem)
    if not stem:
        return "", ""
    stem = re.sub(r"_财务付款数据$", "", stem)
    stem = re.sub(r"_\(\d{4}-\d{2}-\d{2}\)$", "", stem)
    match = re.match(r"^([A-Za-z0-9\-]+)[_\-\s]+(.+)$", stem)
    if not match:
        return "", stem
    return safe_text(match.group(1)), safe_text(match.group(2))


def level1_sort_key(level1: str):
    text = safe_text(level1)
    m = re.match(r"^(\d+)[_\-\.、\s]*(.*)$", text)
    if m:
        return (0, int(m.group(1)), m.group(2))
    return (1, 9999, text)


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
