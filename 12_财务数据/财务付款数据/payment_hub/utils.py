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
