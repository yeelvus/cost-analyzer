# -*- coding: utf-8 -*-
"""拆分工具：

1) 原始拆分 split_raw_merge_by_project
   输入：B_机具费_raw_sum/机具费_raw合并.xlsx
   输出：C_机具费_raw_split/{项目编码}_{项目名称}_机具费.xlsx

2) 标准拆分 split_std_summary_by_project
    输入：D_机具费_std_sum/机具费_汇总.xlsx
   输出：E_机具费_std_split/{项目编码}_{项目名称}_机具费.xlsx
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

from machine.config import (
    FULL_SUMMARY_FILENAME,
    LEGACY_CLEAN_MERGE_NAME,
    LEGACY_STD_DIR,
    LEGACY_SUM_DIR,
    PROJECT_STD_SUFFIX,
    RAW_MERGE_NAME,
    RAW_SPLIT_DIR,
    RAW_SUM_DIR,
    STD_SPLIT_DIR,
    STD_SUM_DIR,
    ensure_data_dirs,
)


def sanitize_filename(s: object, max_len: int = 120, empty_replacement: str = "empty") -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        s = ""
    s = str(s).strip()
    if s.lower() in {"nan", "none", "null"}:
        s = ""
    s = re.sub(r'[\\/:*?"<>|]+', "_", s).strip()
    s = re.sub(r"\s+", "", s)
    if len(s) > max_len:
        s = s[:max_len]
    return s or empty_replacement


def safe_sheet_name(name: str, max_len: int = 31) -> str:
    return sanitize_filename(name, max_len)[:max_len]


def format_project_std_filename(code: object, name: object) -> str:
    code_str = sanitize_filename(code, 40, empty_replacement="未编码")
    name_str = sanitize_filename(name, 80, empty_replacement="未命名项目")
    suffix = PROJECT_STD_SUFFIX if PROJECT_STD_SUFFIX.startswith("_") else f"_{PROJECT_STD_SUFFIX}"
    return f"{code_str}_{name_str}{suffix}"


def resolve_columns(df: pd.DataFrame, desired: list[str]) -> list[str]:
    cols = list(df.columns)
    lower_cols = [str(c).lower() for c in cols]
    found = []
    for token in desired:
        if token in cols:
            found.append(token)
            continue
        t = token.lower()
        matches = [cols[i] for i, c in enumerate(lower_cols) if t in c]
        if matches:
            found.append(matches[0])
            continue
        raise KeyError(f"无法在表头中找到匹配列: {token}")
    return found


def _is_empty_val(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    s = str(v).strip()
    return s == "" or s.lower() in {"nan", "none", "null"}


def resolve_raw_merge_path() -> Path:
    ensure_data_dirs()
    candidates = [
        RAW_SUM_DIR / RAW_MERGE_NAME,
        LEGACY_STD_DIR / RAW_MERGE_NAME,
        LEGACY_STD_DIR / LEGACY_CLEAN_MERGE_NAME,
        RAW_SUM_DIR / LEGACY_CLEAN_MERGE_NAME,
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"未找到过程文件 {RAW_MERGE_NAME}。请先运行 merge-raw。\n"
        f"已查: {[str(c) for c in candidates]}"
    )


def resolve_full_summary_path() -> Path:
    ensure_data_dirs()
    candidates = [
        STD_SUM_DIR / FULL_SUMMARY_FILENAME,
        *sorted(STD_SUM_DIR.glob("*全量标准化*汇总*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True),
        *sorted(LEGACY_SUM_DIR.glob("*全量标准化*汇总*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True),
    ]
    for p in candidates:
        if p.exists() and not p.name.startswith("~$"):
            return p
    raise FileNotFoundError(f"未找到标准汇总文件（期望 {FULL_SUMMARY_FILENAME}）。请先运行 summarize。")


def _clean_project_files(output_dir: Path) -> int:
    removed = 0
    for p in output_dir.glob("*.xlsx"):
        if p.name.startswith("~$") or "合并" in p.name:
            continue
        if p.name.endswith("机具费.xlsx") or "_机具费" in p.name:
            try:
                p.unlink()
                removed += 1
            except Exception:
                pass
    return removed


def split_raw_merge_by_project(
    input_path: Path | None = None,
    output_dir: Path | None = None,
    *,
    clean_old_project_files: bool = True,
) -> list[Path]:
    """原始拆分：机具费_raw合并 → C_机具费_raw_split/ 每项目一个「明细」sheet。"""
    ensure_data_dirs()
    output_dir = Path(output_dir or RAW_SPLIT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    input_path = Path(input_path) if input_path else resolve_raw_merge_path()

    print(f"读取过程文件: {input_path}")
    df = pd.read_excel(input_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"  行数={len(df)} 列={list(df.columns)}")

    code_col = "项目编码" if "项目编码" in df.columns else None
    name_col = "项目名称" if "项目名称" in df.columns else None
    use_col = "使用项目" if "使用项目" in df.columns else None
    if not name_col and not use_col:
        raise SystemExit("合并表中无「项目名称」或「使用项目」列，无法按项目拆分")

    def project_key(row):
        code = row[code_col] if code_col else ""
        name = row[name_col] if name_col else ""
        use = row[use_col] if use_col else ""
        if _is_empty_val(name) and not _is_empty_val(use):
            name = use
        if _is_empty_val(code) and _is_empty_val(name):
            return ("未匹配", "未匹配项目")
        code_s = "未编码" if _is_empty_val(code) else str(code).strip()
        name_s = "未命名项目" if _is_empty_val(name) else str(name).strip()
        return (code_s, name_s)

    keys = df.apply(project_key, axis=1)
    df = df.copy()
    df["_code"] = [k[0] for k in keys]
    df["_name"] = [k[1] for k in keys]

    if clean_old_project_files:
        n = _clean_project_files(output_dir)
        if n:
            print(f"  已清理旧 C_raw_split {n} 个")

    written: list[Path] = []
    groups = df.groupby(["_code", "_name"], dropna=False)
    print(f"\nC_机具费_raw_split：共 {groups.ngroups} 个项目 → {output_dir}")

    for (code, name), group in groups:
        out_path = output_dir / format_project_std_filename(code, name)
        g = group.drop(columns=["_code", "_name"], errors="ignore")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            g.to_excel(writer, sheet_name="明细", index=False)
        written.append(out_path)
        print(f"  ✓ {out_path.name}  ({len(g)} 行)")

    print(f"\n✅ 原始拆分(C)完成：{len(written)} 个文件")
    print(f"   过程文件: {input_path}")
    return written


def split_std_summary_by_project(
    source: Path | None = None,
    target_dir: Path | None = None,
    cols: list[str] | None = None,
    *,
    clean_old: bool = True,
) -> int:
    """标准拆分：标准汇总 → E_机具费_std_split/ 每项目一个文件。"""
    ensure_data_dirs()
    target_dir = Path(target_dir or STD_SPLIT_DIR)
    target_dir.mkdir(parents=True, exist_ok=True)
    source = Path(source) if source else resolve_full_summary_path()
    cols = cols or ["项目编码", "项目名称"]
    print(f"读取 D_机具费_std_sum: {source}")
    print(f"输出 E_机具费_std_split: {target_dir}")

    xl = pd.ExcelFile(source)
    # 优先「汇总」sheet
    sheet = "汇总" if "汇总" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(source, sheet_name=sheet)
    if df.empty:
        print("汇总 sheet 为空")
        return 1

    try:
        actual_cols = resolve_columns(df, cols)
    except KeyError as e:
        print(f"无法拆分: {e}")
        return 2

    if clean_old:
        n = _clean_project_files(target_dir)
        if n:
            print(f"  已清理旧 E_std_split {n} 个")

    projects: dict[tuple[str, str], pd.DataFrame] = {}
    for keys, group in df.groupby(list(actual_cols), dropna=False):
        if not isinstance(keys, tuple):
            keys = (keys,)
        code, name = keys[0], keys[1] if len(keys) > 1 else ""
        if _is_empty_val(code) and _is_empty_val(name):
            continue
        code_s = "未编码" if _is_empty_val(code) else str(code).strip()
        name_s = "未命名项目" if _is_empty_val(name) else str(name).strip()
        projects[(code_s, name_s)] = group.copy()

    if not projects:
        print("未生成任何项目文件")
        return 1

    for (code, name), group in projects.items():
        out_path = target_dir / format_project_std_filename(code, name)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            # 汇总明细放第一个 sheet
            group.to_excel(writer, sheet_name="汇总", index=False)
        print(f"  ✓ {out_path.name}  ({len(group)} 行)")

    print(f"\n✅ 标准拆分(E)完成：{len(projects)} 个文件 → {target_dir}")
    return 0


# 兼容旧名
def split_ledger_by_project(input_path=None, output_dir=None):
    files = split_raw_merge_by_project(input_path, output_dir)
    return files[0] if files else Path(output_dir or RAW_SPLIT_DIR)


def split_summary_by_project(source=None, target_dir=None, cols=None):
    return split_std_summary_by_project(source, target_dir, cols)


def main_split_sum(argv=None):
    p = argparse.ArgumentParser(description="标准汇总 → 标准拆分")
    p.add_argument("--source", default=None)
    p.add_argument("--target_dir", default=None)
    p.add_argument("--cols", nargs="*", default=None)
    args = p.parse_args(argv)
    raise SystemExit(split_std_summary_by_project(args.source, args.target_dir, args.cols))


def main_split_ledger(argv=None):
    p = argparse.ArgumentParser(description="机具费_raw合并 → 原始拆分")
    p.add_argument("--input", default=None)
    p.add_argument("--output-dir", default=None)
    args = p.parse_args(argv)
    split_raw_merge_by_project(args.input, args.output_dir)


if __name__ == "__main__":
    main_split_ledger()
