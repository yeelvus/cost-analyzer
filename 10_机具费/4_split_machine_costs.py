#!/usr/bin/env python3
"""
按 `项目编码`、`项目名称` 将机具费汇总表切分为多个 sheet，输出为一个或多个 Excel 工作簿。

支持逐 sheet 读取源文件并分别拆分输出，每个输入 sheet 对应一个输出工作簿，工作簿内每个项目分组为一个 sheet。

用法:
  python split_machine_costs.py \
    --source "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum/机具费汇总_全量标准化_汇总.xlsx" \
    --target_dir "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_std"

如果不传参数，会使用上面的默认路径。
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable

try:
    import pandas as pd
except Exception as e:
    print("需要 pandas。请在环境中安装：pip install pandas openpyxl")
    raise


DEFAULT_SOURCE = (
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum/机具费汇总_全量标准化_汇总.xlsx"
)
DEFAULT_TARGET = (
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_std"
)


def sanitize_filename(s: object, max_len: int = 120, empty_replacement: str = "empty") -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        s = ""
    s = str(s).strip()
    if s.lower() in {"nan", "none", "null"}:
        s = ""
    # 去除文件名里常见非法字符
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = s.strip()
    if len(s) > max_len:
        s = s[:max_len]
    return s or empty_replacement


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def safe_sheet_name(name: str, max_len: int = 31) -> str:
    name = sanitize_filename(name, max_len)
    return name[:max_len]


def format_project_filename(code: object, name: object) -> str:
    code_str = sanitize_filename(code, 60, empty_replacement="").strip()
    name_str = sanitize_filename(name, 120, empty_replacement="").strip()
    if code_str and name_str:
        base = f"{code_str}_{name_str}"
    elif code_str:
        base = code_str
    elif name_str:
        base = name_str
    else:
        base = "empty_project"
    return f"{base}_机具费.xlsx"


def build_project_workbooks(sheets: dict[str, pd.DataFrame], cols: Iterable[str], target_dir: Path) -> bool:
    projects: dict[tuple[str, str], dict[str, pd.DataFrame]] = {}
    for sheet_name, df in sheets.items():
        try:
            actual_cols = resolve_columns(df, cols)
        except KeyError as e:
            print(f"跳过 sheet '{sheet_name}'：{e}")
            continue

        grouped = df.groupby(list(actual_cols), dropna=False)
        for keys, group in grouped:
            if not isinstance(keys, tuple):
                keys = (keys,)
            code, name = keys[0], keys[1] if len(keys) > 1 else ""
            if (code is None or (isinstance(code, float) and pd.isna(code)) or str(code).strip() == "") and (
                name is None or (isinstance(name, float) and pd.isna(name)) or str(name).strip() == ""
            ):
                print("跳过：项目编码和项目名称均为空的分组")
                continue
            project_key = (str(code), str(name))
            projects.setdefault(project_key, {})[sheet_name] = group.copy()

    if not projects:
        return False

    for (code, name), sheet_map in projects.items():
        file_name = format_project_filename(code, name)
        output_file = target_dir / safe_sheet_name(file_name, max_len=255)
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            sheet_names: dict[str, int] = {}
            for sheet_name, group in sheet_map.items():
                sheet_tab = safe_sheet_name(sheet_name)
                count = sheet_names.get(sheet_tab, 0)
                if count:
                    sheet_tab = safe_sheet_name(f"{sheet_tab}_{count}")
                sheet_names[sheet_tab] = count + 1
                group.to_excel(writer, sheet_name=sheet_tab, index=False)
        print(f"写入: {output_file}")
    return True


def resolve_columns(df: pd.DataFrame, desired: list[str]) -> list[str]:
    """尝试在 df.columns 中找到与 desired 中每项匹配的实际列名。

    匹配策略：
    1) 精确匹配
    2) 不区分大小写的包含匹配（column name contains desired token）
    返回找到的实际列名列表，顺序与 desired 保持一致。
    如果某一项无法匹配，会抛出 KeyError。
    """
    cols = list(df.columns)
    found: list[str] = []
    lower_cols = [c.lower() for c in cols]
    for token in desired:
        if token in cols:
            found.append(token)
            continue
        t = token.lower()
        # 包含匹配
        matches = [cols[i] for i, c in enumerate(lower_cols) if t in c]
        if matches:
            # 取第一个匹配
            found.append(matches[0])
            continue
        raise KeyError(f"无法在表头中找到匹配列: {token}")
    return found


def read_sheets(source: Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(source, sheet_name=None, dtype=object)
    result: dict[str, pd.DataFrame] = {}
    for sheet_name, sheet_df in sheets.items():
        if sheet_df.empty:
            continue
        sheet_df = sheet_df.where(pd.notnull(sheet_df), None)
        result[sheet_name] = sheet_df
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="按项目编码/名称逐 sheet 拆分机具费汇总表")
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    parser.add_argument("--target_dir", default=DEFAULT_TARGET)
    parser.add_argument(
        "--cols",
        nargs="*",
        help="用于分组的列，默认依次为 项目编码 项目名称",
    )
    args = parser.parse_args(argv)

    source = Path(args.source)
    target_dir = Path(args.target_dir)
    cols = args.cols or ["项目编码", "项目名称"]

    if not source.exists():
        print(f"源文件不存在: {source}")
        return 2

    ensure_dir(target_dir)

    print(f"读取: {source}")
    sheets = read_sheets(source)
    if not sheets:
        print("源文件没有可读的数据。请检查是否有至少一个非空 sheet。")
        return 4

    written = build_project_workbooks(sheets, cols, target_dir)
    if not written:
        print("没有有效的拆分结果。请检查 sheet 表头是否包含项目编码/项目名称。")
        return 5

    print("全部完成。输出目录:", target_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
