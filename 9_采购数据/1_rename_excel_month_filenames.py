#!/usr/bin/env python3
"""
步骤1：按所在月份文件夹，将 raw 下 Excel 统一改名为
    {项目名}_{YYYY年MM月}_公司采购台账.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

from _采购路径 import RAW_DIR, extract_month_cn, extract_project_name, is_excel


def build_new_stem(stem: str, folder_month: str | None) -> str:
    if not folder_month:
        return stem
    prefix = extract_project_name(stem)
    return f"{prefix}_{folder_month}_公司采购台账" if prefix else stem


def rename_excel_files(folder: Path, *, recursive: bool = True, dry_run: bool = False) -> None:
    if not folder.is_dir():
        raise FileNotFoundError(f"目录不存在: {folder}")

    paths = [p for p in (folder.rglob("*") if recursive else folder.iterdir()) if is_excel(p)]
    renamed = conflicts = 0

    for path in paths:
        month = extract_month_cn(str(path.parent)) or extract_month_cn(path.name)
        # 向上找月份文件夹
        if not month:
            for parent in path.parents:
                month = extract_month_cn(parent.name)
                if month:
                    break

        new_stem = build_new_stem(path.stem, month)
        if new_stem == path.stem:
            continue

        target = path.with_name(new_stem + path.suffix)
        if target.exists() and target != path:
            conflicts += 1
            print(f"[冲突跳过] {path.name} -> {target.name}")
            continue

        renamed += 1
        print(f"[重命名] {path.name} -> {target.name}")
        if not dry_run:
            path.rename(target)

    mode = "预览" if dry_run else "完成"
    print("-" * 50)
    print(f"{mode} | 扫描 {len(paths)}，重命名 {renamed}，冲突 {conflicts}")


def main() -> None:
    parser = argparse.ArgumentParser(description="按月份文件夹规范采购台账文件名")
    parser.add_argument("folder", nargs="?", default=str(RAW_DIR), help=f"默认: {RAW_DIR}")
    parser.add_argument("--dry-run", action="store_true", help="只预览不改名")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    args = parser.parse_args()

    rename_excel_files(
        Path(args.folder).expanduser().resolve(),
        recursive=not args.no_recursive,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
