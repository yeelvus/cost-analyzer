#!/usr/bin/env python3
"""
从 A_成本统计_raw 中，按项目选取最新月份的 Excel，另存到 B_成本统计_std。

规则：
- 扫描 A 下所有子目录中的 .xlsx / .xls / .xlsm
- 忽略临时文件（~$ 开头）和隐藏文件
- 按文件名「{项目名}_成本统计_{YYYYMM}.ext」解析项目与月份
- 同一项目只选取 YYYYMM 最大的一份
- 以「另存为」方式写出到 B（扁平放置），并去掉只读权限，保证可直接编辑
- 若 B 中已有该项目对应文件，则跳过，不覆盖、不更新
"""

from __future__ import annotations

import os
import re
import shutil
import stat
import subprocess
from collections import defaultdict
from pathlib import Path

# 数据目录（脚本已移至 code 目录，路径写死为绝对路径）
SRC = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集/A_成本统计_raw")
DST = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集/B_成本统计_std")

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
# 支持 正鑫源项目_成本统计_202604.xlsm 以及 正鑫源项目_成本统计_202604(2).xlsm
NAME_RE = re.compile(
    r"^(?P<project>.+?)_成本统计_(?P<yyyymm>\d{6})(?:\(\d+\))?\.(?P<ext>xlsx|xls|xlsm)$",
    re.IGNORECASE,
)

# 目标文件权限：所有者可读写，组/他人只读
WRITABLE_MODE = stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP | stat.S_IROTH  # 0o644


def iter_excel_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        name = p.name
        if name.startswith("~$") or name.startswith("."):
            continue
        if p.suffix.lower() not in EXCEL_EXTS:
            continue
        files.append(p)
    return files


def parse_project_name(filename: str) -> str | None:
    m = NAME_RE.match(filename)
    return m.group("project") if m else None


def existing_projects_in_dst(dst: Path) -> dict[str, Path]:
    """扫描目标目录，返回 {项目名: 已有文件路径}。"""
    found: dict[str, Path] = {}
    if not dst.is_dir():
        return found
    for p in dst.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        if p.suffix.lower() not in EXCEL_EXTS:
            continue
        project = parse_project_name(p.name)
        if project:
            # 同项目多个文件时保留第一个即可（跳过时用）
            found.setdefault(project, p)
    return found


def pick_latest_per_project(files: list[Path]) -> dict[str, tuple[str, Path]]:
    """
    返回 {项目名: (YYYYMM, 文件路径)}
    若同月有多个（如带 (2)），优先选修改时间更新、路径更短的文件。
    """
    groups: dict[str, list[tuple[str, Path]]] = defaultdict(list)
    unmatched: list[Path] = []

    for p in files:
        m = NAME_RE.match(p.name)
        if not m:
            unmatched.append(p)
            continue
        project = m.group("project")
        yyyymm = m.group("yyyymm")
        groups[project].append((yyyymm, p))

    if unmatched:
        print("警告：以下文件命名不匹配，已跳过：")
        for p in unmatched:
            try:
                print(f"  - {p.relative_to(SRC)}")
            except ValueError:
                print(f"  - {p}")

    latest: dict[str, tuple[str, Path]] = {}
    for project, versions in groups.items():
        versions.sort(
            key=lambda item: (
                item[0],
                item[1].stat().st_mtime,
                -len(str(item[1])),
            ),
            reverse=True,
        )
        latest[project] = versions[0]
    return latest


def _ensure_writable(path: Path) -> None:
    """去掉只读/锁定，使文件可编辑。"""
    try:
        if hasattr(os, "chflags"):
            os.chflags(path, 0)
    except OSError:
        pass

    path.chmod(WRITABLE_MODE)

    try:
        subprocess.run(
            ["xattr", "-c", str(path)],
            check=False,
            capture_output=True,
        )
    except OSError:
        pass


def save_as_editable(src: Path, dest: Path) -> None:
    """
    另存为可编辑副本：
    - 只拷贝文件内容，不继承源文件的只读权限
    - 目标强制为可写（0o644）
    """
    if dest.exists():
        _ensure_writable(dest)

    shutil.copyfile(src, dest)
    _ensure_writable(dest)


def copy_latest(clear_dst: bool = False) -> None:
    if not SRC.is_dir():
        raise SystemExit(f"源目录不存在: {SRC}")

    DST.mkdir(parents=True, exist_ok=True)

    if clear_dst:
        for old in DST.iterdir():
            if old.is_file() and not old.name.startswith("."):
                try:
                    _ensure_writable(old)
                except OSError:
                    pass
                old.unlink()
                print(f"已清理: {old.name}")

    files = iter_excel_files(SRC)
    if not files:
        raise SystemExit(f"源目录中未找到 Excel 文件: {SRC}")

    latest = pick_latest_per_project(files)
    if not latest:
        raise SystemExit("未能解析出任何符合命名规则的文件。")

    already = existing_projects_in_dst(DST)

    print(f"源目录: {SRC}")
    print(f"目标目录: {DST}")
    print(f"共解析项目 {len(latest)} 个；目标已有项目 {len(already)} 个\n")

    saved = 0
    skipped = 0
    for project in sorted(latest.keys()):
        yyyymm, src_path = latest[project]
        dest_path = DST / src_path.name

        # 已有同名文件，或同项目任意月份文件 → 跳过，不替换
        if dest_path.exists():
            print(f"  [跳过] {project}  目标已有同名文件: {dest_path.name}")
            skipped += 1
            continue
        if project in already:
            print(
                f"  [跳过] {project}  目标已有对应文件: {already[project].name}"
            )
            skipped += 1
            continue

        save_as_editable(src_path, dest_path)
        saved += 1
        mode = oct(dest_path.stat().st_mode & 0o777)
        print(f"  [新增] [{yyyymm}] {project}  (权限 {mode})")
        print(f"         {src_path.relative_to(SRC)}  ->  {dest_path.name}")

    print(f"\n完成：新增 {saved} 个，跳过 {skipped} 个（目标已存在不覆盖）")
    print(f"目标目录: {DST}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="将各项目最新成本统计 Excel 另存为可编辑文件到 B_成本统计_std（已存在则跳过）"
    )
    parser.add_argument(
        "--clear",
        action="store_true",
        help="另存前先清空目标目录（慎用；默认不覆盖已有文件）",
    )
    args = parser.parse_args()
    copy_latest(clear_dst=args.clear)
