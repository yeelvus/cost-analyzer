#!/usr/bin/env python3
"""
0_Raw/2_成本统计收集 与 0_Std/2_成本统计收集 双向同步。

默认流程（一键运行会先做这一步）：
1. 把 STD 里已有文件全部回写到 RAW 对应月份文件夹（覆盖 RAW 同名文件，保留你在 STD 的修改）
2. 再把 RAW 里各项目「最新月份」补进 STD：
   - STD 已有同一项目同一月份 → 不覆盖，完整保留
   - STD 没有该最新月份 → 新增拷贝，不删 STD 里其它月份

文件名：{项目名}_成本统计_{YYYYMM}.xlsx/.xlsm ，允许空格和 (2) 副本。
"""

from __future__ import annotations

import os
import re
import shutil
import stat
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))
from common.paths import find_named_dir, month_label_from_dirname, raw_data_path, std_data_path  # noqa: E402

# 原始成本统计在 0_Raw/2_成本统计收集；未迁移时回退旧 raw 目录
SRC = raw_data_path(
    "成本统计收集",
    legacy=Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集/A_成本统计_raw"),
)
DST = std_data_path(
    "成本统计收集",
    legacy=Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集/B_成本统计_std"),
)

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
# 支持 正鑫源项目_成本统计_202604.xlsm 以及 正鑫源项目_成本统计_202604(2).xlsm
NAME_RE = re.compile(
    r"^(?P<project>.+?)_成本统计_(?P<yyyymm>\d{6})(?:\(\d+\))?\s*\.(?P<ext>xlsx|xls|xlsm)$",
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


def parse_cost_name(filename: str) -> tuple[str, str] | None:
    m = NAME_RE.match(filename)
    if not m:
        return None
    return m.group("project").strip(), m.group("yyyymm")


def parse_project_name(filename: str) -> str | None:
    parsed = parse_cost_name(filename)
    return parsed[0] if parsed else None


def month_folder_name(yyyymm: str) -> str:
    return f"{int(yyyymm[:4])}年{int(yyyymm[4:6]):02d}月"


def resolve_month_dir(parent: Path, yyyymm: str) -> Path:
    """已有 1_2026年08月 或 2026年08月 都算同一月份。"""
    label = month_folder_name(yyyymm)
    found = find_named_dir(parent, label)
    return found if found is not None else parent / label


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
        project = m.group("project").strip()
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


def iter_std_excels() -> list[Path]:
    if not DST.is_dir():
        return []
    files: list[Path] = []
    for p in DST.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        if p.suffix.lower() not in EXCEL_EXTS:
            continue
        files.append(p)
    return files


def sync_std_back_to_raw() -> tuple[int, int, int]:
    """把 STD 全部文件回写到 RAW 对应月份目录；同名则覆盖 RAW。"""
    added = overwritten = failed = 0
    std_files = iter_std_excels()
    if not std_files:
        print("STD 中没有可回写的 Excel，跳过。")
        return 0, 0, 0

    print(f"\n--- 1/2  STD → RAW（回写修改，共 {len(std_files)} 个）---")
    for src in sorted(std_files, key=lambda p: p.name):
        parsed = parse_cost_name(src.name)
        if not parsed:
            print(f"  [跳过命名] {src.name}")
            continue
        _project, yyyymm = parsed
        dest_dir = resolve_month_dir(SRC, yyyymm)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / src.name
        existed = dest.exists()
        try:
            save_as_editable(src, dest)
        except OSError as e:
            failed += 1
            print(f"  [失败] {src.name} → {dest.relative_to(SRC)}  ({e})")
            continue
        rel = dest.relative_to(SRC)
        if existed:
            overwritten += 1
            print(f"  [覆盖 RAW] [{yyyymm}] {src.name}  →  {rel}")
        else:
            added += 1
            print(f"  [写入 RAW] [{yyyymm}] {src.name}  →  {rel}")
    print(f"回写完成：新增 {added}，覆盖 {overwritten}，失败 {failed}")
    return added, overwritten, failed


def fill_missing_latest_from_raw() -> tuple[int, int]:
    """把 RAW 各项目最新月份补进 STD；已有同一项目同一月份则保留 STD。"""
    DST.mkdir(parents=True, exist_ok=True)
    std_keys: dict[tuple[str, str], Path] = {}
    for p in iter_std_excels():
        parsed = parse_cost_name(p.name)
        if parsed:
            std_keys[parsed] = p

    files = iter_excel_files(SRC)
    if not files:
        raise SystemExit(f"源目录中未找到 Excel 文件: {SRC}")
    latest = pick_latest_per_project(files)
    if not latest:
        raise SystemExit("未能解析出任何符合命名规则的文件。")

    print(f"\n--- 2/2  RAW 最新 → STD（不覆盖已有，共 {len(latest)} 个项目）---")
    added = skipped = 0
    for project in sorted(latest.keys()):
        yyyymm, src_path = latest[project]
        key = (project, yyyymm)
        if key in std_keys:
            skipped += 1
            print(f"  [保留 STD] [{yyyymm}] {project}  已有 {std_keys[key].name}")
            continue
        dest_path = DST / src_path.name
        if dest_path.exists():
            skipped += 1
            print(f"  [保留 STD] [{yyyymm}] {project}  同名已存在 {dest_path.name}")
            continue
        save_as_editable(src_path, dest_path)
        added += 1
        print(f"  [补齐最新] [{yyyymm}] {project}")
        print(f"             {src_path.relative_to(SRC)}  →  {dest_path.name}")
    print(f"补齐完成：新增 {added}，STD 已有故保留 {skipped}")
    return added, skipped


def sync_cost_std_and_raw() -> None:
    if not SRC.is_dir():
        raise SystemExit(f"RAW 目录不存在: {SRC}")
    DST.mkdir(parents=True, exist_ok=True)
    print("=" * 72)
    print("成本文件同步")
    print(f"  RAW: {SRC}")
    print(f"  STD: {DST}")
    print("=" * 72)
    sync_std_back_to_raw()
    fill_missing_latest_from_raw()
    print("\nSTD 现有文件：")
    for p in sorted(iter_std_excels(), key=lambda x: x.name):
        parsed = parse_cost_name(p.name)
        tag = f"{parsed[1]} {parsed[0]}" if parsed else p.name
        print(f"  - {p.name}  ({tag})")
    print(f"\n同步结束。STD 原有文件均未删除、未覆盖。")


def copy_latest(clear_dst: bool = False) -> None:
    if clear_dst:
        print("先把 STD 回写 RAW，再清空 STD 并从 RAW 重建最新（--clear）")
        if DST.is_dir():
            sync_std_back_to_raw()
            for old in list(DST.iterdir()):
                if old.is_file() and not old.name.startswith("."):
                    try:
                        _ensure_writable(old)
                    except OSError:
                        pass
                    old.unlink()
                    print(f"已清理 STD: {old.name}")
        fill_missing_latest_from_raw()
        return
    sync_cost_std_and_raw()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="STD 修改回写 RAW，再把 RAW 各项目最新月份补进 STD（不覆盖 STD 已有文件）"
    )
    parser.add_argument(
        "--clear",
        action="store_true",
        help="回写 RAW 后清空 STD，再从 RAW 重建各项目最新文件（会丢掉 STD 里的旧月份副本）",
    )
    args = parser.parse_args()
    copy_latest(clear_dst=args.clear)
