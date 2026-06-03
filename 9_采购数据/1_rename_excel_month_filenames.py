import argparse
import importlib.util
import re
import subprocess
import sys
from pathlib import Path


EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
MONTH_PATTERN = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")
# 文件夹按关键词动态查找，忽略 A_ 等前缀字母
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(Path(__file__).parent / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    # 若动态加载失败，回退到当前脚本目录，避免脚本直接崩溃
    DEFAULT_FOLDER = str(Path(__file__).parent)
else:
    _路径_mod = importlib.util.module_from_spec(_路径_spec)
    _路径_spec.loader.exec_module(_路径_mod)
    DEFAULT_FOLDER = str(_路径_mod.RAW_DIR)


def extract_folder_month(path: Path) -> str | None:
    """从文件所在目录或其上级目录提取 YYYY年MM月。"""
    for parent in [path.parent, *path.parents]:
        m = MONTH_PATTERN.search(parent.name)
        if not m:
            continue
        year = int(m.group(1))
        month = int(m.group(2))
        return f"{year:04d}年{month:02d}月"
    return None


def extract_project_prefix(stem: str) -> str:
    """提取项目名称部分，保留已有的 项目名称_ 前缀语义。"""
    base = stem.strip()
    if "_" in base:
        return base.split("_", 1)[0].strip()

    m = re.match(r"^(.+?项目)", base)
    if m:
        return m.group(1).strip()

    return base


def build_new_stem(stem: str, folder_month: str | None) -> str:
    """
    Rename pattern example:
    昂纳项目_8月公司采购台账 -> 昂纳项目_2025年08月_公司采购台账

    Rules:
    - 项目名称部分不变。
    - 后缀统一改为“YYYY年MM月_公司采购台账”，月份来自所在文件夹名。
    - 若无法从上级文件夹提取年月，则保持原文件名不变。
    """
    if not folder_month:
        return stem

    prefix = extract_project_prefix(stem)
    if not prefix:
        return stem
    return f"{prefix}_{folder_month}_公司采购台账"


def rename_excel_files(folder: Path, recursive: bool = True, dry_run: bool = False) -> None:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")

    files = folder.rglob("*") if recursive else folder.glob("*")

    total = 0
    changed = 0
    skipped_conflict = 0

    for path in files:
        if not path.is_file() or path.suffix.lower() not in EXCEL_EXTENSIONS:
            continue

        total += 1
        folder_month = extract_folder_month(path)
        new_stem = build_new_stem(path.stem, folder_month)
        if new_stem == path.stem:
            continue

        target = path.with_name(new_stem + path.suffix)

        if target.exists() and target != path:
            skipped_conflict += 1
            print(f"[SKIP-CONFLICT] {path.name} -> {target.name}")
            continue

        changed += 1
        print(f"[RENAME] {path.name} -> {target.name}")
        if not dry_run:
            path.rename(target)

    mode = "DRY-RUN" if dry_run else "DONE"
    print("-" * 60)
    print(f"{mode} | scanned: {total}, renamed: {changed}, conflict_skipped: {skipped_conflict}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rename Excel files like '昂纳项目3月份公司采购台账' to '昂纳项目_3月份公司采购台账'."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=DEFAULT_FOLDER,
        help=f"Target folder path (default: {DEFAULT_FOLDER}).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Scan subfolders recursively. Default behavior is already recursive.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview changes without renaming.",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Only scan the top-level folder.",
    )

    args = parser.parse_args()
    folder = Path(args.folder).expanduser().resolve()
    recursive = not args.no_recursive
    rename_excel_files(folder=folder, recursive=recursive, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
