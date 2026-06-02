import argparse
import importlib.util
import re
import subprocess
import sys
from pathlib import Path


EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
# 文件夹按关键词动态查找，忽略 A_ 等前缀字母
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(Path(__file__).parent / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    # 若动态加载失败，回退到当前脚本目录，避免脚本直接崩溃
    DEFAULT_FOLDER = str(Path(__file__).parent)
else:
    _路径_mod = importlib.util.module_from_spec(_路径_spec)
    _路径_spec.loader.exec_module(_路径_mod)
    DEFAULT_FOLDER = str(_路径_mod.RAW_DIR)


def build_new_stem(stem: str) -> str:
    """
    Rename pattern example:
    昂纳项目3月份公司采购台账 -> 昂纳项目_3月份公司采购台账

    Rules:
    - Match when text contains "...项目" followed immediately by month token.
    - Month token supports Arabic and common Chinese numerals.
    - Keep names unchanged if already separated (e.g., 项目_3月份...).
    """
    pattern = re.compile(
        r"^(?P<prefix>.+?项目)(?P<zhi>至?)(?P<month>(?:[0-9]{1,2}|[一二三四五六七八九十]{1,3})月份?.*)$"
    )
    m = pattern.match(stem)
    if not m:
        return stem

    prefix = m.group("prefix")
    month_part = m.group("month")  # "至" 已被单独捕获并丢弃

    # If already separated somehow, leave unchanged.
    if prefix.endswith("_") or month_part.startswith(("_", "-", " ")):
        return stem

    return f"{prefix}_{month_part}"


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
        new_stem = build_new_stem(path.stem)
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
