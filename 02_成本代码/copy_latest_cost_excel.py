#!/usr/bin/env python3
"""
01_Raw/02_成本统计收集 与 02_Std/02_成本统计收集 双向同步。

默认流程：
1. 把 STD 里已有文件按原文件名回写到 RAW 对应月份文件夹（覆盖 RAW 同名文件，保留 STD 上的修改）
2. 再按项目名称匹配库，把 RAW 里各标准项目的「最新月份」放到 STD：
   - 别名并到标准名。例如「腾洋光学」「腾洋项目」只留「腾洋项目」一份
   - 未在匹配库中的名称保持原名，不和其他项目合并
   - 只做精确别名匹配，不做模糊匹配，也不写回匹配库
   - STD 已有该标准项目的同一月份 → 保留 STD 内容，必要时改成标准文件名
   - RAW 有更新的月份 → 拷贝进 STD，并删掉 STD 里同一标准项目的旧文件或别名文件
   - RAW 月份目录里的历史文件不删除

STD 文件名：{YYYYMM}_{项目编码}_{标准项目名}_成本统计.xlsx/.xlsm
RAW 文件名保持收集时的原名，允许空格和 (2) 副本。
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
    "成本统计收集_raw",
    legacy=Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_成本统计收集_raw"),
)
DST = std_data_path(
    "成本统计收集",
    legacy=Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/02_成本统计收集"),
)

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
# RAW：正鑫源项目_成本统计_202604.xlsm ，允许 (2) 副本
NAME_RE = re.compile(
    r"^(?P<project>.+?)_成本统计_(?P<yyyymm>\d{6})(?:\(\d+\))?\s*\.(?P<ext>xlsx|xls|xlsm)$",
    re.IGNORECASE,
)
# STD：202609_THCC-2026-002_东威科技_成本统计.xlsm
STD_NAME_RE = re.compile(
    r"^(?P<yyyymm>\d{6})_(?P<code>[^_]+)_(?P<project>.+)_成本统计\s*\.(?P<ext>xlsx|xls|xlsm)$",
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
    """返回 (项目名, YYYYMM)。同时识别收集原名和标准目录名。"""
    coded = STD_NAME_RE.match(filename)
    if coded:
        return coded.group("project").strip(), coded.group("yyyymm")
    m = NAME_RE.match(filename)
    if not m:
        return None
    return m.group("project").strip(), m.group("yyyymm")


def parse_project_name(filename: str) -> str | None:
    parsed = parse_cost_name(filename)
    return parsed[0] if parsed else None


_MATCHER = None


def _matcher():
    """只读匹配库。不读项目基本信息表，避免把未确认的名称写进别名。"""
    global _MATCHER
    if _MATCHER is None:
        from common.project_name_matcher import ProjectMatcher

        _MATCHER = ProjectMatcher(load_info=False)
    return _MATCHER


def _match_hit(name: str) -> dict | None:
    name = name.strip()
    if not name:
        return None
    hit = _matcher().match(name, fuzzy=False)
    if not hit or hit.get("source") in {"unmatched_keep", "unmatched"}:
        return None
    return hit


def canonical_project(name: str) -> str:
    """精确别名 → 标准项目名。未登记则保持原名。"""
    hit = _match_hit(name)
    if not hit:
        return name.strip()
    canon = str(hit.get("project_name") or "").strip()
    return canon or name.strip()


def project_code(name: str) -> str:
    """标准项目编码。没有正式编码时用「未编码」。"""
    hit = _match_hit(name)
    if not hit:
        return "未编码"
    code = str(hit.get("project_code") or "").strip()
    if not code or code in {"需手动核对", "未匹配", "未编码"}:
        return "未编码"
    return code


def std_filename(canon: str, yyyymm: str, suffix: str, code: str | None = None) -> str:
    """日期_项目编码_项目名称_成本统计。日期用统计月份 YYYYMM。"""
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    code_text = (code if code is not None else project_code(canon)).strip() or "未编码"
    code_text = re.sub(r'[\\/:*?"<>|]+', "_", code_text)
    name_text = re.sub(r'[\\/:*?"<>|]+', "_", canon.strip()) or "未命名项目"
    return f"{yyyymm}_{code_text}_{name_text}_成本统计{suffix.lower()}"


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
    返回 {标准项目名: (YYYYMM, 文件路径)}。
    文件名里的别名先并到匹配库标准名，再在每个标准项目里取最新月份。
    同月多份时，优先标准名本身，其次修改时间更新、路径更短的文件。
    """
    groups: dict[str, list[tuple[str, Path]]] = defaultdict(list)
    literals: dict[str, set[str]] = defaultdict(set)
    unmatched: list[Path] = []

    for p in files:
        parsed = parse_cost_name(p.name)
        if not parsed:
            unmatched.append(p)
            continue
        project, yyyymm = parsed
        canon = canonical_project(project)
        groups[canon].append((yyyymm, p))
        literals[canon].add(project)

    if unmatched:
        print("警告：以下文件命名不匹配，已跳过：")
        for p in unmatched:
            try:
                print(f"  - {p.relative_to(SRC)}")
            except ValueError:
                print(f"  - {p}")

    latest: dict[str, tuple[str, Path]] = {}
    for canon, versions in groups.items():
        versions.sort(
            key=lambda item: (
                item[0],
                1 if parse_project_name(item[1].name) == canon else 0,
                item[1].stat().st_mtime,
                -len(str(item[1])),
            ),
            reverse=True,
        )
        latest[canon] = versions[0]
        names = literals[canon]
        if len(names) > 1:
            print(f"  合并别名 → {canon}: {', '.join(sorted(names))}")
    return latest


def _ensure_writable(path: Path) -> None:
    """去掉只读和锁定，使复制后的文件可以修改。"""
    if not path.is_file():
        return
    try:
        if hasattr(os, "chflags"):
            os.chflags(path, 0)
    except OSError:
        pass
    mode = path.stat().st_mode
    path.chmod(mode | stat.S_IRUSR | stat.S_IWUSR)
    if not os.access(path, os.W_OK):
        path.chmod(WRITABLE_MODE)
    try:
        subprocess.run(
            ["chmod", "u+w", str(path)],
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


def raw_dest_for_std(std_path: Path) -> Path | None:
    """STD 回写到 RAW 时，覆盖同一标准项目、同一月份里已有的收集文件。

    这样 STD 改成标准文件名之后，不会在月份目录里再复制出第二份。
    """
    parsed = parse_cost_name(std_path.name)
    if not parsed:
        return None
    literal, yyyymm = parsed
    canon = canonical_project(literal)
    month_dir = resolve_month_dir(SRC, yyyymm)
    peers: list[Path] = []
    if month_dir.is_dir():
        for path in month_dir.iterdir():
            if not path.is_file() or path.name.startswith("~$") or path.name.startswith("."):
                continue
            if path.suffix.lower() not in EXCEL_EXTS:
                continue
            found = parse_cost_name(path.name)
            if not found:
                continue
            raw_literal, raw_month = found
            if raw_month == yyyymm and canonical_project(raw_literal) == canon:
                peers.append(path)
    if not peers:
        return month_dir / std_path.name
    peers.sort(
        key=lambda path: (
            path.name == std_path.name,
            parse_project_name(path.name) == literal,
            parse_project_name(path.name) == canon,
            path.stat().st_mtime,
        ),
        reverse=True,
    )
    return peers[0]


def sync_std_back_to_raw(*, dry_run: bool = False) -> tuple[int, int, int]:
    """把 STD 修改回写到 RAW 对应月份里的同一项目文件。"""
    added = overwritten = failed = 0
    std_files = iter_std_excels()
    if not std_files:
        print("STD 中没有可回写的 Excel，跳过。")
        return 0, 0, 0

    print(f"\n--- 1/2  STD → RAW（回写修改，共 {len(std_files)} 个）---")
    for src in sorted(std_files, key=lambda p: p.name):
        dest = raw_dest_for_std(src)
        if dest is None:
            print(f"  [跳过命名] {src.name}")
            continue
        parsed = parse_cost_name(src.name)
        yyyymm = parsed[1] if parsed else ""
        existed = dest.exists()
        rel = dest.relative_to(SRC)
        if dry_run:
            verb = "将覆盖 RAW" if existed else "将写入 RAW"
            print(f"  [{verb}] [{yyyymm}] {src.name}  →  {rel}")
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            save_as_editable(src, dest)
        except OSError as e:
            failed += 1
            print(f"  [失败] {src.name} → {rel}  ({e})")
            continue
        if existed:
            overwritten += 1
            print(f"  [覆盖 RAW] [{yyyymm}] {src.name}  →  {rel}")
        else:
            added += 1
            print(f"  [写入 RAW] [{yyyymm}] {src.name}  →  {rel}")
    if dry_run:
        print("回写预览结束（未写文件）。")
    else:
        print(f"回写完成：新增 {added}，覆盖 {overwritten}，失败 {failed}")
    return added, overwritten, failed


def _std_by_canonical() -> dict[str, list[tuple[str, Path]]]:
    grouped: dict[str, list[tuple[str, Path]]] = defaultdict(list)
    for path in iter_std_excels():
        parsed = parse_cost_name(path.name)
        if not parsed:
            print(f"  [跳过命名] {path.name}")
            continue
        project, yyyymm = parsed
        grouped[canonical_project(project)].append((yyyymm, path))
    return grouped


def _drop_other_std(keep: Path, items: list[tuple[str, Path]], *, dry_run: bool) -> int:
    removed = 0
    keep_resolved = keep.resolve() if keep.exists() else None
    for _yyyymm, path in items:
        if not path.exists():
            continue
        if keep_resolved is not None and path.resolve() == keep_resolved:
            continue
        removed += 1
        if dry_run:
            print(f"  [将移除重复] {path.name}")
            continue
        try:
            _ensure_writable(path)
            path.unlink()
        except OSError as e:
            print(f"  [移除失败] {path.name}  ({e})")
            continue
        print(f"  [移除重复] {path.name}")
    return removed


def _rename_std(src: Path, dest: Path) -> Path:
    _ensure_writable(src)
    if dest.exists() and dest.resolve() != src.resolve():
        _ensure_writable(dest)
        dest.unlink()
    src.rename(dest)
    _ensure_writable(dest)
    return dest


def fill_missing_latest_from_raw(*, dry_run: bool = False) -> tuple[int, int, int]:
    """每个标准项目在 STD 只留最新一份，文件名使用标准项目名。"""
    if not dry_run:
        DST.mkdir(parents=True, exist_ok=True)
    files = iter_excel_files(SRC)
    if not files:
        raise SystemExit(f"源目录中未找到 Excel 文件: {SRC}")
    print("\n--- 别名合并 ---")
    latest = pick_latest_per_project(files)
    if not latest:
        raise SystemExit("未能解析出任何符合命名规则的文件。")

    std_groups = _std_by_canonical()
    canons = sorted(set(latest) | set(std_groups))
    print(f"\n--- 2/2  RAW 最新 → STD（每个标准项目一份，共 {len(canons)} 个）---")
    added = kept = removed = 0
    for canon in canons:
        std_items = std_groups.get(canon, [])
        if canon not in latest:
            std_items.sort(key=lambda item: (item[0], item[1].stat().st_mtime), reverse=True)
            yyyymm, keeper = std_items[0]
            dest = DST / std_filename(canon, yyyymm, keeper.suffix)
            print(f"  [仅 STD] [{yyyymm}] {canon}  RAW 中没有，保留 {keeper.name}")
            if keeper.name != dest.name:
                print(f"             改名为 {dest.name}")
                if not dry_run:
                    keeper = _rename_std(keeper, dest)
            elif not dry_run:
                _ensure_writable(keeper)
            kept += 1
            removed += _drop_other_std(keeper, std_items, dry_run=dry_run)
            continue

        yyyymm, src_path = latest[canon]
        dest = DST / std_filename(canon, yyyymm, src_path.suffix)
        same_month = [path for month, path in std_items if month == yyyymm]
        literal = parse_project_name(src_path.name) or ""
        if same_month:
            same_month.sort(
                key=lambda path: (path.name == dest.name, path.stat().st_mtime),
                reverse=True,
            )
            keeper = same_month[0]
            if keeper.name == dest.name:
                print(f"  [保留 STD] [{yyyymm}] {canon}  已有 {keeper.name}")
                if not dry_run:
                    _ensure_writable(keeper)
            else:
                print(f"  [改用标准名] [{yyyymm}] {canon}")
                print(f"             {keeper.name}  →  {dest.name}")
                if not dry_run:
                    keeper = _rename_std(keeper, dest)
            kept += 1
            removed += _drop_other_std(dest if dest.exists() else keeper, std_items, dry_run=dry_run)
            continue

        print(f"  [补齐最新] [{yyyymm}] {canon}")
        if literal != canon:
            print(f"             来源别名 {literal}")
        print(f"             {src_path.relative_to(SRC)}  →  {dest.name}")
        if not dry_run:
            save_as_editable(src_path, dest)
        added += 1
        removed += _drop_other_std(dest, std_items, dry_run=dry_run)
    if dry_run:
        print(f"预览：将新增 {added}，保留或改名 {kept}，将移除重复 {removed}")
    else:
        unlocked = unlock_outputs()
        print(f"补齐完成：新增 {added}，保留或改名 {kept}，移除重复 {removed}，可修改 {unlocked}")
    return added, kept, removed


def unlock_outputs() -> int:
    """标准目录里的成本文件一律改为当前用户可写。"""
    count = 0
    for path in iter_std_excels():
        before = path.stat().st_mode
        _ensure_writable(path)
        if not os.access(path, os.W_OK):
            raise OSError(f"复制后的文件仍然只读: {path}")
        if not before & stat.S_IWUSR:
            print(f"  [改为可修改] {path.name}")
        count += 1
    return count


def _print_std_now() -> None:
    print("\nSTD 现有文件：")
    for p in sorted(iter_std_excels(), key=lambda x: x.name):
        parsed = parse_cost_name(p.name)
        if not parsed:
            print(f"  - {p.name}")
            continue
        project, yyyymm = parsed
        canon = canonical_project(project)
        mark = "" if canon == project else f"  标准名 {canon}"
        print(f"  - {p.name}  ({yyyymm} {project}{mark})")


def sync_cost_std_and_raw(*, dry_run: bool = False) -> None:
    if not SRC.is_dir():
        raise SystemExit(f"RAW 目录不存在: {SRC}")
    if not dry_run:
        DST.mkdir(parents=True, exist_ok=True)
    print("=" * 72)
    print("成本文件同步" + ("（预览，不写文件）" if dry_run else ""))
    print(f"  RAW: {SRC}")
    print(f"  STD: {DST}")
    print("=" * 72)
    sync_std_back_to_raw(dry_run=dry_run)
    fill_missing_latest_from_raw(dry_run=dry_run)
    if not dry_run:
        _print_std_now()
        print("\n同步结束。RAW 月份目录中的历史文件都还在。STD 每个标准项目只留最新一份。")


def copy_latest(clear_dst: bool = False, *, dry_run: bool = False) -> None:
    if clear_dst:
        print("先把 STD 回写 RAW，再清空 STD 并从 RAW 重建最新（--clear）")
        if DST.is_dir():
            sync_std_back_to_raw(dry_run=dry_run)
            if dry_run:
                for old in sorted(iter_std_excels(), key=lambda p: p.name):
                    print(f"将清理 STD: {old.name}")
            else:
                for old in list(DST.iterdir()):
                    if old.is_file() and not old.name.startswith("."):
                        try:
                            _ensure_writable(old)
                        except OSError:
                            pass
                        old.unlink()
                        print(f"已清理 STD: {old.name}")
        fill_missing_latest_from_raw(dry_run=dry_run)
        if not dry_run:
            _print_std_now()
        return
    sync_cost_std_and_raw(dry_run=dry_run)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="STD 修改回写 RAW，再按标准项目名把各项目最新月份放进 STD（同项目只留一份）"
    )
    parser.add_argument(
        "--clear",
        action="store_true",
        help="回写 RAW 后清空 STD，再从 RAW 重建各标准项目的最新文件",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印将要复制、改名和移除的文件，不写磁盘",
    )
    args = parser.parse_args()
    copy_latest(clear_dst=args.clear, dry_run=args.dry_run)
