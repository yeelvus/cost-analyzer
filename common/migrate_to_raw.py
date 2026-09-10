# -*- coding: utf-8 -*-
"""把各领域 *_raw 原始文件迁到 01成本统计/0_Raw/。

用法：
  python3 -m common.migrate_to_raw
  python3 -m common.migrate_to_raw --dry-run
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import DATA_ROOT, dir_key, ensure_raw_root, resolve_under
from settings import DATA_ROOT as _DATA_ROOT

README_TEXT = """此目录中的原始文件已迁到：

  {dest}

以后请把新收到的原始资料放到 0_Raw（也可写成 Raw）下对应目录。
目录名可以带 0_、1_ 这类序号，程序会自动忽略。
"""


def _move_contents(src: Path, dst: Path, *, dry_run: bool) -> int:
    """把 src 下条目移到 dst（同名则跳过文件、合并目录）。"""
    if not src.is_dir():
        return 0
    dst.mkdir(parents=True, exist_ok=True)
    moved = 0
    for item in list(src.iterdir()):
        if item.name.startswith("~$"):
            continue
        if item.name.startswith("已迁移"):
            continue
        target = dst / item.name
        if dry_run:
            print(f"  DRY  {item}  →  {target}")
            moved += 1
            continue
        if target.exists():
            if item.is_dir() and target.is_dir():
                moved += _move_contents(item, target, dry_run=False)
                try:
                    item.rmdir()
                except OSError:
                    pass
            else:
                print(f"  skip exists: {target}")
            continue
        shutil.move(str(item), str(target))
        moved += 1
    return moved


def _leave_pointer(src: Path, dest: Path, *, dry_run: bool) -> None:
    if dry_run or not src.is_dir():
        return
    note = src / "已迁移到_0_Raw.txt"
    try:
        note.write_text(README_TEXT.format(dest=dest), encoding="utf-8")
    except OSError as exc:
        print(f"  ⚠️ 无法写说明: {exc}")


def _find_existing(*candidates: Path) -> Path | None:
    for p in candidates:
        if p.is_dir():
            return p
    return None


def _domain_dir(keyword: str) -> Path | None:
    if not DATA_ROOT.is_dir():
        return None
    for child in DATA_ROOT.iterdir():
        if child.is_dir() and dir_key(child.name) == dir_key(keyword):
            return child
    for child in DATA_ROOT.iterdir():
        if child.is_dir() and keyword in child.name:
            return child
    return None


def planned_moves() -> list[tuple[str, Path, tuple[str, ...]]]:
    """(说明, 源目录, raw 下的逻辑段)."""
    cost = _domain_dir("成本统计收集")
    budget = _domain_dir("预算部分")
    resp = _domain_dir("责任成本")
    sub = _domain_dir("分包合同")
    proc = _domain_dir("采购台账")
    mix = _domain_dir("商混站")
    debt = _domain_dir("材料欠款")
    machine = _domain_dir("机具费")
    labor = _domain_dir("人工费")
    finance = _domain_dir("财务数据") or _domain_dir("财务付款")

    jobs: list[tuple[str, Path | None, tuple[str, ...]]] = [
        ("成本统计收集", (cost / "A_成本统计_raw") if cost else None, ("成本统计收集",)),
        ("广联达人材机", (budget / "1_广联达导出人材料机_raw") if budget else None, ("预算部分",)),
        ("责任成本", (resp / "1_责任成本_raw") if resp else None, ("责任成本",)),
        ("分包合同台账", (sub / "1_分包合同台账_raw") if sub else None, ("分包合同",)),
        ("采购台账", (proc / "A_采购部采购台账_raw") if proc else None, ("采购台账",)),
        ("商混站", (mix / "A_商混站_raw") if mix else None, ("商混站",)),
        ("材料欠款", (debt / "A_材料欠款_raw") if debt else None, ("材料欠款",)),
        ("机具费", (machine / "A_机具费_raw") if machine else None, ("机具费",)),
        (
            "人工费台账",
            (labor / "A_1_人工费台账" / "A_人工费台账_raw") if labor else None,
            ("人工费", "人工费台账"),
        ),
        (
            "机票",
            (labor / "A_2_飞签机票" / "A_飞签台账_raw") if labor else None,
            ("人工费", "机票"),
        ),
        (
            "财务付款",
            (finance / "A_财务付款数据" / "A_财务付款_raw") if finance else None,
            ("财务付款",),
        ),
    ]
    result = []
    for title, src, parts in jobs:
        if src is None or not src.is_dir():
            continue
        if (src / "已迁移到_0_Raw.txt").exists() and not any(
            p.name != "已迁移到_0_Raw.txt" and not p.name.startswith("~$")
            for p in src.iterdir()
        ):
            continue
        result.append((title, src, parts))
    return result


def run(*, dry_run: bool = False) -> None:
    raw_root = DATA_ROOT / "0_Raw" if dry_run else ensure_raw_root()
    # 空的 Raw/ 升为 0_Raw
    if not dry_run:
        plain = DATA_ROOT / "Raw"
        numbered = DATA_ROOT / "0_Raw"
        if plain.is_dir() and not numbered.exists():
            try:
                empty = not any(plain.iterdir())
            except OSError:
                empty = False
            if empty:
                plain.rename(numbered)
                raw_root = numbered
                print(f"已将 Raw/ 重命名为 {numbered}")
        raw_root = ensure_raw_root()

    print(f"DATA_ROOT = {DATA_ROOT}")
    print(f"RAW_ROOT  = {raw_root}")
    jobs = planned_moves()
    if not jobs:
        print("没有需要迁移的 raw 目录（可能已经迁过）。")
        return

    total = 0
    for title, src, parts in jobs:
        dest = resolve_under(raw_root, *parts, create=not dry_run)
        print(f"\n[{title}]")
        print(f"  from {src}")
        print(f"  to   {dest}")
        n = _move_contents(src, dest, dry_run=dry_run)
        total += n
        print(f"  条目 {n}")
        if not dry_run:
            _leave_pointer(src, dest, dry_run=False)
    print(f"\n完成，共处理 {total} 个条目。dry_run={dry_run}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="迁移原始数据到 0_Raw")
    parser.add_argument("--dry-run", action="store_true", help="只打印，不移动")
    args = parser.parse_args(argv)
    if DATA_ROOT != _DATA_ROOT:
        pass
    run(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
