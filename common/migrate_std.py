# -*- coding: utf-8 -*-
"""把各领域标准化中间结果迁到 01成本统计/0_Std。

用法：
  python3 -m common.migrate_std
  python3 -m common.migrate_std --dry-run
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import DATA_ROOT, dir_key, ensure_std_root, resolve_under

README_TEXT = """此目录中的标准化中间文件已迁到：

  {dest}

以后由程序写入 0_Std（也可写成 Std）。目录名可以带 0_、1_ 这类序号，程序会自动忽略。
"""

# 汇总表已在 0_Sum，迁 std 时跳过
_SKIP_NAMES = {
    "【总汇总】所有标段.xlsx",
    "责任成本汇总.xlsx",
    "完整处理报告.xlsx",
    "人工费汇总.xlsx",
    "机票汇总.xlsx",
    "分包合同台账_全量汇总.xlsx",
    "分包合同汇总逐月对比汇总.xlsx",
}


def _domain(keyword: str) -> Path | None:
    if not DATA_ROOT.is_dir():
        return None
    for child in DATA_ROOT.iterdir():
        if child.is_dir() and dir_key(child.name) == dir_key(keyword):
            return child
    for child in DATA_ROOT.iterdir():
        if child.is_dir() and keyword in child.name:
            return child
    return None


def _move_item(src: Path, dst: Path, *, dry_run: bool) -> int:
    if not src.exists():
        return 0
    if src.name.startswith("~$") or src.name.startswith("."):
        return 0
    if src.name.startswith("已迁移"):
        return 0
    if src.name in _SKIP_NAMES:
        print(f"  skip sum file: {src.name}")
        return 0
    if dry_run:
        print(f"  DRY  {src}  →  {dst}")
        return 1
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        if src.is_dir() and dst.is_dir():
            n = 0
            for child in list(src.iterdir()):
                n += _move_item(child, dst / child.name, dry_run=False)
            try:
                src.rmdir()
            except OSError:
                pass
            return n
        print(f"  skip exists: {dst}")
        return 0
    shutil.move(str(src), str(dst))
    return 1


def _move_contents(src: Path, dst: Path, *, dry_run: bool) -> int:
    if not src.is_dir():
        return 0
    n = 0
    for item in list(src.iterdir()):
        n += _move_item(item, dst / item.name, dry_run=dry_run)
    return n


def _leave_pointer(src: Path, dest: Path, *, dry_run: bool) -> None:
    if dry_run or not src.is_dir():
        return
    note = src / "已迁移到_0_Std.txt"
    try:
        note.write_text(README_TEXT.format(dest=dest), encoding="utf-8")
    except OSError as exc:
        print(f"  ⚠️ 无法写说明: {exc}")


def _job(title: str, src: Path, dest: Path, *, dry_run: bool) -> int:
    if not src.exists():
        print(f"\n[{title}] 跳过（不存在）{src}")
        return 0
    print(f"\n[{title}] {src}  →  {dest}")
    n = _move_contents(src, dest, dry_run=dry_run) if src.is_dir() else _move_item(src, dest / src.name, dry_run=dry_run)
    if n and src.is_dir():
        _leave_pointer(src, dest, dry_run=dry_run)
    print(f"  处理 {n} 项")
    return n


def run(*, dry_run: bool = False) -> None:
    std_root = DATA_ROOT / "0_Std" if dry_run else ensure_std_root()
    print(f"STD_ROOT = {std_root}")
    n = 0

    cost = _domain("成本统计收集")
    if cost:
        n += _job(
            "成本统计收集",
            cost / "B_成本统计_std",
            resolve_under(std_root, "成本统计收集", create=not dry_run),
            dry_run=dry_run,
        )

    budget = _domain("预算部分")
    if budget:
        n += _job(
            "预算人材机",
            budget / "2_人材机_std",
            resolve_under(std_root, "预算部分", create=not dry_run),
            dry_run=dry_run,
        )

    resp = _domain("责任成本")
    if resp:
        n += _job(
            "责任成本",
            resp / "2_责任成本_std",
            resolve_under(std_root, "责任成本", create=not dry_run),
            dry_run=dry_run,
        )

    sub = _domain("分包合同")
    if sub:
        n += _job(
            "分包合同",
            sub / "2_分包合同台账_std",
            resolve_under(std_root, "分包合同", create=not dry_run),
            dry_run=dry_run,
        )

    proc = _domain("采购台账")
    if proc:
        dest_base = resolve_under(std_root, "采购台账", create=not dry_run)
        for folder in (
            "B_采购部采购台账_extract",
            "C_采购部采购台账_std",
            "D_采购部采购台账_raw_sum",
            "F_采购部采购台账_std_split",
        ):
            n += _job(
                f"采购 {folder}",
                proc / folder,
                dest_base / folder,
                dry_run=dry_run,
            )

    machine = _domain("机具费")
    if machine:
        dest_base = resolve_under(std_root, "机具费", create=not dry_run)
        for folder in ("B_机具费_raw_sum", "C_机具费_raw_split", "E_机具费_std_split"):
            n += _job(
                f"机具 {folder}",
                machine / folder,
                dest_base / folder,
                dry_run=dry_run,
            )

    labor = _domain("人工费")
    if labor:
        dest_base = resolve_under(std_root, "人工费", create=not dry_run)
        n += _job(
            "人工费台账",
            labor / "A_1_人工费台账" / "B_人工费台账_std",
            dest_base / "B_人工费台账_std",
            dry_run=dry_run,
        )
        n += _job(
            "机票",
            labor / "A_2_飞签机票" / "B_飞签机票_std",
            dest_base / "B_飞签机票_std",
            dry_run=dry_run,
        )

    finance = _domain("财务数据") or _domain("财务付款")
    if finance:
        pay = finance / "A_财务付款数据"
        dest_base = resolve_under(std_root, "财务付款", create=not dry_run)
        for folder in (
            "A1_财务付款_raw_std",
            "A2_财务付款_classified",
            "B_财务付款_std",
            "D_财务付款_学习池",
        ):
            n += _job(
                f"财务 {folder}",
                pay / folder,
                dest_base / folder,
                dry_run=dry_run,
            )

    print(f"\n完成，共处理 {n} 项。dry_run={dry_run}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="迁移标准化中间结果到 0_Std")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    run(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
