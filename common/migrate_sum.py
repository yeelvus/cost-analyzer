# -*- coding: utf-8 -*-
"""把各领域合并汇总 Excel 迁到 01成本统计/0_Sum。

用法：
  python3 -m common.migrate_sum
  python3 -m common.migrate_sum --dry-run
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import DATA_ROOT, dir_key, ensure_sum_root, resolve_under
from settings import REPORT_ROOT


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


def _move_file(src: Path, dst: Path, *, dry_run: bool) -> int:
    if not src.exists() or not src.is_file():
        return 0
    if src.name.startswith("~$"):
        return 0
    if dry_run:
        print(f"  DRY  {src}  →  {dst}")
        return 1
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists() and dst.resolve() == src.resolve():
        return 0
    if dst.exists():
        print(f"  skip exists: {dst.name}")
        return 0
    shutil.move(str(src), str(dst))
    return 1


def _move_xlsx_in_dir(src_dir: Path, dst_dir: Path, *, dry_run: bool, names: set[str] | None = None) -> int:
    if not src_dir.is_dir():
        return 0
    n = 0
    for p in src_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            continue
        if names and p.name not in names and not any(p.name.startswith(x.replace(".xlsx", "")) for x in (names or [])):
            # if names given, only those exact names
            if p.name not in names:
                continue
        n += _move_file(p, dst_dir / p.name, dry_run=dry_run)
    return n


def run(*, dry_run: bool = False) -> None:
    sum_root = DATA_ROOT / "0_Sum" if dry_run else ensure_sum_root()
    print(f"SUM_ROOT = {sum_root}")
    n = 0

    jobs: list[tuple[str, Path, str]] = []

    cost = _domain("成本统计收集")
    if cost:
        jobs.append(("成本统计收集", cost / "B_成本统计_sum" / "项目统计数据汇总_全月份.xlsx", "2_成本统计收集"))

    budget = _domain("预算部分")
    if budget:
        jobs.append(("预算总汇总", budget / "2_人材机_std" / "【总汇总】所有标段.xlsx", "3_预算部分"))

    resp = _domain("责任成本")
    if resp:
        jobs.append(("责任成本汇总", resp / "2_责任成本_std" / "责任成本汇总.xlsx", "3.1_责任成本"))

    sub = _domain("分包合同")
    if sub:
        d = sub / "3_分包合同_sum"
        dest = resolve_under(sum_root, "分包合同", create=not dry_run)
        print(f"\n[分包合同] {d} → {dest}")
        n += _move_xlsx_in_dir(d, dest, dry_run=dry_run)

    proc = _domain("采购台账")
    if proc:
        dest = resolve_under(sum_root, "采购台账", create=not dry_run)
        for subdir in ("E_采购部采购台账_sum", "D_采购部采购台账_raw_sum"):
            src = proc / subdir
            print(f"\n[采购 {subdir}] {src} → {dest}")
            n += _move_xlsx_in_dir(src, dest, dry_run=dry_run)

    mix = _domain("商混站")
    if mix:
        dest = resolve_under(sum_root, "商混站", create=not dry_run)
        print(f"\n[商混站] {mix / 'B_商混站_std'} → {dest}")
        n += _move_xlsx_in_dir(mix / "B_商混站_std", dest, dry_run=dry_run)
        if REPORT_ROOT.is_dir():
            for name in ("搅拌站混凝土汇总报表.xlsx", "搅拌站混凝土汇总报表_仅有项目编码.xlsx"):
                n += _move_file(REPORT_ROOT / name, dest / name, dry_run=dry_run)

    debt = _domain("材料欠款")
    if debt:
        dest = resolve_under(sum_root, "材料欠款", create=not dry_run)
        print(f"\n[材料欠款] → {dest}")
        n += _move_xlsx_in_dir(debt / "B_材料欠款_std", dest, dry_run=dry_run)

    machine = _domain("机具费")
    if machine:
        dest = resolve_under(sum_root, "机具费", create=not dry_run)
        print(f"\n[机具费] → {dest}")
        n += _move_xlsx_in_dir(machine / "D_机具费_std_sum", dest, dry_run=dry_run)

    labor = _domain("人工费")
    if labor:
        dest = resolve_under(sum_root, "人工费", create=not dry_run)
        print(f"\n[人工费] → {dest}")
        n += _move_file(
            labor / "A_1_人工费台账" / "B_人工费台账_std" / "人工费汇总.xlsx",
            dest / "人工费汇总.xlsx",
            dry_run=dry_run,
        )
        n += _move_file(
            labor / "A_2_飞签机票" / "B_飞签机票_std" / "机票汇总.xlsx",
            dest / "机票汇总.xlsx",
            dry_run=dry_run,
        )
        if REPORT_ROOT.is_dir():
            n += _move_file(REPORT_ROOT / "人工费汇总.xlsx", dest / "人工费汇总.xlsx", dry_run=dry_run)

    finance = _domain("财务数据") or _domain("财务付款")
    if finance:
        src = finance / "A_财务付款数据" / "C_财务付款_sum"
        dest = resolve_under(sum_root, "财务付款", create=not dry_run)
        print(f"\n[财务付款] {src} → {dest}")
        if src.is_dir():
            # 组目录下的汇总 xlsx
            for p in src.rglob("*.xlsx"):
                if p.name.startswith("~$"):
                    continue
                rel = p.relative_to(src)
                n += _move_file(p, dest / rel, dry_run=dry_run)

    contract = _domain("合同管理")
    if contract:
        dest = resolve_under(sum_root, "合同管理", create=not dry_run)
        print(f"\n[合同管理] → {dest}")
        n += _move_file(
            contract / "合同管理_总包分包全量汇总.xlsx",
            dest / "合同管理_总包分包全量汇总.xlsx",
            dry_run=dry_run,
        )

    for title, src, part in jobs:
        dest_dir = resolve_under(sum_root, part, create=not dry_run)
        print(f"\n[{title}] {src.name} → {dest_dir}")
        n += _move_file(src, dest_dir / src.name, dry_run=dry_run)

    print(f"\n完成，共处理 {n} 个文件。dry_run={dry_run}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="迁移合并汇总 Excel 到 0_Sum")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    run(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
