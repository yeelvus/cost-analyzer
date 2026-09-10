# -*- coding: utf-8 -*-
"""把已有 HTML 从 Z_report 迁到 01成本统计/0_html。

用法：
  python3 -m common.migrate_html
  python3 -m common.migrate_html --dry-run
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.paths import DATA_ROOT, ensure_html_root
from settings import REPORT_ROOT


def _move(src: Path, dst: Path, *, dry_run: bool) -> int:
    if not src.exists():
        return 0
    if dry_run:
        print(f"  DRY  {src}  →  {dst}")
        return 1
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        if src.is_dir() and dst.is_dir():
            n = 0
            for child in src.iterdir():
                n += _move(child, dst / child.name, dry_run=False)
            return n
        print(f"  skip exists: {dst}")
        return 0
    shutil.move(str(src), str(dst))
    return 1


def run(*, dry_run: bool = False) -> None:
    html_root = DATA_ROOT / "0_html" if dry_run else ensure_html_root()
    print(f"HTML_ROOT = {html_root}")
    print(f"Z_report  = {REPORT_ROOT}")
    if not REPORT_ROOT.is_dir():
        print("Z_report 不存在，无需迁移。")
        return

    n = 0
    # 根目录 html
    for p in REPORT_ROOT.iterdir():
        if p.is_file() and p.suffix.lower() == ".html":
            n += _move(p, html_root / p.name, dry_run=dry_run)

    # 看板整夹（含 json/js）
    dash = REPORT_ROOT / "综合成本看板"
    if dash.is_dir():
        n += _move(dash, html_root / "综合成本看板", dry_run=dry_run)

    # 机具费 html（xlsx 留在 Z_report）
    mach = REPORT_ROOT / "机具费分析"
    if mach.is_dir():
        dest = html_root / "机具费分析"
        for p in mach.glob("*.html"):
            n += _move(p, dest / p.name, dry_run=dry_run)

    # 财务付款 html
    fin = REPORT_ROOT / "财务付款"
    if fin.is_dir():
        dest = html_root / "财务付款"
        for p in fin.glob("*.html"):
            n += _move(p, dest / p.name, dry_run=dry_run)

    # 采购分析 HTML 子目录
    proc_html = REPORT_ROOT / "C_采购数据分析" / "HTML"
    if proc_html.is_dir():
        dest = html_root / "采购数据分析" / "HTML"
        for p in proc_html.glob("*.html"):
            n += _move(p, dest / p.name, dry_run=dry_run)

    print(f"\n完成，共处理 {n} 个条目。dry_run={dry_run}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="迁移 HTML 到 0_html")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    run(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
