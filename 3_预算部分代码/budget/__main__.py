# -*- coding: utf-8 -*-
"""
python -m budget [command]

commands:
  glodon         广联达导出 → 人材机 std + 总汇总 Excel
  dashboard      读取最新总汇总 → HTML 看板
  all            glodon + dashboard（默认）
  responsibility 责任成本处理
  units-test     运行单位换算自检
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _ensure_path():
    root = Path(__file__).resolve().parents[1]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    code_root = root.parent
    if str(code_root) not in sys.path:
        sys.path.insert(0, str(code_root))


def cmd_glodon(args):
    from budget.glodon.pipeline import run_glodon_pipeline

    run_glodon_pipeline(
        input_dir=args.input,
        output_dir=args.output,
        skip_summary=args.skip_summary,
    )


def cmd_dashboard(args):
    from budget.config import DASHBOARD_HTML_NAME, GLODON_STD_DIR
    from budget.dashboard import (
        build_html,
        build_payload,
        find_latest_summary_file,
        run_source_script,
    )

    output_dir = Path(args.output_dir or GLODON_STD_DIR)
    if getattr(args, "run_source", False):
        run_source_script()

    if getattr(args, "summary_file", None):
        summary_file = Path(args.summary_file)
        if not summary_file.exists():
            raise FileNotFoundError(f"指定总汇总不存在: {summary_file}")
    else:
        summary_file = find_latest_summary_file(output_dir)

    print(f"使用总汇总: {summary_file}")
    payload = build_payload(summary_file)
    html = build_html(payload)
    out_html = Path(getattr(args, "output", None) or (output_dir / DASHBOARD_HTML_NAME))
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(html, encoding="utf-8")
    print(f"看板已生成: {out_html}")


def cmd_all(args):
    cmd_glodon(args)

    class _DashArgs:
        run_source = False
        summary_file = None
        output_dir = args.output
        output = None

    cmd_dashboard(_DashArgs())


def cmd_responsibility(args):
    from budget.responsibility import main as resp_main

    resp_main()


def cmd_units_test(args):
    from budget.units import (
        normalize_unit,
        sum_quantities_to_base,
        to_base_quantity,
        unit_matches_material,
    )

    assert normalize_unit("吨") == "TON"
    assert normalize_unit("kg") == "KG"
    total, bad = sum_quantities_to_base([10, 500], ["吨", "kg"], "钢筋")
    assert abs(total - 10500) < 1e-6 and not bad
    assert unit_matches_material("吨", "钢筋")
    r = to_base_quantity(1, "袋", "钢筋")
    assert not r.ok
    print("units self-check OK (10吨+500kg=10500kg)")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="budget",
        description="预算处理工具集（广联达人材机 / 看板 / 责任成本）",
    )
    sub = p.add_subparsers(dest="command")

    g = sub.add_parser("glodon", help="广联达 raw → std + 总汇总")
    g.add_argument("-i", "--input", default=None, help="raw 输入目录")
    g.add_argument("-o", "--output", default=None, help="std 输出目录")
    g.add_argument("--skip-summary", action="store_true", help="不生成跨项目总汇总")
    g.set_defaults(func=cmd_glodon)

    d = sub.add_parser("dashboard", help="生成 HTML 看板")
    d.add_argument("--run-source", action="store_true", help="先跑广联达 pipeline")
    d.add_argument("--summary-file", default=None, help="指定总汇总 xlsx")
    d.add_argument("--output-dir", default=None, help="总汇总所在目录")
    d.add_argument("-o", "--output", default=None, help="输出 HTML 路径")
    d.set_defaults(func=cmd_dashboard)

    a = sub.add_parser("all", help="广联达处理 + 看板（默认）")
    a.add_argument("-i", "--input", default=None)
    a.add_argument("-o", "--output", default=None)
    a.add_argument("--skip-summary", action="store_true")
    a.set_defaults(func=cmd_all)

    r = sub.add_parser("responsibility", help="责任成本处理")
    r.set_defaults(func=cmd_responsibility)

    u = sub.add_parser("units-test", help="单位换算自检")
    u.set_defaults(func=cmd_units_test)

    return p


def main(argv=None):
    _ensure_path()
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.command:
        # 默认全流程
        args = parser.parse_args(["all"])
    args.func(args)


if __name__ == "__main__":
    main()
