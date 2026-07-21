#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据分析模块统一入口。

用法：
  python 00_分析入口.py --list
  python 00_分析入口.py dashboard lifecycle
  python 00_分析入口.py dashboard unified
  python 00_分析入口.py dashboard comprehensive
  python 00_分析入口.py dashboard budget-vs-process
  python 00_分析入口.py analysis cost
  python 00_分析入口.py analysis margin
  python 00_分析入口.py analysis subpackage-payment
  python 00_分析入口.py reports generate-config
"""

from __future__ import annotations

import argparse
import runpy
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_CODE = _HERE.parent
if str(_CODE) not in sys.path:
    sys.path.insert(0, str(_CODE))

# name -> relative path under S_数据分析模块
COMMANDS: dict[str, dict[str, str]] = {
    "dashboard": {
        "lifecycle": "dashboards/build_full_lifecycle_dashboard.py",
        "unified": "dashboards/build_unified_cost_dashboard.py",
        "kanban": "dashboards/build_cost_kanban.py",
        "comprehensive": "dashboards/build_comprehensive_cost_analysis_html.py",
        "company-monthly": "dashboards/build_company_cost_monthly_html.py",
        "budget-vs-process": "dashboards/build_budget_vs_process.py",
    },
    "analysis": {
        "cost": "analysis/成本_数据分析.py",
        "margin": "analysis/预计毛利率对比_数据统计与环比.py",
        "subpackage-payment": "analysis/3_分包付款汇总分析_按日期与月增量.py",
    },
    "reports": {
        "generate-config": "reports/generate_report_config.py",
        "pipeline": "reports/contract_report_pipeline.py",
    },
}


def list_commands() -> None:
    print("可用命令：\n")
    for group, items in COMMANDS.items():
        print(f"  [{group}]")
        for name, rel in items.items():
            print(f"    {name:22s}  →  {rel}")
        print()
    print("示例: python 00_分析入口.py dashboard unified -- --month 2026年06月")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="S_数据分析模块入口（看板 / 分析 / 报告）",
    )
    parser.add_argument("--list", action="store_true", help="列出全部子命令")
    parser.add_argument(
        "group",
        nargs="?",
        choices=list(COMMANDS.keys()),
        help="命令分组: dashboard | analysis | reports",
    )
    parser.add_argument("name", nargs="?", help="分组内脚本名（见 --list）")
    parser.add_argument(
        "script_args",
        nargs=argparse.REMAINDER,
        help="传给目标脚本的参数；建议用 -- 分隔",
    )
    args = parser.parse_args(argv)

    if args.list or not args.group:
        list_commands()
        if not args.list and not args.group:
            return
        if args.list:
            return

    group = COMMANDS[args.group]
    if not args.name or args.name not in group:
        print(f"请指定 {args.group} 下的脚本名：")
        for n in group:
            print(f"  - {n}")
        raise SystemExit(2)

    target = _HERE / group[args.name]
    if not target.is_file():
        raise SystemExit(f"脚本不存在: {target}")

    # 去掉前导 --
    extra = list(args.script_args or [])
    if extra and extra[0] == "--":
        extra = extra[1:]

    sys.argv = [str(target), *extra]
    # 在 code 根执行上下文
    runpy.run_path(str(target), run_name="__main__")


if __name__ == "__main__":
    main()
