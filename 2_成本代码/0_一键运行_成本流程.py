#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
SCRIPT_SEQUENCE = [
    BASE_DIR / "成本_汇总按月收集成本统计.py",
    BASE_DIR / "成本_数据分析.py",
    BASE_DIR / "预计毛利率对比_数据统计与环比.py",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="一键顺序运行成本统计、成本分析和预计毛利率对比脚本"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本",
    )
    return parser.parse_args()


def run_script(script_path: Path) -> None:
    print(f"\n{'=' * 72}")
    print(f"[开始] {script_path.name}")
    print(f"{'=' * 72}")
    subprocess.run(
        [sys.executable, str(script_path)],
        check=True,
        cwd=BASE_DIR,
    )
    print(f"[完成] {script_path.name}")


def main() -> None:
    args = parse_args()

    missing_scripts = [path for path in SCRIPT_SEQUENCE if not path.exists()]
    if missing_scripts:
        missing_text = "\n".join(str(path) for path in missing_scripts)
        raise FileNotFoundError(f"以下脚本不存在:\n{missing_text}")

    failed_scripts: list[tuple[Path, int]] = []
    for script_path in SCRIPT_SEQUENCE:
        try:
            run_script(script_path)
        except subprocess.CalledProcessError as exc:
            failed_scripts.append((script_path, exc.returncode))
            print(f"[失败] {script_path.name}，退出码: {exc.returncode}")
            if not args.continue_on_error:
                raise SystemExit(exc.returncode) from exc

    print(f"\n{'=' * 72}")
    if failed_scripts:
        print("全部流程已执行，但存在失败步骤:")
        for script_path, code in failed_scripts:
            print(f"- {script_path.name}: exit code {code}")
        raise SystemExit(1)

    print("全部成本流程执行完成")
    for script_path in SCRIPT_SEQUENCE:
        print(f"- {script_path.name}")
    print(f"{'=' * 72}")


if __name__ == "__main__":
    main()