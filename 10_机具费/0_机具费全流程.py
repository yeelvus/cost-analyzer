from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
STEP_FILES = [
    "1_月份标准化_使用时间优先.py",
    "2_机具使用汇总.py",
    "3_项目汇总机具费月度增量.py",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="机具费全流程执行：标准化 -> 汇总 -> 月度增量")
    parser.add_argument(
        "--python",
        default=sys.executable,
        help="用于执行子脚本的 Python 可执行文件路径（默认当前解释器）",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="传给第1步标准化脚本的输入文件路径（可选）",
    )
    parser.add_argument(
        "--suffix",
        default=None,
        help="传给第1步标准化脚本的输出后缀（可选）",
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续步骤（默认失败即停止）",
    )
    return parser.parse_args()


def build_command(python_exec: str, script_path: Path, step_index: int, args: argparse.Namespace) -> list[str]:
    command = [python_exec, str(script_path)]

    # 仅第1步支持输入文件与后缀参数透传。
    if step_index == 1:
        if args.input:
            command.extend(["--input", args.input])
        if args.suffix:
            command.extend(["--suffix", args.suffix])

    return command


def run_step(step_index: int, total_steps: int, script_name: str, command: list[str]) -> tuple[bool, float]:
    print(f"\n[{step_index}/{total_steps}] 开始执行: {script_name}")
    print("命令:", " ".join(command))

    start = time.time()
    completed = subprocess.run(command, cwd=str(SCRIPT_DIR))
    elapsed = time.time() - start

    if completed.returncode == 0:
        print(f"[{step_index}/{total_steps}] 完成: {script_name} (耗时 {elapsed:.1f}s)")
        return True, elapsed

    print(
        f"[{step_index}/{total_steps}] 失败: {script_name} "
        f"(退出码 {completed.returncode}, 耗时 {elapsed:.1f}s)"
    )
    return False, elapsed


def main() -> None:
    args = parse_args()

    python_exec = str(Path(args.python).expanduser())
    total_steps = len(STEP_FILES)
    durations: list[float] = []

    print("=== 机具费全流程开始 ===")
    print("执行目录:", SCRIPT_DIR)
    print("Python:", python_exec)

    for idx, script_name in enumerate(STEP_FILES, start=1):
        script_path = SCRIPT_DIR / script_name
        if not script_path.exists():
            raise FileNotFoundError(f"未找到脚本: {script_path}")

        command = build_command(python_exec, script_path, idx, args)
        success, elapsed = run_step(idx, total_steps, script_name, command)
        durations.append(elapsed)

        if not success and not args.continue_on_error:
            print("\n流程中止：请先修复失败步骤后重试。")
            raise SystemExit(1)

    print("\n=== 机具费全流程结束 ===")
    for idx, (script_name, elapsed) in enumerate(zip(STEP_FILES, durations), start=1):
        print(f"  {idx}. {script_name}: {elapsed:.1f}s")


if __name__ == "__main__":
    main()
