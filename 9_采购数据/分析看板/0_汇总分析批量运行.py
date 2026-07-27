"""汇总分析批量运行器
同时并发执行两个分析看板脚本：
  - 钢筋价格趋势分析.py  → C_数据分析/*.xlsx + *.html
  - 桩成本分析看板.py    → C_数据分析/*.xlsx + *.html
"""
from __future__ import annotations

import concurrent.futures
import subprocess
import sys
import time
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
PYTHON = sys.executable

SCRIPTS: list[Path] = [
    SCRIPT_DIR / "钢筋价格趋势分析.py",
    SCRIPT_DIR / "桩成本分析看板.py",
]


def run_script(script: Path) -> tuple[str, bool, float, str]:
    t0 = time.time()
    result = subprocess.run(
        [PYTHON, str(script)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed = time.time() - t0
    ok = result.returncode == 0
    output = result.stdout
    if result.stderr:
        output += f"\n[STDERR]\n{result.stderr}"
    return script.name, ok, elapsed, output


def main() -> None:
    print("=" * 60)
    print("采购数据分析看板 —— 批量并发运行")
    print(f"脚本目录: {SCRIPT_DIR}")
    print(f"运行脚本数: {len(SCRIPTS)}")
    print("=" * 60)

    missing = [s for s in SCRIPTS if not s.exists()]
    if missing:
        for s in missing:
            print(f"[错误] 脚本不存在: {s}")
        sys.exit(1)

    start = time.time()
    results: list[tuple[str, bool, float, str]] = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=len(SCRIPTS)) as pool:
        futures = {pool.submit(run_script, s): s for s in SCRIPTS}
        for future in concurrent.futures.as_completed(futures):
            results.append(future.result())

    total = time.time() - start
    print()
    all_ok = True
    for name, ok, elapsed, output in sorted(results, key=lambda x: x[0]):
        status = "✓ 成功" if ok else "✗ 失败"
        print(f"{'─' * 60}")
        print(f"{status}  [{name}]  耗时 {elapsed:.1f}s")
        print(output.rstrip())
        if not ok:
            all_ok = False

    print(f"{'=' * 60}")
    print(f"全部完成（并发总耗时 {total:.1f}s）")
    if not all_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
