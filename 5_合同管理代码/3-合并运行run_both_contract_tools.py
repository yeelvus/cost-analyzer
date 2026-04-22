import subprocess
import sys
from pathlib import Path


def run_script(script_path: Path) -> None:
    print(f"\n=== Running: {script_path.name} ===")
    result = subprocess.run([sys.executable, str(script_path)])
    if result.returncode != 0:
        raise SystemExit(f"Script failed: {script_path.name} (exit {result.returncode})")


def pick_first_existing(base_dir: Path, candidates: list[str], label: str) -> Path:
    """从候选脚本名中选择第一个存在的文件。"""
    for name in candidates:
        path = base_dir / name
        if path.exists():
            print(f"Using {label}: {path.name}")
            return path

    joined = " | ".join(candidates)
    raise SystemExit(f"Missing {label} script, tried: {joined}")


def main() -> None:
    base_dir = Path(__file__).resolve().parent

    script_1 = pick_first_existing(
        base_dir,
        [
            "03_合同管理-汇总.py",
        ],
        "step-1",
    )
    script_2 = pick_first_existing(
        base_dir,
        [
            "04-总包分包合同额统计.py",
            "contract_summary.py",
        ],
        "step-2",
    )

    run_script(script_1)
    run_script(script_2)

    print("\n=== All scripts completed successfully ===")


if __name__ == "__main__":
    main()
