from __future__ import annotations

import argparse
import importlib.util
import re
from pathlib import Path
from types import ModuleType

from openpyxl import load_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
SUMMARY_INPUT_ROOT = SCRIPT_DIR.parent.parent / "9_采购台账📦/B_采购部采购台账_sum"
ANALYSIS_OUTPUT_ROOT = SCRIPT_DIR.parent.parent / "C_数据分析"
INPUT_GLOB = "*_材料采购_merged.xlsx"
DEFAULT_OUTPUT_SUFFIX = "_标准化汇总"
DEFAULT_RULES = ["桩规格标准化.py", "钢筋规格标准化.py","钢筋网片归类修正.py"]


def extract_year_month_from_name(path: Path) -> tuple[int, int] | None:
    m = re.search(r"(?<!\d)((?:19|20)\d{2})年(0?[1-9]|1[0-2])月", path.stem)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def find_latest_summary_file() -> Path | None:
    if not SUMMARY_INPUT_ROOT.is_dir():
        return None
    candidates = [p for p in SUMMARY_INPUT_ROOT.glob(INPUT_GLOB) if p.is_file()]
    if not candidates:
        return None

    def sort_key(p: Path):
        ym = extract_year_month_from_name(p)
        if ym is not None:
            return (1, ym[0], ym[1], p.stat().st_mtime)
        return (0, 0, 0, p.stat().st_mtime)

    return max(candidates, key=sort_key)


def build_output_path(input_file: Path, custom_output: str, suffix: str) -> Path:
    if custom_output:
        return Path(custom_output).expanduser().resolve()
    ANALYSIS_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    return ANALYSIS_OUTPUT_ROOT / f"{input_file.stem}{suffix}{input_file.suffix}"


def load_module_from_file(rule_file: Path) -> ModuleType:
    if not rule_file.exists():
        raise FileNotFoundError(f"未找到规则文件: {rule_file}")

    module_name = f"rule_{rule_file.stem}_{abs(hash(str(rule_file)))}"
    spec = importlib.util.spec_from_file_location(module_name, rule_file)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载规则模块: {rule_file}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_rule_on_workbook(module: ModuleType, workbook, rule_label: str) -> dict[str, int]:
    if not hasattr(module, "process_worksheet"):
        raise AttributeError(f"规则 {rule_label} 缺少 process_worksheet(worksheet) 函数")

    process_worksheet = getattr(module, "process_worksheet")

    stats = {
        "sheets": 0,
        "matched": 0,
        "updated": 0,
        "prefixed": 0,
        "unresolved": 0,
        "skipped": 0,
    }

    for ws in workbook.worksheets:
        result = process_worksheet(ws)
        stats["sheets"] += 1

        if result.get("skipped"):
            stats["skipped"] += 1
            continue

        stats["matched"] += int(result.get("matched_rows", 0))
        stats["updated"] += int(result.get("updated_rows", 0))
        stats["prefixed"] += int(result.get("prefixed_rows", 0))
        stats["unresolved"] += len(result.get("unresolved_rows", []))

    return stats


def parse_rules(rules_arg: str) -> list[Path]:
    if not rules_arg.strip():
        names = DEFAULT_RULES
    else:
        names = [x.strip() for x in rules_arg.split(",") if x.strip()]

    if not names:
        raise ValueError("未提供任何规则文件")

    paths: list[Path] = []
    for name in names:
        p = Path(name).expanduser()
        if not p.is_absolute():
            p = SCRIPT_DIR / p
        paths.append(p.resolve())
    return paths


def parse_args():
    parser = argparse.ArgumentParser(description="一次性串行执行多个标准化规则，仅输出一个最终Excel")
    parser.add_argument(
        "input_file",
        nargs="?",
        default="",
        help="输入 Excel 文件；不传时自动读取 B_汇总数据 下最新 *_材料采购_merged.xlsx",
    )
    parser.add_argument(
        "--rules",
        default=",".join(DEFAULT_RULES),
        help="按顺序执行的规则文件列表（逗号分隔），例如: 桩规格标准化.py,钢筋规格标准化.py",
    )
    parser.add_argument(
        "--suffix",
        default=DEFAULT_OUTPUT_SUFFIX,
        help="默认输出后缀（仅当未传 --output-file 时生效）",
    )
    parser.add_argument(
        "--output-file",
        default="",
        help="自定义输出文件路径；不传时输出到 C_数据分析",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.input_file:
        input_file = Path(args.input_file).expanduser().resolve()
    else:
        latest = find_latest_summary_file()
        if latest is None:
            raise FileNotFoundError(f"未在 {SUMMARY_INPUT_ROOT} 找到 {INPUT_GLOB}")
        input_file = latest.resolve()

    if not input_file.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    output_file = build_output_path(input_file, args.output_file, args.suffix)
    rule_files = parse_rules(args.rules)

    print(f"输入文件: {input_file}")
    print("规则顺序:")
    for i, rf in enumerate(rule_files, start=1):
        print(f"  {i}. {rf.name}")

    workbook = load_workbook(input_file)

    summary: list[tuple[str, dict[str, int]]] = []
    for rf in rule_files:
        module = load_module_from_file(rf)
        stats = run_rule_on_workbook(module, workbook, rf.name)
        summary.append((rf.name, stats))
        print(
            f"规则 {rf.name}: 命中 {stats['matched']} 行, 更新 {stats['updated']} 行, "
            f"名称加前缀 {stats['prefixed']} 行, 未识别 {stats['unresolved']} 行"
        )

    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)

    print("\n处理完成（单文件输出）")
    print(f"输出文件: {output_file}")

    print("\n执行汇总:")
    for name, stats in summary:
        print(
            f"- {name}: 命中={stats['matched']}, 更新={stats['updated']}, "
            f"前缀={stats['prefixed']}, 未识别={stats['unresolved']}, 跳过sheet={stats['skipped']}"
        )


if __name__ == "__main__":
    main()
