"""
采购全流程一键运行：
  步骤1：1_rename_excel_month_filenames.py  —— 重命名原始文件（项目名_月份）
  步骤2：2_采购部单文件标准化输出.py        —— 逐文件标准化，写入 B_std
  步骤3：3_采购部数据汇总.py               —— 按月汇总 B_std，写入 B_sum
  步骤4：2_采购原始台账项目覆盖宽表.py      —— 生成项目×月份覆盖宽表，复制最新文件到 simplify
"""

import importlib.util
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).parent


def load_main(script_name: str):
    script_path = CURRENT_DIR / script_name
    spec = importlib.util.spec_from_file_location(script_name, str(script_path))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.main


def run_step(step: int, script_name: str):
    print("\n" + "=" * 100)
    print(f"步骤 {step}：{script_name}")
    print("=" * 100)
    main_func = load_main(script_name)
    # 隔离 sys.argv，避免 argparse 误读流水线参数
    _orig_argv = sys.argv
    sys.argv = [script_name]
    try:
        main_func()
    finally:
        sys.argv = _orig_argv
    print(f"\n步骤 {step} 完成：{script_name}")


if __name__ == "__main__":
    run_step(1, "1_rename_excel_month_filenames.py")
    run_step(2, "2_采购部单文件标准化输出.py")
    run_step(3, "3_采购部数据汇总.py")
    run_step(4, "2_采购原始台账项目覆盖宽表.py")
    print("\n" + "=" * 100)
    print("全流程完成")
    print("=" * 100)
