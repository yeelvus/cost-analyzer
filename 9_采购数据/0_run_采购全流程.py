"""
采购全流程一键运行：
  步骤1：*_rename_excel_month_filenames.py  —— 重命名原始文件（项目名_月份）
  步骤2：*_采购部单文件标准化输出.py        —— 逐文件标准化，写入 B_std
  步骤3：*_采购部数据汇总.py               —— 按月汇总 B_std，写入 B_sum
  步骤4：*_采购原始台账项目覆盖宽表.py      —— 生成项目×月份覆盖宽表，复制最新文件到 simplify

脚本按名称关键词查找，忽略编号前缀（如 1_ / 3_），编号变更后无需修改本文件。
"""

import importlib.util
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).parent

# 动态加载 _采购路径 模块以使用 find_script
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(CURRENT_DIR / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    raise ImportError(f"无法加载路径配置模块: {CURRENT_DIR / '_采购路径.py'}")
_路径_mod = importlib.util.module_from_spec(_路径_spec)
_路径_spec.loader.exec_module(_路径_mod)
find_script = _路径_mod.find_script


def load_main(keyword: str):
    script_path = find_script(keyword)
    if not script_path.exists():
        raise FileNotFoundError(f"未找到含关键词 '{keyword}' 的脚本: {CURRENT_DIR}")
    spec = importlib.util.spec_from_file_location(script_path.name, str(script_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载脚本模块: {script_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.main


def run_step(step: int, keyword: str):
    script_path = find_script(keyword)
    print("\n" + "=" * 100)
    print(f"步骤 {step}：{script_path.name}")
    print("=" * 100)
    main_func = load_main(keyword)
    # 隔离 sys.argv，避免 argparse 误读流水线参数
    _orig_argv = sys.argv
    sys.argv = [script_path.name]
    try:
        main_func()
    finally:
        sys.argv = _orig_argv
    print(f"\n步骤 {step} 完成：{script_path.name}")


if __name__ == "__main__":
    run_step(1, "rename_excel_month_filenames")
    run_step(2, "采购部单文件标准化输出")
    run_step(3, "采购部数据汇总")
    run_step(4, "采购原始台账项目覆盖宽表")
    print("\n" + "=" * 100)
    print("全流程完成")
    print("=" * 100)
