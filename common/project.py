from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

# code/ 根目录（common/project.py 的上两级）
_CODE_ROOT = Path(__file__).resolve().parents[1]

# 单例键：确保全进程只加载一次，所有调用方共享同一内存实例
_SINGLETON_KEY = "_project_name_matcher_singleton"


def _load_matcher():
    if _SINGLETON_KEY in sys.modules:
        return sys.modules[_SINGLETON_KEY]
    matcher_path = _CODE_ROOT / "0_项目基本信息" / "project_name_matcher.py"
    spec = importlib.util.spec_from_file_location(_SINGLETON_KEY, matcher_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载 project_name_matcher: {matcher_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[_SINGLETON_KEY] = module  # 注册后再执行，防止循环导入
    spec.loader.exec_module(module)
    return module


_matcher = _load_matcher()

build_standard_projects_from_dataframe = _matcher.build_standard_projects_from_dataframe
load_or_init_db = _matcher.load_or_init_db
save_db = _matcher.save_db
get_project_options = _matcher.get_project_options
get_project_lookup = _matcher.get_project_lookup
add_mapping = _matcher.add_mapping
mark_unmatched = _matcher.mark_unmatched
clear_unmatched_aliases = _matcher.clear_unmatched_aliases
is_unmatched_alias = _matcher.is_unmatched_alias
match_project = _matcher.match_project
match_project_name = _matcher.match_project
prompt_unmatched_project_mappings = _matcher.prompt_unmatched_project_mappings

__all__ = [
    "build_standard_projects_from_dataframe",
    "load_or_init_db",
    "save_db",
    "get_project_options",
    "get_project_lookup",
    "add_mapping",
    "mark_unmatched",
    "clear_unmatched_aliases",
    "is_unmatched_alias",
    "match_project",
    "match_project_name",
    "prompt_unmatched_project_mappings",
]
