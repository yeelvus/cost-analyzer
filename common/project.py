from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

# code/ 根目录（common/project.py 的上两级）
_CODE_ROOT = Path(__file__).resolve().parents[1]

# 兼容新旧目录名：优先选择真正包含 project_name_matcher.py 的目录，
# 以避免旧目录存在但缺失核心文件时被错误选中。
_PROJECT_INFO_CANDIDATES = [
    _CODE_ROOT / "00_项目基本信息",
    _CODE_ROOT / "0_项目基本信息",
]

# 单例键：确保全进程只加载一次，所有调用方共享同一内存实例
_SINGLETON_KEY = "_project_name_matcher_singleton"


def _resolve_project_info_dir() -> Path:
    for candidate in _PROJECT_INFO_CANDIDATES:
        if (candidate / "project_name_matcher.py").exists():
            return candidate
    for candidate in _PROJECT_INFO_CANDIDATES:
        if candidate.exists():
            return candidate
    searched = ", ".join(str(p) for p in _PROJECT_INFO_CANDIDATES)
    raise FileNotFoundError(f"未找到项目基础信息目录，已搜索：{searched}")


def _load_matcher():
    if _SINGLETON_KEY in sys.modules:
        return sys.modules[_SINGLETON_KEY]
    matcher_path = _resolve_project_info_dir() / "project_name_matcher.py"
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
