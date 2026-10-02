# -*- coding: utf-8 -*-
"""外部依赖：项目匹配库、Excel 样式工具。"""

from __future__ import annotations

import importlib.util
import sys
from functools import lru_cache
from pathlib import Path
from types import ModuleType
from typing import Callable, Optional

from .pkg import CODE_ROOT

DESKTOP_CODE = Path("/Users/wyattsun/Desktop/code")


def resolve_project_info_dir() -> Optional[Path]:
    """定位 excel_formatting_utils 所在目录（00_项目基本信息）。"""
    candidates = [
        DESKTOP_CODE / "00_项目基本信息",
        DESKTOP_CODE / "0_项目基本信息",
        CODE_ROOT.parent.parent / "00_项目基本信息",
        CODE_ROOT.parent.parent / "0_项目基本信息",
        CODE_ROOT.parent / "00_项目基本信息",
        CODE_ROOT.parent / "0_项目基本信息",
    ]
    for p in candidates:
        if (p / "excel_formatting_utils.py").is_file():
            return p
    return None


def _load_module_from_path(module_name: str, path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载: {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


@lru_cache(maxsize=1)
def get_excel_layout_fn() -> Callable:
    info = resolve_project_info_dir()
    if not info:
        return lambda *a, **k: None
    path = info / "excel_formatting_utils.py"
    if not path.is_file():
        return lambda *a, **k: None
    try:
        mod = _load_module_from_path("excel_formatting_utils_hub", path)
        return getattr(mod, "apply_common_excel_layout", lambda *a, **k: None)
    except Exception as exc:
        print(f"⚠️ 加载 excel 样式模块失败: {exc}")
        return lambda *a, **k: None


@lru_cache(maxsize=1)
def get_project_matcher_module() -> ModuleType:
    """加载项目名称匹配模块（common.project_name_matcher）。"""
    repo_root = DESKTOP_CODE
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    from common import project_name_matcher

    return project_name_matcher


def project_match_db_path() -> str:
    repo_root = DESKTOP_CODE
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    from settings import MATCH_DB_PATH

    return str(MATCH_DB_PATH)


def get_project_matcher(load_info: bool = True):
    """
    推荐：返回 ProjectMatcher 实例。
        m = get_project_matcher()
        name, code = m.resolve("永强")
    """
    from common.project import ProjectMatcher

    return ProjectMatcher(db_path=project_match_db_path(), load_info=load_info)


def get_subject_mapper(data_dir: Optional[Path] = None):
    """直接用 learn_classify 适配器，不依赖旧 subject_mapper 文件逻辑。"""
    parent = str(CODE_ROOT)
    if parent not in sys.path:
        sys.path.insert(0, parent)
    from learn_classify.adapter import SubjectMapper

    if data_dir is not None:
        return SubjectMapper(data_dir=data_dir)
    repo = CODE_ROOT.parent.parent
    if str(repo) not in sys.path:
        sys.path.insert(0, str(repo))
    from settings import CLASSIFY_DATA_DIR

    return SubjectMapper(data_dir=CLASSIFY_DATA_DIR)
