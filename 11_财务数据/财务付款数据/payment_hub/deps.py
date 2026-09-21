# -*- coding: utf-8 -*-
"""外部依赖探测：项目匹配库、Excel 样式工具。"""

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
    candidates = [
        CODE_ROOT.parent.parent / "0_项目基本信息",  # code/0_项目基本信息
        CODE_ROOT.parent / "0_项目基本信息",
        DESKTOP_CODE / "0_项目基本信息",
    ]
    for p in candidates:
        if p.is_dir():
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
    """加载 project_name_matcher 模块（兼容旧函数式 API）。"""
    info = resolve_project_info_dir()
    if not info:
        raise FileNotFoundError("未找到 0_项目基本信息 目录（需要 project_name_matcher.py）")
    path = info / "project_name_matcher.py"
    if not path.is_file():
        raise FileNotFoundError(f"未找到: {path}")
    if str(info) not in sys.path:
        sys.path.insert(0, str(info))
    return _load_module_from_path("project_name_matcher_hub", path)


def project_match_db_path() -> str:
    info = resolve_project_info_dir()
    if not info:
        return str(CODE_ROOT / "project_name_match_db.json")
    return str(info / "project_name_match_db.json")


def get_project_matcher(load_info: bool = True):
    """
    推荐：返回 ProjectMatcher 实例。
        m = get_project_matcher()
        name, code = m.resolve("永强")
    """
    mod = get_project_matcher_module()
    cls = getattr(mod, "ProjectMatcher", None)
    if cls is None:
        raise AttributeError("project_name_matcher 缺少 ProjectMatcher，请更新 0_项目基本信息")
    return cls(db_path=project_match_db_path(), load_info=load_info)


def get_subject_mapper(data_dir: Optional[Path] = None):
    """直接用 learn_classify 适配器，不依赖旧 subject_mapper 文件逻辑。"""
    # 确保 财务付款数据 在 path 上
    parent = str(CODE_ROOT)
    if parent not in sys.path:
        sys.path.insert(0, parent)
    from learn_classify.adapter import SubjectMapper

    if data_dir is not None:
        return SubjectMapper(data_dir=data_dir)
    return SubjectMapper(data_dir=CODE_ROOT / "learn_classify" / "data")
