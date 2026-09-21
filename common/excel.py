from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
_PROJECT_INFO_CANDIDATES = [
    _CODE_ROOT / "00_项目基本信息",
    _CODE_ROOT / "0_项目基本信息",
]


def _resolve_excel_utils_path() -> Path:
    for candidate in _PROJECT_INFO_CANDIDATES:
        path = candidate / "excel_formatting_utils.py"
        if path.exists():
            return path
    for candidate in _PROJECT_INFO_CANDIDATES:
        path = candidate / "excel_formatting_utils.py"
        if candidate.exists():
            return path
    searched = ", ".join(str(p / 'excel_formatting_utils.py') for p in _PROJECT_INFO_CANDIDATES)
    raise FileNotFoundError(f"未找到 excel_formatting_utils.py，已搜索：{searched}")


_spec = importlib.util.spec_from_file_location("_excel_formatting_utils_runtime", _resolve_excel_utils_path())
if _spec is None or _spec.loader is None:
    raise ImportError(f"无法加载 excel_formatting_utils: {_resolve_excel_utils_path()}")
_excel_utils = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _excel_utils
_spec.loader.exec_module(_excel_utils)

apply_common_excel_layout = _excel_utils.apply_common_excel_layout
set_header_row_style = _excel_utils.set_header_row_style
set_columns_number_format = _excel_utils.set_columns_number_format
set_auto_filter_range = _excel_utils.set_auto_filter_range
style_full_row = _excel_utils.style_full_row
auto_adjust_column_widths = _excel_utils.auto_adjust_column_widths

__all__ = [
    "apply_common_excel_layout",
    "set_header_row_style",
    "set_columns_number_format",
    "set_auto_filter_range",
    "style_full_row",
    "auto_adjust_column_widths",
]
