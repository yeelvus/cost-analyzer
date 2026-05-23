from __future__ import annotations

from pathlib import Path

from settings import CODE_ROOT, DATA_ROOT, REPORT_ROOT, code_path, data_path, report_path


def find_child_dir(base: Path, keyword: str) -> Path:
    if base.is_dir():
        for child in sorted(base.iterdir()):
            if child.is_dir() and keyword in child.name:
                return child
    return base / keyword


def find_script(script_dir: Path, keyword: str) -> Path:
    if script_dir.is_dir():
        for script in sorted(script_dir.iterdir()):
            if script.is_file() and script.suffix == ".py" and keyword in script.name:
                return script
    return script_dir / keyword


__all__ = [
    "CODE_ROOT",
    "DATA_ROOT",
    "REPORT_ROOT",
    "code_path",
    "data_path",
    "report_path",
    "find_child_dir",
    "find_script",
]
