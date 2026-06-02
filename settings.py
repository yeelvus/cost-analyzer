from __future__ import annotations

import os
from pathlib import Path


CODE_ROOT = Path(__file__).resolve().parent
DATA_ROOT = Path(
    os.environ.get(
        "COST_DATA_ROOT",
        "/Users/wyattsun/Desktop/邯郸中材/01成本统计",
    )
)
REPORT_ROOT = Path(os.environ.get("COST_REPORT_ROOT", str(DATA_ROOT / "Z_report")))


def data_path(*parts: str) -> Path:
    return DATA_ROOT.joinpath(*parts)


def report_path(*parts: str) -> Path:
    return REPORT_ROOT.joinpath(*parts)


def code_path(*parts: str) -> Path:
    return CODE_ROOT.joinpath(*parts)


MATCH_DB_PATH = CODE_ROOT / "0_项目基本信息" / "project_name_match_db.json"
