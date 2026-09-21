# -*- coding: utf-8 -*-
"""
从项目基本信息表刷新匹配库。

推荐直接：
  python -m common.project_name_matcher refresh
"""

from __future__ import annotations

import sys
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.project import DEFAULT_DB_PATH, ProjectMatcher, find_info_excel

if __name__ == "__main__":
    db_path = DEFAULT_DB_PATH
    excel_path = find_info_excel()
    if not excel_path:
        excel_path = Path(
            "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
        )

    print(f"DB:    {db_path}")
    print(f"Excel: {excel_path}")
    m = ProjectMatcher(db_path=db_path, load_info=False)
    m.refresh_from_excel(excel_path)
    print(f"\n✅ 完成。当前标准项目: {m.project_count}")
