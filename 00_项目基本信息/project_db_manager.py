# -*- coding: utf-8 -*-
"""
兼容入口：从项目基本信息表刷新匹配库。

推荐直接：
  python project_name_matcher.py refresh
  或  ProjectMatcher().refresh_from_excel()
"""

from __future__ import annotations

from pathlib import Path

from project_name_matcher import (  # noqa: F401 — 再导出，兼容旧 import
    DEFAULT_DB_PATH,
    ProjectMatcher,
    add_mapping,
    build_standard_projects_from_dataframe,
    clear_unmatched_aliases,
    find_info_excel,
    get_project_lookup,
    get_project_options,
    is_unmatched_alias,
    load_matcher,
    load_or_init_db,
    mark_unmatched,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
    update_db_from_excel,
)

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
