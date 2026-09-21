"""项目名称匹配公共入口。实现见 common.project_name_matcher。"""
from __future__ import annotations

from common.project_name_matcher import (
    DEFAULT_DB_PATH,
    ProjectMatcher,
    add_mapping,
    build_standard_projects_from_dataframe,
    clear_matcher_cache,
    clear_unmatched_aliases,
    collect_names_needing_prompt,
    default_db_path,
    fill_excel_project_columns,
    find_info_excel,
    get_matcher,
    get_project_lookup,
    get_project_options,
    is_unmatched_alias,
    load_matcher,
    load_or_init_db,
    mark_unmatched,
    match_project,
    peek_excel_sheets_and_headers,
    prompt_unmatched_project_mappings,
    save_db,
    strip_project_prefix_from_summary,
    update_db_from_excel,
)

# 历史别名：部分脚本把 match_project 叫作 match_project_name
match_project_name = match_project

__all__ = [
    "DEFAULT_DB_PATH",
    "ProjectMatcher",
    "add_mapping",
    "build_standard_projects_from_dataframe",
    "clear_matcher_cache",
    "clear_unmatched_aliases",
    "collect_names_needing_prompt",
    "default_db_path",
    "fill_excel_project_columns",
    "find_info_excel",
    "get_matcher",
    "get_project_lookup",
    "get_project_options",
    "is_unmatched_alias",
    "load_matcher",
    "load_or_init_db",
    "mark_unmatched",
    "match_project",
    "match_project_name",
    "peek_excel_sheets_and_headers",
    "prompt_unmatched_project_mappings",
    "save_db",
    "strip_project_prefix_from_summary",
    "update_db_from_excel",
]
