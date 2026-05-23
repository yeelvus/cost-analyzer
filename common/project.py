from __future__ import annotations

from importlib import import_module

_matcher = import_module("0_项目基本信息.project_name_matcher")

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
]
