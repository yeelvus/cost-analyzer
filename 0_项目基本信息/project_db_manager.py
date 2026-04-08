from project_name_matcher import (  # noqa: F401
    add_mapping,
    build_standard_projects_from_dataframe,
    clear_unmatched_aliases,
    get_project_lookup,
    get_project_options,
    is_unmatched_alias,
    load_or_init_db,
    mark_unmatched,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
    update_db_from_excel,
)


# ====================== 主程序 ======================
if __name__ == "__main__":
    # 兼容旧入口：仅保留参数与调用，核心逻辑已统一到 project_name_matcher.py。
    db_path = "db.json"
    excel_path = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"

    # 如果你的 Excel 列名与下面不同，请在这里修改
    update_db_from_excel(
        db_path=db_path,
        excel_path=excel_path,
        status_column="项目状态",      # ← 如果状态列不是“项目状态”，请改成实际列名
        name_column="项目名称",
        code_column="项目编码"
    )

    # 更新完后可以继续使用原来的匹配功能
    print("\n✅ 所有更新已保存到 db.json")
    print("   你现在可以继续调用 match_project()、prompt_unmatched_project_mappings() 等函数进行匹配。")