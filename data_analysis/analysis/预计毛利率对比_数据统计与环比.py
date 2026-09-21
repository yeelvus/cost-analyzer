
# --- path bootstrap (S_数据分析模块) ---
import sys
from pathlib import Path as _PathForBootstrap
try:
    from S_数据分析模块._bootstrap import ensure_code_root, ensure_script_dir
    ensure_code_root(_PathForBootstrap(__file__))
    ensure_script_dir(__file__)
except Exception:
    _p = _PathForBootstrap(__file__).resolve()
    for _c in [_p.parent, *_p.parents]:
        if (_c / "settings.py").exists():
            if str(_c) not in sys.path:
                sys.path.insert(0, str(_c))
            break
    if str(_p.parent) not in sys.path:
        sys.path.insert(0, str(_p.parent))
# --- end bootstrap ---

import importlib.util
import re
from pathlib import Path

import pandas as pd


CONFIG = {
    "source_dir": r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/1_成本报公司",
    "output_file": r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/1_成本报公司/对比文件/项目预计毛利率对比_数据统计与环比.xlsx",
    "matcher_file": r"/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_matcher.py",
    "match_db_file": r"/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json",
    "project_info_file": r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx",
    "excel_formatting_utils_file": r"/Users/wyattsun/Desktop/code/0_项目基本信息/excel_formatting_utils.py",
}

FILE_PATTERN = re.compile(r"^(\d{4})年(\d{1,2})月泰国子公司各项目成本统计表\.xlsx$")
PROJECT_HEADER_CANDIDATES = {"项目名称", "项目名"}
MARGIN_HEADER_CANDIDATES = {"预计毛利率", "毛利率", "预计利润率"}


def load_module(module_path, module_name):
    module_spec = importlib.util.spec_from_file_location(module_name, str(module_path))
    if module_spec is None or module_spec.loader is None:
        raise ImportError(f"无法加载模块: {module_path}")
    module = importlib.util.module_from_spec(module_spec)
    module_spec.loader.exec_module(module)
    return module


def normalize_header_text(value):
    if value is None:
        return ""
    return str(value).strip().replace("\n", "").replace(" ", "")


def parse_margin_value(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        is_percent = text.endswith("%")
        text = text.rstrip("%")
        try:
            numeric = float(text)
        except ValueError:
            return None
        if is_percent:
            return numeric / 100.0
        return numeric / 100.0 if 1 < abs(numeric) <= 100 else numeric
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    return numeric / 100.0 if 1 < abs(numeric) <= 100 else numeric


def find_header_row_and_cols(raw_df):
    max_scan_rows = min(30, len(raw_df))
    for row_idx in range(max_scan_rows):
        row_values = [normalize_header_text(v) for v in raw_df.iloc[row_idx].tolist()]
        project_col_idx = None
        margin_col_idx = None
        for col_idx, header_text in enumerate(row_values):
            if project_col_idx is None and header_text in PROJECT_HEADER_CANDIDATES:
                project_col_idx = col_idx
            if margin_col_idx is None and header_text in MARGIN_HEADER_CANDIDATES:
                margin_col_idx = col_idx
        if project_col_idx is not None and margin_col_idx is not None:
            return row_idx, project_col_idx, margin_col_idx
    return None, None, None


def parse_month_from_filename(file_name):
    match = FILE_PATTERN.match(file_name)
    if not match:
        return None, None
    year = int(match.group(1))
    month = int(match.group(2))
    month_text = f"{year}年{month}月"
    month_ts = pd.Timestamp(year=year, month=month, day=1)
    return month_text, month_ts


def collect_source_files(source_dir):
    source_path = Path(source_dir)
    files = [
        p for p in sorted(source_path.glob("*.xlsx"), key=lambda x: x.name)
        if FILE_PATTERN.match(p.name) and not p.name.startswith("~$")
    ]
    return files


def extract_project_margin_rows(file_path, month_text, month_ts):
    all_rows = []
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=object)
        header_row, project_col_idx, margin_col_idx = find_header_row_and_cols(raw_df)
        if header_row is None:
            continue

        data_df = raw_df.iloc[header_row + 1 :, [project_col_idx, margin_col_idx]].copy()
        data_df.columns = ["项目名称", "预计毛利率"]
        data_df["项目名称"] = data_df["项目名称"].apply(lambda v: "" if pd.isna(v) else str(v).strip())
        data_df = data_df[data_df["项目名称"] != ""]
        data_df["预计毛利率"] = data_df["预计毛利率"].apply(parse_margin_value)
        data_df = data_df.dropna(subset=["预计毛利率"])
        if data_df.empty:
            continue

        data_df.insert(0, "月份", month_text)
        data_df.insert(1, "月份日期", month_ts)
        data_df.insert(2, "来源文件", file_path.name)
        data_df.insert(3, "来源Sheet", sheet_name)
        all_rows.append(data_df)
        break

    if not all_rows:
        return pd.DataFrame(columns=["月份", "月份日期", "来源文件", "来源Sheet", "项目名称", "预计毛利率"])
    return pd.concat(all_rows, ignore_index=True)


def init_matcher_and_db(config):
    matcher_module = load_module(Path(config["matcher_file"]), "project_name_matcher_for_margin_compare")

    standard_projects = []
    project_info_file = Path(config["project_info_file"])
    if project_info_file.exists():
        try:
            project_info_df = pd.read_excel(project_info_file)
            standard_projects = matcher_module.build_standard_projects_from_dataframe(project_info_df)
        except Exception as exc:
            print(f"⚠️ 加载项目基本信息失败，将仅使用已有匹配库: {exc}")

    match_db_file = Path(config["match_db_file"])
    match_db = matcher_module.load_or_init_db(str(match_db_file), standard_projects)
    return matcher_module, match_db, match_db_file


def apply_project_matching(df, matcher_module, match_db):
    standardized_names = []
    project_codes = []
    match_sources = []

    for raw_name in df["项目名称"].tolist():
        hit = matcher_module.match_project(raw_name, match_db)
        if hit:
            standardized_names.append(str(hit.get("project_name", raw_name)).strip() or str(raw_name).strip())
            project_codes.append(str(hit.get("project_code", "需手动核对")).strip() or "需手动核对")
            match_sources.append(str(hit.get("source", "alias")))
        else:
            matcher_module.mark_unmatched(match_db, raw_name)
            standardized_names.append(str(raw_name).strip())
            project_codes.append("未匹配")
            match_sources.append("unmatched")

    out_df = df.copy()
    out_df["标准项目名称"] = standardized_names
    out_df["项目编码"] = project_codes
    out_df["匹配来源"] = match_sources
    return out_df


def build_ring_ratio_sheet(stats_df):
    if stats_df.empty:
        return pd.DataFrame()

    ordered_months = (
        stats_df[["月份", "月份日期"]]
        .drop_duplicates()
        .sort_values("月份日期")["月份"]
        .tolist()
    )

    base = stats_df[["标准项目名称", "月份", "月份日期", "预计毛利率"]].copy()
    base = base.sort_values(["标准项目名称", "月份日期"])
    base["毛利率环比变化"] = base.groupby("标准项目名称")["预计毛利率"].diff()

    wide_df = base.pivot_table(
        index="标准项目名称",
        columns="月份",
        values="毛利率环比变化",
        aggfunc="last",
    )
    wide_df = wide_df.reindex(columns=ordered_months)
    wide_df.index.name = "项目名称"
    return wide_df.reset_index()


def write_result_excel(stats_df, ring_ratio_df, output_file, formatting_module):
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    stats_export_df = stats_df[["月份", "标准项目名称", "项目编码", "项目名称", "预计毛利率", "匹配来源", "来源文件", "来源Sheet"]].copy()
    stats_export_df = stats_export_df.rename(columns={"标准项目名称": "项目名称(标准)", "项目名称": "项目名称(原始)"})

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        stats_export_df.to_excel(writer, index=False, sheet_name="数据统计")
        ring_ratio_df.to_excel(writer, index=False, sheet_name="环比变化")

        ws_stats = writer.book["数据统计"]
        formatting_module.apply_common_excel_layout(
            ws_stats,
            numeric_start_col=5,
            percent_headers=["预计毛利率"],
            percent_number_format="0.00%",
            freeze_panes="A2",
        )

        ws_ring = writer.book["环比变化"]
        formatting_module.apply_common_excel_layout(
            ws_ring,
            numeric_start_col=2,
            freeze_panes="B2",
        )

        if ws_ring.max_column >= 2:
            formatting_module.set_columns_number_format(
                ws_ring,
                col_indices=range(2, ws_ring.max_column + 1),
                number_format="0.00%",
                start_row=2,
            )


def main():
    source_files = collect_source_files(CONFIG["source_dir"])
    if not source_files:
        print(f"⚠️ 未找到符合命名规则的源文件: {CONFIG['source_dir']}")
        return

    print(f"✓ 找到源文件 {len(source_files)} 个")
    all_parts = []
    for file_path in source_files:
        month_text, month_ts = parse_month_from_filename(file_path.name)
        if month_text is None:
            continue
        part_df = extract_project_margin_rows(file_path, month_text, month_ts)
        if part_df.empty:
            print(f"⚠️ 未在文件中识别到 项目名称/预计毛利率 列: {file_path.name}")
            continue
        all_parts.append(part_df)
        print(f"✓ 已提取: {file_path.name} -> {len(part_df)} 条")

    if not all_parts:
        print("⚠️ 没有可用数据，未生成输出文件。")
        return

    raw_df = pd.concat(all_parts, ignore_index=True)
    matcher_module, match_db, db_path = init_matcher_and_db(CONFIG)
    stats_df = apply_project_matching(raw_df, matcher_module, match_db)
    stats_df = stats_df.sort_values(["月份日期", "标准项目名称"]).reset_index(drop=True)

    ring_ratio_df = build_ring_ratio_sheet(stats_df)

    formatting_module = load_module(Path(CONFIG["excel_formatting_utils_file"]), "excel_formatting_utils_for_margin_compare")
    write_result_excel(stats_df, ring_ratio_df, CONFIG["output_file"], formatting_module)

    matcher_module.save_db(str(db_path), match_db)
    print(f"✓ 输出完成: {CONFIG['output_file']}")


if __name__ == "__main__":
    main()
