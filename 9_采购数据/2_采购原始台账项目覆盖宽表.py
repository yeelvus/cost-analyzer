from __future__ import annotations

import importlib.util
import re
from pathlib import Path

import shutil

import pandas as pd
from openpyxl import load_workbook


# 文件夹按关键词动态查找，忽略 A_ / C_ 等前缀字母
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(Path(__file__).parent / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    # 动态加载失败时使用脚本目录兜底，避免直接崩溃
    RAW_ROOT = Path(__file__).parent
    OUTPUT_ROOT = Path(__file__).parent
else:
    _路径_mod = importlib.util.module_from_spec(_路径_spec)
    _路径_spec.loader.exec_module(_路径_mod)
    RAW_ROOT = _路径_mod.RAW_DIR
    OUTPUT_ROOT = _路径_mod.SIMPLIFY_DIR
OUTPUT_FILE = OUTPUT_ROOT / "采购原始台账_项目月份覆盖宽表.xlsx"
PROJECT_INFO_FILE = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx")
PROJECT_MATCH_DB_FILE = Path("/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json")
PROJECT_MATCHER_MODULE_PATH = Path("/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_matcher.py")
EXCEL_FORMATTING_UTILS_PATH = Path("/Users/wyattsun/Desktop/code/0_项目基本信息/excel_formatting_utils.py")

MONTH_PATTERN = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")


def _load_module_from_path(module_name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(file_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_matcher = _load_module_from_path("project_name_matcher", PROJECT_MATCHER_MODULE_PATH)
build_standard_projects_from_dataframe = _matcher.build_standard_projects_from_dataframe
load_or_init_db = _matcher.load_or_init_db
save_db = _matcher.save_db
match_project = _matcher.match_project
mark_unmatched = _matcher.mark_unmatched
add_mapping = _matcher.add_mapping
prompt_unmatched_project_mappings = _matcher.prompt_unmatched_project_mappings

_excel_formatting_utils = _load_module_from_path("excel_formatting_utils", EXCEL_FORMATTING_UTILS_PATH)
apply_common_excel_layout = _excel_formatting_utils.apply_common_excel_layout


def month_key(month_text: str) -> tuple[int, int]:
    y, m = month_text.split("年")
    m = m.replace("月", "")
    return int(y), int(m)


def extract_month_from_folder(folder_name: str) -> str | None:
    m = MONTH_PATTERN.search(folder_name)
    if not m:
        return None
    year = int(m.group(1))
    month = int(m.group(2))
    return f"{year}年{month:02d}月"


def extract_project_name(file_name: str) -> str:
    stem = Path(file_name).stem.strip()
    if "_" in stem:
        return stem.split("_", 1)[0].strip()
    m = re.match(r"^(.+?项目)", stem)
    if m:
        return m.group(1).strip()
    return stem


def list_excel_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            continue
        files.append(p)
    return files


def load_project_config() -> list[dict]:
    if not PROJECT_INFO_FILE.is_file():
        print(f"警告: 未找到项目基础信息文件: {PROJECT_INFO_FILE}")
        return []

    try:
        df = pd.read_excel(PROJECT_INFO_FILE, sheet_name=0)
        return build_standard_projects_from_dataframe(df)
    except Exception as exc:
        print(f"警告: 读取项目基础信息失败: {exc}")
        return []


def resolve_project_by_matcher(raw_name: str, match_db: dict) -> tuple[str, str]:
    hit = match_project(raw_name, match_db)
    if hit:
        std_name = str(hit.get("project_name", raw_name)).strip() or raw_name
        code = str(hit.get("project_code", "")).strip()
        add_mapping(match_db, raw_name, std_name)
        return std_name, code

    mark_unmatched(match_db, raw_name)
    return raw_name, ""


def build_coverage_dataframe(raw_root: Path, match_db: dict) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    records: list[dict] = []
    unmatched_raw_names: set[str] = set()

    month_dirs = [p for p in raw_root.iterdir() if p.is_dir()]
    for month_dir in sorted(month_dirs, key=lambda x: x.name):
        month_text = extract_month_from_folder(month_dir.name)
        if not month_text:
            continue

        for excel_path in list_excel_files(month_dir):
            raw_project_name = extract_project_name(excel_path.name)
            if not raw_project_name:
                continue

            std_project_name, project_code = resolve_project_by_matcher(raw_project_name, match_db)
            if std_project_name == raw_project_name and project_code == "":
                unmatched_raw_names.add(raw_project_name)

            records.append(
                {
                    "月份": month_text,
                    "原始项目名称": raw_project_name,
                    "项目名称": std_project_name,
                    "项目编码": project_code,
                    "文件名": excel_path.name,
                    "文件路径": str(excel_path),
                }
            )

    mapped_count = 0
    if unmatched_raw_names:
        print(f"\n发现 {len(unmatched_raw_names)} 个未匹配项目，打开一次性批量匹配窗口...")
        batch_results = prompt_unmatched_project_mappings(
            sorted(unmatched_raw_names),
            match_db,
            window_title="采购原始台账-批量项目匹配",
            db_path=str(PROJECT_MATCH_DB_FILE),
        )
        mapped_count = sum(1 for _, std in batch_results.items() if std)

        if mapped_count:
            for rec in records:
                raw_name = rec.get("原始项目名称", "")
                hit = match_project(raw_name, match_db)
                if hit:
                    rec["项目名称"] = str(hit.get("project_name", raw_name)).strip() or raw_name
                    rec["项目编码"] = str(hit.get("project_code", "")).strip()

    if not records:
        raise RuntimeError(f"未在目录中扫描到可用台账文件: {raw_root}")

    detail_df = pd.DataFrame(records)

    # 同一月份同一项目可能多个文件，这里统计文件数量
    grouped = (
        detail_df
        .groupby(["项目编码", "项目名称", "月份"], as_index=False)
        .agg(文件数量=("文件名", "count"))
    )

    months = sorted(grouped["月份"].unique().tolist(), key=month_key)

    pivot_count = (
        grouped
        .pivot(index=["项目编码", "项目名称"], columns="月份", values="文件数量")
        .reindex(columns=months)
        .fillna(0)
        .astype(int)
    )

    # 元信息：首次出现月份 / 最新出现月份 / 覆盖月数
    presence = pivot_count.gt(0)
    first_month = []
    last_month = []
    month_covered = []

    for _, row in presence.iterrows():
        hit_months = [m for m in months if bool(row[m])]
        if hit_months:
            first_month.append(hit_months[0])
            last_month.append(hit_months[-1])
            month_covered.append(len(hit_months))
        else:
            first_month.append("")
            last_month.append("")
            month_covered.append(0)

    wide_df = pivot_count.copy()
    wide_df.insert(0, "覆盖月数", month_covered)
    wide_df.insert(0, "最新出现月份", last_month)
    wide_df.insert(0, "首次出现月份", first_month)
    wide_df = wide_df.reset_index()

    # 便于阅读：0 显示为空字符串
    month_cols = months
    for col in month_cols:
        wide_df[col] = wide_df[col].apply(lambda x: "" if int(x) == 0 else int(x))

    wide_df = wide_df.sort_values(["首次出现月份", "项目编码", "项目名称"], ascending=[True, True, True]).reset_index(drop=True)
    return wide_df, detail_df, mapped_count


def format_output_excel(path: Path) -> None:
    wb = load_workbook(path)

    for ws in wb.worksheets:
        if ws.title == "项目月份覆盖宽表":
            apply_common_excel_layout(
                ws,
                numeric_start_col=5,
                header_row=1,
                header_height=28,
                freeze_panes="D2",
                thousand_number_format="#,##0",
                min_col_width=11,
                max_col_width=28,
                extra_padding=2,
            )
        else:
            apply_common_excel_layout(
                ws,
                numeric_start_col=ws.max_column + 1,
                header_row=1,
                header_height=28,
                freeze_panes="A2",
                min_col_width=11,
                max_col_width=40,
                extra_padding=2,
            )

    wb.save(path)


def copy_latest_files(detail_df: pd.DataFrame, dest_dir: Path) -> int:
    """对每个项目取最新月份的文件，复制到 dest_dir，返回复制文件数。"""
    dest_dir.mkdir(parents=True, exist_ok=True)

    # 按项目名称分组，每组取月份最大的那行（若有多文件取第一个）
    detail_df = detail_df.copy()
    detail_df["_month_key"] = detail_df["月份"].apply(month_key)

    latest_rows = (
        detail_df.sort_values("_month_key")
        .groupby("项目名称", as_index=False)
        .last()
    )

    copied = 0
    for _, row in latest_rows.iterrows():
        src = Path(row["文件路径"])
        if not src.is_file():
            print(f"  [跳过] 源文件不存在: {src}")
            continue
        dst = dest_dir / src.name
        shutil.copy2(src, dst)
        copied += 1

    return copied


def main() -> None:
    if not RAW_ROOT.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {RAW_ROOT}")

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    project_config = load_project_config()
    match_db = load_or_init_db(str(PROJECT_MATCH_DB_FILE), project_config)

    wide_df, detail_df, mapped_count = build_coverage_dataframe(RAW_ROOT, match_db)
    save_db(str(PROJECT_MATCH_DB_FILE), match_db)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        wide_df.to_excel(writer, sheet_name="项目月份覆盖宽表", index=False)
        detail_df.to_excel(writer, sheet_name="明细映射", index=False)

    format_output_excel(OUTPUT_FILE)

    # 复制每个项目最新月份文件到 simplify 目录
    copied_count = copy_latest_files(detail_df, OUTPUT_ROOT)

    month_cols = [c for c in wide_df.columns if re.fullmatch(r"(?:19|20)\d{2}年(?:0[1-9]|1[0-2])月", str(c))]
    print("处理完成")
    print(f"扫描目录: {RAW_ROOT}")
    print(f"项目数: {len(wide_df)}")
    print(f"本次新增项目映射: {mapped_count}")
    print(f"月份数: {len(month_cols)}")
    if month_cols:
        print(f"月份范围: {month_cols[0]} ~ {month_cols[-1]}")
    print(f"复制最新文件数: {copied_count}")
    print(f"输出目录: {OUTPUT_ROOT}")


if __name__ == "__main__":
    main()
