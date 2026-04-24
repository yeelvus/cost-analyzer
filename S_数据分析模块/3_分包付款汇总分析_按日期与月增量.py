import os
import re
from datetime import datetime
from pathlib import Path
import importlib.util
import pandas as pd

_SCRIPT_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _SCRIPT_DIR.parent.parent  # Desktop 下一层


def _resolve_input_dir() -> Path:
    candidates = [
        Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_分包对账与付款/B_分包对账与付款_sum"),
        _PROJECT_ROOT / "邯郸中材" / "01成本统计" / "12-财务数据💰" / "A_分包对账与付款" / "B_分包对账与付款_sum",
        _PROJECT_ROOT / "01成本统计" / "12-财务数据💰" / "A_分包对账与付款" / "B_分包对账与付款_sum",
    ]
    for p in candidates:
        if p.is_dir():
            return p
    return candidates[0]


def _resolve_format_util_path() -> Path | None:
    candidates = [
        _SCRIPT_DIR.parent / "0_项目基本信息" / "excel_formatting_utils.py",
        Path("/Users/wyattsun/Desktop/code/0_项目基本信息/excel_formatting_utils.py"),
        _PROJECT_ROOT / "code" / "0_项目基本信息" / "excel_formatting_utils.py",
        Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/19_代码/0_项目基本信息/excel_formatting_utils.py"),
    ]
    for p in candidates:
        if p.is_file():
            return p
    return None


INPUT_DIR = _resolve_input_dir()
OUTPUT_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report/财务分包付款_分析文件")
FORMAT_UTIL_PATH = _resolve_format_util_path()


def load_apply_common_excel_layout():
    """动态加载格式化工具，找不到时返回空降级函数。"""
    if FORMAT_UTIL_PATH is None or not FORMAT_UTIL_PATH.is_file():
        print(f"警告: 未找到 excel_formatting_utils.py，已跳过输出样式设置。")
        return lambda *a, **kw: None
    spec = importlib.util.spec_from_file_location("excel_formatting_utils", str(FORMAT_UTIL_PATH))
    if spec is None or spec.loader is None:
        print(f"警告: 无法加载格式化工具: {FORMAT_UTIL_PATH}")
        return lambda *a, **kw: None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not hasattr(module, "apply_common_excel_layout"):
        print("警告: 格式化工具中未找到 apply_common_excel_layout")
        return lambda *a, **kw: None
    return module.apply_common_excel_layout


def extract_date_from_filename(file_name: str) -> str | None:
    """从文件名提取 yyyy-mm-dd 日期。"""
    m = re.search(r"(\d{4}-\d{2}-\d{2})", file_name)
    if not m:
        return None
    return m.group(1)


def pick_latest_files_by_prefix(files: list[Path]) -> list[Path]:
    """同前缀文件仅保留后缀时间最新的一份。"""
    # key: prefix, value: (timestamp_text, file_path)
    latest_map: dict[str, tuple[str, Path]] = {}
    for file_path in files:
        stem = file_path.stem
        m = re.match(r"^(.*)_(\d{8}_\d{6})$", stem)
        if m:
            prefix = m.group(1)
            ts = m.group(2)
        else:
            # 不符合后缀规则的文件，按自身文件名作为唯一前缀保留
            prefix = stem
            ts = "00000000_000000"
        old = latest_map.get(prefix)
        if old is None or ts > old[0]:
            latest_map[prefix] = (ts, file_path)
    selected = [v[1] for v in latest_map.values()]
    return sorted(selected, key=lambda p: p.name)


def to_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def build_project_fallback_from_subitems(file_path: Path, key_cols: list[str]) -> pd.DataFrame:
    """当项目汇总sheet金额为空时，用子项汇总sheet按项目聚合兜底。"""
    fallback_sheet = "项目名称子项汇总（万泰铢）"
    amount_col = "付款汇总（万泰铢）"
    try:
        sub_df = pd.read_excel(file_path, sheet_name=fallback_sheet, dtype=object)
    except Exception as exc:
        print(f" ⚠️ 项目汇总兜底读取失败 {file_path.name}: {exc}")
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    if sub_df.empty:
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    sub_df.columns = [str(c).strip() for c in sub_df.columns]
    if key_cols[0] not in sub_df.columns or amount_col not in sub_df.columns:
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    use_df = sub_df[[key_cols[0], amount_col]].copy()
    use_df[key_cols[0]] = use_df[key_cols[0]].fillna("").astype(str).str.strip()
    use_df = use_df[use_df[key_cols[0]] != ""]
    use_df["汇总金额"] = to_numeric_series(use_df[amount_col])
    use_df = use_df.groupby(key_cols, dropna=False)["汇总金额"].sum().reset_index()
    print(f" ℹ️ 已使用子项汇总兜底重建项目汇总: {file_path.name}")
    return use_df


def read_category_rows(file_path: Path, sheet_name: str, key_cols: list[str]) -> pd.DataFrame:
    """读取一个类别sheet并标准化为 key_cols + [汇总金额]。"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=object)
    except Exception as exc:
        print(f" ⚠️ 读取失败 {file_path.name} / {sheet_name}: {exc}")
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    if df.empty:
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    df.columns = [str(c).strip() for c in df.columns]
    amount_col = "付款汇总（万泰铢）"
    missing_keys = [c for c in key_cols if c not in df.columns]

    if missing_keys or amount_col not in df.columns:
        print(
            f" ⚠️ 缺少字段 {file_path.name} / {sheet_name}: "
            f"keys缺失={missing_keys}, amount存在={amount_col in df.columns}"
        )
        return pd.DataFrame(columns=key_cols + ["汇总金额"])

    use_df = df[key_cols + [amount_col]].copy()
    for col in key_cols:
        use_df[col] = use_df[col].fillna("").astype(str).str.strip()

    use_df = use_df[use_df[key_cols[0]] != ""]

    raw_amount = pd.to_numeric(use_df[amount_col], errors="coerce")

    # 项目汇总金额为空时尝试用子项兜底
    if (
        sheet_name == "项目汇总（万泰铢）"
        and raw_amount.notna().sum() == 0
        and key_cols == ["项目名称"]
    ):
        return build_project_fallback_from_subitems(file_path, key_cols)

    use_df["汇总金额"] = raw_amount.fillna(0.0)
    use_df = use_df.drop(columns=[amount_col])

    # 同一文件中可能存在重复key，先汇总一次
    use_df = use_df.groupby(key_cols, dropna=False)["汇总金额"].sum().reset_index()
    return use_df


def build_wide_tables(long_df: pd.DataFrame, key_cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """由长表生成“按日期汇总”与“按月增加量”宽表。"""
    if long_df.empty:
        return pd.DataFrame(columns=key_cols), pd.DataFrame(columns=key_cols)

    long_df = long_df.copy()
    long_df["日期"] = pd.to_datetime(long_df["日期"], errors="coerce")
    long_df = long_df[long_df["日期"].notna()]
    if long_df.empty:
        return pd.DataFrame(columns=key_cols), pd.DataFrame(columns=key_cols)

    long_df["日期"] = long_df["日期"].dt.strftime("%Y-%m-%d")

    wide_sum = long_df.pivot_table(
        index=key_cols,
        columns="日期",
        values="汇总金额",
        aggfunc="sum",
        fill_value=0.0,
    )

    ordered_date_cols = sorted(wide_sum.columns, key=lambda x: datetime.strptime(x, "%Y-%m-%d"))
    wide_sum = wide_sum[ordered_date_cols]

    wide_inc = wide_sum.diff(axis=1)
    if not wide_inc.empty:
        first_col = wide_inc.columns[0]
        # 首月无上月可比，不统计增量，保留为空
        wide_inc[first_col] = pd.NA

    wide_sum = wide_sum.reset_index()
    wide_inc = wide_inc.reset_index()

    return wide_sum, wide_inc


def aggregate_category(files: list[Path], sheet_name: str, key_cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    all_rows = []
    for file_path in files:
        date_text = extract_date_from_filename(file_path.name)
        if not date_text:
            print(f" ⚠️ 跳过（文件名无日期）: {file_path.name}")
            continue

        one_df = read_category_rows(file_path, sheet_name, key_cols)
        if one_df.empty:
            continue

        one_df["日期"] = date_text
        all_rows.append(one_df)

    if not all_rows:
        return pd.DataFrame(columns=key_cols), pd.DataFrame(columns=key_cols)

    long_df = pd.concat(all_rows, ignore_index=True)
    return build_wide_tables(long_df, key_cols)


def main():
    print("=" * 70)
    print("开始生成：分包付款多期汇总与按月增加量分析")
    print("=" * 70)
    print(f"输入目录: {INPUT_DIR}")
    print(f"输出目录: {OUTPUT_DIR}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    all_files = sorted(
        [p for p in INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.name,
    )

    files = pick_latest_files_by_prefix(all_files)

    if not files:
        print("❌ 未找到可处理的xlsx文件。")
        return

    print(f"原始文件数量: {len(all_files)}")
    print(f"按前缀去重后数量: {len(files)}")
    for p in files:
        print(f" - {p.name}")

    project_key_cols = ["项目名称"]
    subitem_key_cols = ["项目名称", "子项"]

    print("\n1) 处理 项目汇总（万泰铢）...")
    project_sum_df, project_inc_df = aggregate_category(
        files=files,
        sheet_name="项目汇总（万泰铢）",
        key_cols=project_key_cols,
    )

    print("2) 处理 项目名称子项汇总（万泰铢）...")
    subitem_sum_df, subitem_inc_df = aggregate_category(
        files=files,
        sheet_name="项目名称子项汇总（万泰铢）",
        key_cols=subitem_key_cols,
    )

    # 生成不带时间后缀的文件名
    out_path = OUTPUT_DIR / "分包付款多期对比_汇总与月增量.xlsx"

    apply_common_excel_layout = load_apply_common_excel_layout()

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        project_sum_df.to_excel(writer, sheet_name="项目汇总_汇总", index=False)
        project_inc_df.to_excel(writer, sheet_name="项目汇总_按月增加量", index=False)
        subitem_sum_df.to_excel(writer, sheet_name="子项汇总_汇总", index=False)
        subitem_inc_df.to_excel(writer, sheet_name="子项汇总_按月增加量", index=False)

        wb = writer.book
        layout_specs = [
            ("项目汇总_汇总", len(project_key_cols) + 1),
            ("项目汇总_按月增加量", len(project_key_cols) + 1),
            ("子项汇总_汇总", len(subitem_key_cols) + 1),
            ("子项汇总_按月增加量", len(subitem_key_cols) + 1),
        ]

        for sheet_name, numeric_start_col in layout_specs:
            ws = wb[sheet_name]
            apply_common_excel_layout(
                ws,
                numeric_start_col=numeric_start_col,
                header_row=1,
                freeze_panes="A2",
                thousand_number_format="#,##0.00",
            )

    print("\n" + "=" * 70)
    print("✅ 分析文件生成完成")
    print(f"输出文件: {out_path}")
    print("说明: 共4个sheet，分别对应项目/子项的汇总与按月增加量。")
    print("=" * 70)


if __name__ == "__main__":
    main()