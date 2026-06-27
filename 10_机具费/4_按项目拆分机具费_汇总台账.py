from __future__ import annotations
import re
import sys
from pathlib import Path

import pandas as pd


INPUT_FILE_NAME = "clean_机具费_raw合并.xlsx"
DEFAULT_INPUT_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_std"
)
FALLBACK_INPUT_DIRS = [
    Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/A_机具费_raw"),
    Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum"),
]
OUTPUT_DIR = DEFAULT_INPUT_DIR
DEFAULT_YEAR = "2025"

OUTPUT_COLUMNS = [
    "序号",
    "机具种类",
    "型号规格",
    "台班费(泰铢)",
    "本次结算台班（小时、天、月)",
    "加班(小时)",
    "加班费+超时费",
    "进出场费/运费",
    "移动费＋拆装费+接管费",
    "结算费用(泰铢)",
    "增值税7%",
    "机具租赁预扣税3%",
    "机具租赁预扣税5%",
    "本次费用合计（泰铢）",
    "使用日期",
    "年月",
    "出租单位",
    "项目名称",
    "项目编码",
    "备注",
]

SOURCE_COLUMNS = [
    "序号",
    "机具种类",
    "型号规格",
    "台班费(泰铢)",
    "本次结算台班（小时、天、月)",
    "加班(小时)",
    "加班费+超时费",
    "进出场费/运费",
    "移动费＋拆装费+接管费",
    "结算费用(泰铢)",
    "增值税7%",
    "机具租赁预扣税3%",
    "机具租赁预扣税5%",
    "本次费用合计（泰铢）",
    "使用日期",
    "出租单位",
    "月份",
    "项目名称",
    "项目编码",
    "备注",
]

COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "序号": ("序号",),
    "机具种类": ("机具种类",),
    "型号规格": ("型号规格",),
    "台班费(泰铢)": ("台班费(泰铢)", "台班费泰铢", "台班费"),
    "本次结算台班（小时、天、月)": (
        "本次结算台班（小时、天、月)",
        "本次结算台班(小时、天、月)",
        "本次结算台班（小时、天、月）",
        "本次结算台班",
    ),
    "加班(小时)": ("加班(小时)", "加班（小时）", "加班小时", "加班"),
    "加班费+超时费": (
        "加班费+超时费",
        "加班费＋超时费",
        "加班费超时费",
        "加班费（小时）",
    ),
    "进出场费/运费": ("进出场费/运费", "进出场费", "运费"),
    "移动费＋拆装费+接管费": (
        "移动费＋拆装费+接管费",
        "移动费+拆装费+接管费",
        "移动费拆装费接管费",
        "运费＋拆装费",
        "运费+拆装费",
    ),
    "结算费用(泰铢)": (
        "结算费用(泰铢)",
        "结算费用（泰铢）",
        "结算费用泰铢",
        "结算费用",
        "本次结算费用(泰铢)",
        "本次结算费用（泰铢）",
        "本次结算费用",
    ),
    "增值税7%": ("增值税7%",),
    "机具租赁预扣税3%": ("机具租赁预扣税3%", "机具租赁预扣税3％"),
    "机具租赁预扣税5%": ("机具租赁预扣税5%", "机具租赁预扣税5％"),
    "本次费用合计（泰铢）": (
        "本次费用合计（泰铢）",
        "本次费用合计(泰铢)",
        "本次费用合计泰铢",
        "本次费用合计",
        "费用合计",
    ),
    "使用日期": ("使用日期", "使用时间"),
    "出租单位": ("出租单位",),
    "月份": ("月份",),
    "备注": ("备注",),
    "项目名称": ("项目名称",),
    "项目编码": ("项目编码",),
}

HEADER_TOKENS = ("序号", "机具种类", "型号规格")


def normalize_col(text: object) -> str:
    value = "" if text is None else str(text)
    value = re.sub(r"\s+", "", value)
    value = value.replace("（", "(").replace("）", ")")
    value = value.replace("＋", "+")
    return value


def safe_filename(name: object) -> str:
    value = "" if name is None else str(name).strip()
    if not value:
        value = "未匹配项目"
    return re.sub(r"[\\/:*?\"<>|]", "_", value)


def detect_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 20) -> int:
    scan_rows = min(max_scan_rows, len(raw_df))
    best_idx = None
    best_hits = -1
    for row_idx in range(scan_rows):
        row_text = "".join(
            str(v).strip() for v in raw_df.iloc[row_idx].tolist() if str(v) != "nan"
        )
        hits = sum(1 for token in HEADER_TOKENS if token in row_text)
        if "月份" in row_text:
            hits += 2
        if hits > best_hits:
            best_hits = hits
            best_idx = row_idx
    if best_idx is None or best_hits < 3:
        raise RuntimeError("未识别到表头行，请检查源文件结构")
    return best_idx


def has_project_value(value: object) -> bool:
    if pd.isna(value):
        return False
    text = str(value).strip()
    if not text:
        return False
    return text.lower() != "nan"


def is_total_like_row(row: pd.Series) -> bool:
    check_cols = ["序号", "机具种类", "型号规格", "备注"]
    parts: list[str] = []
    for col in check_cols:
        if col not in row.index:
            continue
        value = row[col]
        if pd.isna(value):
            continue
        parts.append(str(value).strip())
    row_text = " ".join(parts)
    return any(token in row_text for token in ("合计", "小计", "总计"))


def parse_month_key(value: object) -> str | None:
    if pd.isna(value):
        return None
    if hasattr(value, "year") and hasattr(value, "month"):
        return f"{int(value.year):04d}{int(value.month):02d}"

    text = str(value).strip()
    if not text:
        return None

    match = re.search(r"(20\d{2})[-/](0[1-9]|1[0-2])", text)
    if match:
        return f"{match.group(1)}{match.group(2)}"

    match = re.search(r"(20\d{2})\s*[年./-]\s*(1[0-2]|0?[1-9])", text)
    if match:
        return f"{int(match.group(1)):04d}{int(match.group(2)):02d}"

    match = re.search(r"(?<!\d)(1[0-2]|0?[1-9])\s*月", text)
    if match:
        return f"{DEFAULT_YEAR}{int(match.group(1)):02d}"

    return None


def format_month_label(month_key: object) -> str:
    text = "" if month_key is None else str(month_key).strip()
    match = re.fullmatch(r"(20\d{2})(0[1-9]|1[0-2])", text)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    return f"{DEFAULT_YEAR}-01"


def resolve_input_file() -> Path:
    candidate = DEFAULT_INPUT_DIR / INPUT_FILE_NAME
    if candidate.exists():
        return candidate

    for fallback_dir in FALLBACK_INPUT_DIRS:
        fallback = fallback_dir / INPUT_FILE_NAME
        if fallback.exists():
            print(f"警告: 在默认输入目录未找到文件，改用备选路径: {fallback}")
            return fallback

    raise FileNotFoundError(
        f"输入文件不存在: {INPUT_FILE_NAME}\n"
        f"已检查路径: {DEFAULT_INPUT_DIR} 和 {', '.join(str(p) for p in FALLBACK_INPUT_DIRS)}"
    )


def extract_project_name_from_sheet(raw_df: pd.DataFrame, sheet_name: str) -> str:
    scan_rows = min(10, len(raw_df))
    for row_idx in range(scan_rows):
        for value in raw_df.iloc[row_idx].tolist():
            if pd.isna(value):
                continue
            text = str(value).strip()
            match = re.search(r"项目名称[:：]\s*(.+)", text)
            if match:
                project_name = match.group(1).strip()
                if project_name:
                    return project_name
    return sheet_name.strip()


def load_sheet_data(input_file: Path, sheet_name: str) -> tuple[pd.DataFrame, str]:
    raw_df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
    header_row = detect_header_row(raw_df)
    project_raw_name = extract_project_name_from_sheet(raw_df, sheet_name)

    data_df = pd.read_excel(input_file, sheet_name=sheet_name, header=header_row)
    normalized_to_source = {normalize_col(col): col for col in data_df.columns}

    selected = pd.DataFrame()
    for target_col in SOURCE_COLUMNS:
        source_col = None
        for alias in COLUMN_ALIASES.get(target_col, (target_col,)):
            alias_key = normalize_col(alias)
            if alias_key in normalized_to_source:
                source_col = normalized_to_source[alias_key]
                break
        selected[target_col] = data_df[source_col] if source_col is not None else None

    selected = selected.dropna(
        how="all",
        subset=["机具种类", "型号规格", "本次费用合计（泰铢）", "使用日期", "月份"],
    )
    selected = selected[~selected.apply(is_total_like_row, axis=1)].copy()
    selected["原始项目名称"] = project_raw_name
    return selected, project_raw_name


def main() -> None:
    input_file = resolve_input_file()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    excel_file = pd.ExcelFile(input_file)
    all_rows: list[pd.DataFrame] = []

    for sheet_name in excel_file.sheet_names:
        selected, project_raw_name = load_sheet_data(input_file, sheet_name)
        if selected.empty:
            continue
        all_rows.append(selected)

    if not all_rows:
        raise ValueError("未读取到可拆分的数据行")

    merged_df = pd.concat(all_rows, ignore_index=True)

    if "项目名称" not in merged_df.columns:
        merged_df["项目名称"] = ""
    if "项目编码" not in merged_df.columns:
        merged_df["项目编码"] = "需手动核对"

    merged_df["月份键"] = merged_df["月份"].apply(parse_month_key)
    merged_df = merged_df[merged_df["月份键"].apply(has_project_value)].copy()
    merged_df["年月"] = merged_df["月份键"].apply(format_month_label)

    output_count = 0
    for (month_key, project_name, project_code), group_df in merged_df.groupby(
        ["月份键", "项目名称", "项目编码"],
        dropna=False,
    ):
        month_text = str(month_key).strip()
        if not month_text:
            continue
        project_text = safe_filename(project_name)
        output_path = OUTPUT_DIR / f"{month_text}_{project_text}机具费.xlsx"
        result_df = group_df[OUTPUT_COLUMNS].copy()
        result_df.to_excel(output_path, index=False)
        output_count += 1
        print(f"输出: {output_path} | 行数={len(result_df)}")

    print(f"完成。共输出 {output_count} 个文件。")


if __name__ == "__main__":
    main()
