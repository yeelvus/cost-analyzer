import os
import re
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd
from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent
_CODE_ROOT = SCRIPT_DIR.parent

if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.project import (
    build_standard_projects_from_dataframe,
    load_or_init_db,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
)
from common.excel import apply_common_excel_layout
from settings import MATCH_DB_PATH as _SETTINGS_MATCH_DB_PATH


# ====================== 配置区 ======================
class Config:
    INPUT_SOURCE = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_2_飞签机票/A_飞签台账_raw"
    INPUT_GLOB = "*.xlsx"

    MAPPING_EXCEL_PATH = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
    MATCH_DB_PATH = str(_SETTINGS_MATCH_DB_PATH)

    OUTPUT_STD = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/11_人工费🙎🏻‍♂️/A_2_飞签机票/B_飞签机票_std"

    MONEY_COL = "本次机票价格（泰铢）"

    # 目标保留列（顺序即输出顺序）
    TARGET_COLUMNS = [
        "所属单位",
        "姓名",
        "人次",
        "订票时间",
        "年-月",
        "去程机票日期",
        "回程机票日期",
        "本次机票航班号",
        "舱位",
        "本次机票价格（泰铢）",
        "人均机票价格（泰铢）",
        "机票类型",
        "票代",
        "护照",
        "项目编码",
        "项目名称",
        "原始项目名称",
        "_来源文件",
    ]


# ====================== 辅助函数 ======================
def _sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|\r\n\t]', "_", str(name)).strip()


def _parse_money(series: pd.Series) -> pd.Series:
    def to_num(val):
        if pd.isna(val):
            return 0.0
        text = str(val).strip().replace(",", "").replace("，", "").replace("THB", "").replace("thb", "")
        is_neg = text.startswith("(") and text.endswith(")")
        text = text.strip("()")
        text = re.sub(r"[^0-9.\-]", "", text)
        try:
            num = float(text)
        except ValueError:
            num = 0.0
        return -num if is_neg else num

    return series.apply(to_num)


def _parse_person_count(series: pd.Series) -> pd.Series:
    """根据姓名列解析人次：按常见分隔符拆分后计数。"""
    _sep = re.compile(r"[，,、/\\\s\u3000]+")

    def count_names(val):
        if pd.isna(val):
            return 1
        parts = [p.strip() for p in _sep.split(str(val)) if p.strip()]
        return max(len(parts), 1)

    return series.apply(count_names)


def _parse_year_month(series: pd.Series) -> pd.Series:
    """从日期列解析出 'YYYY-MM' 字符串。"""
    converted = pd.to_datetime(series, errors="coerce", dayfirst=False)
    return converted.dt.strftime("%Y-%m").where(converted.notna(), other="")


def _pick_col(columns, candidates: List[str]) -> Optional[str]:
    """返回第一个在 columns 中能模糊匹配到的候选列名。"""
    for cand in candidates:
        for col in columns:
            if cand in str(col):
                return col
    return None


INVALID_NAMES = {"", "nan", "none", "null", "n/a", "na", "-", "--", "/"}


# ====================== 读取所有原始 Excel ======================
def load_all_raw(input_dir: Path) -> pd.DataFrame:
    files = sorted(input_dir.glob(Config.INPUT_GLOB))
    if not files:
        raise FileNotFoundError(f"未在 {input_dir} 找到任何 .xlsx 文件")

    frames: List[pd.DataFrame] = []
    for f in files:
        print(f"  📄 读取：{f.name}")
        try:
            df_s = pd.read_excel(f, sheet_name=0)
            df_s.dropna(how="all", inplace=True)
            df_s.dropna(axis=1, how="all", inplace=True)
            df_s.columns = [str(c).strip() for c in df_s.columns]
            df_s["_来源文件"] = f.name
            frames.append(df_s)
        except Exception as e:
            print(f"  ⚠️ 跳过（读取失败）：{f.name} — {e}")

    if not frames:
        raise ValueError("所有文件均无有效数据")

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined.columns = [str(c).strip() for c in combined.columns]
    return combined


# ====================== 列标准化 ======================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {
        "所属单位":       ["所属单位", "单位", "部门"],
        "姓名":           ["姓名", "员工姓名", "人员"],
        "订票时间":       ["订票时间", "订票日期", "购票时间", "购票日期"],
        "去程机票日期":   ["去程机票日期", "去程日期", "出发日期"],
        "回程机票日期":   ["回程机票日期", "回程日期", "返回日期"],
        "本次机票航班号": ["本次机票航班号", "航班号", "航班"],
        "舱位":           ["舱位", "座位等级"],
        "本次机票价格（泰铢）": ["本次机票价格", "机票价格", "票价", "金额"],
        "机票类型":       ["机票类型", "类型"],
        "票代":           ["票代", "票务代理"],
        "护照":           ["护照", "护照号", "证件"],
    }

    rename_dict: dict = {}
    for standard, candidates in col_map.items():
        found = _pick_col(df.columns, candidates)
        if found and found not in rename_dict:
            rename_dict[found] = standard

    # "项目"列使用精确匹配，防止误命中"项目名称"等子串列
    if "项目" not in rename_dict.values():
        for exact_name in ["项目", "所属项目"]:
            if exact_name in df.columns and exact_name not in rename_dict:
                rename_dict[exact_name] = "项目"
                break

    df = df.rename(columns=rename_dict)

    # 补全缺失列
    for col in list(col_map.keys()):
        if col not in df.columns:
            df[col] = ""

    return df


# ====================== 项目名称标准化 ======================
def standardize_projects(df: pd.DataFrame) -> pd.DataFrame:
    df["原始项目名称"] = df["项目"].astype(str).str.strip()

    if not os.path.exists(Config.MAPPING_EXCEL_PATH):
        print("⚠️ 映射表不存在，跳过项目标准化")
        df["项目编码"] = ""
        df["项目名称"] = df["原始项目名称"]
        return df

    mapping_df = pd.read_excel(Config.MAPPING_EXCEL_PATH)
    mapping_df.columns = [str(c).strip() for c in mapping_df.columns]

    def pick(cols, candidates):
        for c in candidates:
            if c in cols:
                return c
        return None

    code_col   = pick(mapping_df.columns, ["项目编码", "项目代码", "项目编号", "编码"])
    name_col   = pick(mapping_df.columns, [
        "项目名称（合同管理）", "项目名称(合同管理)", "合同管理项目名称",
        "项目名称（成本梳理）", "项目名称(成本梳理)", "成本梳理项目名称",
        "项目名称", "标准项目名称",
    ])
    status_col = pick(mapping_df.columns, ["备注", "项目状态", "状态"])

    matcher_df = pd.DataFrame({
        "项目名称": mapping_df[name_col],
        "项目编码": mapping_df[code_col] if code_col else "",
        "备注":     mapping_df[status_col] if status_col else "需手动核对",
    })

    standard_projects = build_standard_projects_from_dataframe(matcher_df)
    match_db = load_or_init_db(Config.MATCH_DB_PATH, standard_projects)

    def _match(raw_name: str):
        raw = str(raw_name).strip()
        if raw.lower() in INVALID_NAMES:
            return ("", raw)
        hit = match_project(raw, match_db) if match_db else None
        if hit:
            return (
                hit.get("project_code", "") or "",
                hit.get("project_name", raw) or raw,
            )
        return ("", raw)

    results = df["原始项目名称"].apply(_match)
    df["项目编码"] = results.apply(lambda x: x[0])
    df["项目名称"] = results.apply(lambda x: x[1])

    # 收集未命中项目（编码为空且名称等于原始名称）
    unmatched_mask = (df["项目编码"] == "") & (df["项目名称"] == df["原始项目名称"])
    unmatched_raw_names = df.loc[unmatched_mask, "原始项目名称"].unique().tolist()
    unmatched_raw_names = [n for n in unmatched_raw_names if n.lower() not in INVALID_NAMES]

    if unmatched_raw_names:
        print(f"⚠️ 有 {len(unmatched_raw_names)} 个项目未自动匹配，弹窗手动选择...")
        manual_map = prompt_unmatched_project_mappings(
            unmatched_raw_names, match_db, db_path=Config.MATCH_DB_PATH
        )
        # 将手动匹配结果回写到 df
        for raw_name, std_name in manual_map.items():
            if not std_name:
                continue
            row_mask = df["原始项目名称"] == raw_name
            # 从 db 重新查一次以获取编码
            hit = match_project(raw_name, match_db)
            code = hit.get("project_code", "") or "" if hit else ""
            name = hit.get("project_name", std_name) or std_name if hit else std_name
            df.loc[row_mask, "项目编码"] = code
            df.loc[row_mask, "项目名称"] = name

    save_db(Config.MATCH_DB_PATH, match_db)
    print("✅ 项目名称标准化完成")
    return df


# ====================== 按项目拆分导出 ======================
def export_per_project_excel(df: pd.DataFrame, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    non_numeric = {"所属单位", "姓名", "订票时间", "年-月", "去程机票日期", "回程机票日期",
                   "本次机票航班号", "舱位", "机票类型", "票代", "护照", "项目编码", "项目名称",
                   "原始项目名称", "_来源文件"}

    projects = df[["项目编码", "项目名称"]].drop_duplicates()
    print(f"\n📂 按项目拆分输出（共 {len(projects)} 个项目）...")

    for _, row in projects.iterrows():
        code = str(row["项目编码"]).strip()
        name = str(row["项目名称"]).strip()

        prefix = f"{code}_{name}" if code and code != "nan" else name
        filename = _sanitize_filename(prefix) + "_机票.xlsx"
        out_path = output_dir / filename

        mask = (df["项目名称"] == row["项目名称"]) & (df["项目编码"] == row["项目编码"])
        proj_df = df[mask].copy()

        if proj_df.empty:
            continue

        display_cols = [c for c in Config.TARGET_COLUMNS if c in proj_df.columns]

        # 月度汇总
        monthly = (
            proj_df.groupby("年-月")[Config.MONEY_COL]
            .sum()
            .reset_index()
            .sort_values("年-月")
            .rename(columns={Config.MONEY_COL: "机票费用合计（泰铢）"})
        )

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            proj_df[display_cols].to_excel(writer, sheet_name="明细数据", index=False)
            monthly.to_excel(writer, sheet_name="月度汇总", index=False)

        wb = load_workbook(out_path)
        for ws in wb.worksheets:
            headers = [cell.value for cell in ws[1]]
            numeric_start = next(
                (i + 1 for i, h in enumerate(headers) if h and str(h) not in non_numeric),
                3,
            )
            apply_common_excel_layout(
                ws,
                numeric_start_col=numeric_start,
                freeze_panes="B2",
                thousand_number_format="#,##0.00",
            )
        wb.save(out_path)
        print(f"  ✅ {filename}")

    print(f"✅ 按项目拆分输出完成：{output_dir}")


# ====================== 主程序 ======================
def main():
    print("🚀 开始处理机票台账...")

    input_dir = Path(Config.INPUT_SOURCE)
    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在：{input_dir}")

    # 1. 读取所有原始文件
    df = load_all_raw(input_dir)
    print(f"✅ 原始数据共 {len(df)} 行")

    # 2. 列标准化
    df = normalize_columns(df)
    df = df.loc[:, ~df.columns.duplicated()]  # 去除重命名后可能产生的重复列

    # 3. 金额转换
    df[Config.MONEY_COL] = _parse_money(df[Config.MONEY_COL])

    # 3.1 人次 & 人均价格
    df["人次"] = _parse_person_count(df["姓名"])
    df["人均机票价格（泰铢）"] = (df[Config.MONEY_COL] / df["人次"]).round(2)

    # 4. 年-月列
    df["年-月"] = _parse_year_month(df["订票时间"])

    # 5. 清洗无效项目行
    df = df[df["项目"].apply(lambda x: str(x).strip().lower() not in INVALID_NAMES).values]
    df = df[~df["项目"].astype(str).str.contains("合计|总计|小计|Grand", case=False, na=False).values]
    df.reset_index(drop=True, inplace=True)

    # 6. 项目名称标准化
    df = standardize_projects(df)

    # ========== 聚合 ==========
    proj_summary = (
        df.groupby(["项目编码", "项目名称"], as_index=False)[Config.MONEY_COL]
        .sum()
        .rename(columns={Config.MONEY_COL: "机票费用合计（泰铢）"})
        .sort_values("机票费用合计（泰铢）", ascending=False)
    )

    monthly_pivot = df.pivot_table(
        index=["项目编码", "项目名称"],
        columns="年-月",
        values=Config.MONEY_COL,
        aggfunc="sum",
        fill_value=0,
    )
    if not monthly_pivot.empty:
        cols_sorted = sorted(
            monthly_pivot.columns,
            key=lambda s: (int(str(s).split("-")[0]), int(str(s).split("-")[1])) if "-" in str(s) else (0, 0),
        )
        monthly_pivot = monthly_pivot.reindex(columns=cols_sorted)

    # ========== 输出主汇总 Excel ==========
    out_dir = Path(Config.OUTPUT_STD)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "机票汇总.xlsx"

    display_cols = [c for c in Config.TARGET_COLUMNS if c in df.columns]

    non_numeric_headers = {"所属单位", "姓名", "订票时间", "年-月", "去程机票日期", "回程机票日期",
                           "本次机票航班号", "舱位", "机票类型", "票代", "护照", "项目编码", "项目名称",
                           "原始项目名称", "_来源文件"}}

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Sheet1: 基础汇总（所有行）
        df[display_cols].to_excel(writer, sheet_name="基础汇总", index=False)
        # Sheet2: 按项目汇总成本
        proj_summary.to_excel(writer, sheet_name="按项目汇总", index=False)
        # Sheet3: 按项目×月度透视
        if not monthly_pivot.empty:
            monthly_pivot.to_excel(writer, sheet_name="月度透视")

    wb = load_workbook(out_path)
    for ws in wb.worksheets:
        headers = [cell.value for cell in ws[1]]
        numeric_start = next(
            (i + 1 for i, h in enumerate(headers) if h and str(h) not in non_numeric_headers),
            3,
        )
        apply_common_excel_layout(
            ws,
            numeric_start_col=numeric_start,
            freeze_panes="B2",
            thousand_number_format="#,##0.00",
        )
    wb.save(out_path)
    print(f"✅ 主汇总已输出：{out_path}")

    # ========== 按项目拆分 Excel ==========
    export_per_project_excel(
        df=df,
        output_dir=out_dir / "按项目拆分",
    )

    print("\n🎉 机票台账处理完成！")


if __name__ == "__main__":
    main()
