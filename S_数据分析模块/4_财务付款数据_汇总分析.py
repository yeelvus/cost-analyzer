import os
import re
import sys
import shutil
import importlib.util
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook

# ====================== 配置区 ======================
class Config:
    INPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/B_财务付款_std"
    
    # 输出路径（同时支持两个位置）
    OUTPUT_Z = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report/财务付款数据汇总.xlsx"
    OUTPUT_ORIGINAL = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/C_财务付款_sum/财务付款数据汇总.xlsx"
    
    AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]


CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve_project_info_dir() -> Optional[str]:
    """探测项目基本信息目录"""
    candidates = [
        os.path.normpath(os.path.join(CURRENT_DIR, "..", "..", "0_项目基本信息")),
        os.path.normpath(os.path.join(CURRENT_DIR, "..", "0_项目基本信息")),
        "/Users/wyattsun/Desktop/code/0_项目基本信息",
    ]
    for p in candidates:
        if os.path.isdir(p):
            return p
    return None


PROJECT_INFO_DIR = _resolve_project_info_dir()
if PROJECT_INFO_DIR and PROJECT_INFO_DIR not in sys.path:
    sys.path.append(PROJECT_INFO_DIR)


def load_formatter():
    """加载 Excel 样式模块"""
    if not PROJECT_INFO_DIR:
        print("⚠️ 未找到项目信息目录，跳过样式设置。")
        return lambda *a, **k: None

    module_path = Path(PROJECT_INFO_DIR) / "excel_formatting_utils.py"
    if not module_path.exists():
        print(f"⚠️ 未找到格式模块: {module_path}")
        return lambda *a, **k: None

    try:
        spec = importlib.util.spec_from_file_location("excel_formatting_utils", module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module.apply_common_excel_layout
    except Exception as e:
        print(f"⚠️ 加载格式模块失败: {e}")
        return lambda *a, **k: None


APPLY_COMMON_EXCEL_LAYOUT = load_formatter()


# ====================== 核心函数 ======================
def parse_folder_tag(folder_name: str) -> str:
    match = re.search(r"(20\d{6})", folder_name)
    if match:
        return match.group(1)
    digits = "".join(ch for ch in folder_name if ch.isdigit())
    return digits[:8] if len(digits) >= 8 else folder_name


def safe_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def extract_project_summary(file_path: Path, ymd: str) -> Optional[Dict]:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"⚠️ 读取失败 {file_path.name}: {e}")
        return None

    if df.empty or "实付金额" not in df.columns:
        if "实付金额" not in df.columns and not df.empty:
            print(f"⚠️ 缺少实付金额列: {file_path.name}")
        return None

    # 过滤全零金额行
    amt_cols = [c for c in Config.AMOUNT_COLUMNS if c in df.columns]
    if amt_cols:
        numeric = df[amt_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
        df = df[numeric.sum(axis=1) > 0]

    if df.empty:
        return None

    # 项目编码 & 名称
    project_code = ""
    if "项目编码" in df.columns:
        codes = df["项目编码"].dropna().astype(str).str.strip()
        codes = codes[~codes.str.lower().isin(["", "nan", "none"])]
        if not codes.empty:
            project_code = codes.iloc[0]

    project_name = ""
    if "项目名称" in df.columns:
        names = df["项目名称"].dropna().astype(str).str.strip()
        names = names[~names.str.lower().isin(["", "nan", "none"])]
        if not names.empty:
            project_name = names.iloc[0]

    if not project_name:
        project_name = file_path.stem

    paid = pd.to_numeric(df["实付金额"], errors="coerce").fillna(0)
    vat = pd.to_numeric(df.get("增值税额", 0), errors="coerce").fillna(0)

    return {
        "项目编码": safe_text(project_code),
        "项目名称": safe_text(project_name),
        "年月日": ymd,
        "付款总金额": float((paid + vat).sum()),
    }


def build_summary_dataframe() -> pd.DataFrame:
    input_root = Path(Config.INPUT_ROOT)
    if not input_root.is_dir():
        raise FileNotFoundError(f"输入目录不存在: {input_root}")

    rows = []
    for folder in sorted(input_root.iterdir()):
        if not folder.is_dir():
            continue

        ymd = parse_folder_tag(folder.name)
        print(f"\n处理目录: {folder.name}")

        for file_path in folder.glob("*.xlsx"):
            if file_path.name.startswith("~$"):
                continue
            row = extract_project_summary(file_path, ymd)
            if row:
                rows.append(row)

    if not rows:
        return pd.DataFrame(columns=["项目编码", "项目名称"])

    df = pd.DataFrame(rows)

    # 聚合
    grouped = df.groupby(["项目名称", "年月日"], dropna=False).agg(
        项目编码=("项目编码", lambda s: next((v for v in s if safe_text(v)), "")),
        付款总金额=("付款总金额", "sum")
    ).reset_index()

    # 项目编码去重
    code_df = grouped.groupby("项目名称", dropna=False).agg(
        项目编码=("项目编码", lambda s: next((v for v in s if safe_text(v)), ""))
    ).reset_index()

    # 透视表
    pivot = grouped.pivot_table(
        index="项目名称", columns="年月日", values="付款总金额", aggfunc="sum"
    ).reset_index()

    date_cols = sorted(c for c in pivot.columns if c != "项目名称")
    pivot = pivot[["项目名称"] + date_cols]

    result = code_df.merge(pivot, on="项目名称", how="left")
    return result[["项目编码", "项目名称"] + date_cols]


def build_increase_sheet(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    inc_df = df.copy()
    date_cols = [c for c in df.columns if c not in ("项目编码", "项目名称")]
    if not date_cols:
        return inc_df

    numeric = {col: pd.to_numeric(df[col], errors="coerce") for col in date_cols}

    for i, col in enumerate(date_cols):
        if i == 0:
            inc_df[col] = "-"
            continue
        prev = date_cols[i - 1]
        diff = numeric[col] - numeric[prev]
        both_valid = numeric[col].notna() & numeric[prev].notna()
        inc_df[col] = diff.where(both_valid, "-")

    return inc_df


def apply_formatting(output_path: str):
    try:
        wb = load_workbook(output_path)
        for sheet in ["付款总金额汇总", "较上期增加汇总"]:
            ws = wb[sheet]
            headers = [cell.value for cell in ws[1]]
            amount_headers = headers[2:]

            APPLY_COMMON_EXCEL_LAYOUT(
                ws,
                numeric_start_col=ws.max_column + 1,
                amount_headers=amount_headers,
                header_row=1,
                header_height=40,
                freeze_panes="A2",
            )
        wb.save(output_path)
        print(f"✅ 样式应用完成: {Path(output_path).name}")
    except Exception as e:
        print(f"⚠️ 样式应用失败: {e}")


def main():
    print("🚀 开始处理财务付款数据汇总...")

    summary_df = build_summary_dataframe()
    increase_df = build_increase_sheet(summary_df)

    # 输出策略：先生成主文件，再复制
    output_paths = [Config.OUTPUT_Z, Config.OUTPUT_ORIGINAL]

    for path_str in output_paths:
        path = Path(path_str)
        path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="付款总金额汇总", index=False)
            increase_df.to_excel(writer, sheet_name="较上期增加汇总", index=False)

        apply_formatting(str(path))
        print(f"✅ 输出完成 → {path}")

    print("\n🎉 全部处理完成！")
    print(f"项目数量: {len(summary_df)}")
    print(f"日期列数: {len([c for c in summary_df.columns if c not in ('项目编码','项目名称')])}")


if __name__ == "__main__":
    main()