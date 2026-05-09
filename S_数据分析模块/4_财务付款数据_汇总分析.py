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
        if spec is None or spec.loader is None:
            print(f"⚠️ 无法创建模块加载器: {module_path}")
            return lambda *a, **k: None
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


def ymd_sort_key(value: str) -> int:
    text = safe_text(value)
    if re.fullmatch(r"\d{8}", text):
        return int(text)
    if re.fullmatch(r"\d{6}", text):
        return int(f"{text}01")
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 8:
        return int(digits[:8])
    if len(digits) >= 6:
        return int(f"{digits[:6]}01")
    return -1


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

    amount_totals = {}
    for col in Config.AMOUNT_COLUMNS:
        source = df[col] if col in df.columns else pd.Series([0] * len(df), index=df.index)
        amount_totals[col] = float(pd.to_numeric(source, errors="coerce").fillna(0).sum())

    row = {
        "项目编码": safe_text(project_code),
        "项目名称": safe_text(project_name),
        "年月日": ymd,
    }
    row.update(amount_totals)
    return row


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
        return pd.DataFrame(columns=["项目编码", "项目名称", "最新月份"] + Config.AMOUNT_COLUMNS)

    df = pd.DataFrame(rows)

    # 先按项目+月份聚合每月数据
    agg_map = {col: (col, "sum") for col in Config.AMOUNT_COLUMNS}
    grouped = df.groupby(["项目名称", "年月日"], dropna=False).agg(
        项目编码=("项目编码", lambda s: next((v for v in s if safe_text(v)), "")),
        **agg_map,
    ).reset_index()

    # 每个项目仅保留最新月份
    grouped["_ymd_sort"] = grouped["年月日"].astype(str).map(ymd_sort_key)
    latest = (
        grouped.sort_values(["项目名称", "_ymd_sort", "年月日"], ascending=[True, True, True])
        .groupby("项目名称", dropna=False, as_index=False)
        .tail(1)
        .drop(columns=["_ymd_sort"])
        .rename(columns={"年月日": "最新月份"})
    )

    return latest[["项目编码", "项目名称", "最新月份"] + Config.AMOUNT_COLUMNS].reset_index(drop=True)


def apply_formatting(output_path: str):
    try:
        wb = load_workbook(output_path)
        for sheet in ["最新月份汇总"]:
            ws = wb[sheet]
            headers = [cell.value for cell in ws[1]]
            amount_headers = [h for h in headers if h in Config.AMOUNT_COLUMNS]

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

    # 输出策略：先生成主文件，再复制
    output_paths = [Config.OUTPUT_Z, Config.OUTPUT_ORIGINAL]

    for path_str in output_paths:
        path = Path(path_str)
        path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="最新月份汇总", index=False)

        apply_formatting(str(path))
        print(f"✅ 输出完成 → {path}")

    print("\n🎉 全部处理完成！")
    print(f"项目数量: {len(summary_df)}")
    print(f"输出列: {', '.join(summary_df.columns)}")


if __name__ == "__main__":
    main()