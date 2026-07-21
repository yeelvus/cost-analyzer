# -*- coding: utf-8 -*-
"""机具费 HTML / Excel 分析报告。"""
from __future__ import annotations

import glob
import os
import re
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from machine.config import REPORT_DIR, SUM_DIR

warnings.filterwarnings("ignore")


def find_latest_summary(sum_dir: Path | None = None) -> Path:
    sum_dir = Path(sum_dir or SUM_DIR)
    files = [Path(f) for f in glob.glob(str(sum_dir / "*已汇总*.xlsx")) if not Path(f).name.startswith("~$")]
    if not files:
        files = [Path(f) for f in glob.glob(str(sum_dir / "*汇总*.xlsx")) if not Path(f).name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"未找到汇总文件: {sum_dir}")
    return max(files, key=lambda p: p.stat().st_mtime)


def detect_header_row(df_raw: pd.DataFrame) -> int:
    for idx, row in df_raw.iterrows():
        row_str = "".join(row.astype(str)).lower()
        if "区域" in row_str and ("标准机具种类" in row_str or "机具种类" in row_str):
            return int(idx) if isinstance(idx, (int, float)) else 0
        if "机具种类" in row_str and ("台班" in row_str or "项目" in row_str):
            return int(idx) if isinstance(idx, (int, float)) else 0
    return 0


def smart_find_column(df, possible_names, target_name):
    if target_name in df.columns:
        return target_name
    for name in possible_names:
        for col in df.columns:
            a = name.lower().replace("_", "").replace(" ", "")
            b = str(col).lower().replace("_", "").replace(" ", "")
            if a in b:
                return col
    return None


def parse_month(month_val):
    if pd.isna(month_val):
        return None
    month_str = str(month_val).strip()
    if "月" in month_str:
        m = re.search(r"(\d+)", month_str)
        if m:
            try:
                return pd.Timestamp(year=datetime.now().year, month=int(m.group(1)), day=1)
            except Exception:
                return None
    if len(month_str) == 6 and month_str.isdigit():
        try:
            return pd.to_datetime(month_str, format="%Y%m")
        except Exception:
            pass
    # YYYY-MM
    if re.match(r"20\d{2}-\d{2}$", month_str):
        try:
            return pd.to_datetime(month_str + "-01")
        except Exception:
            pass
    try:
        return pd.to_datetime(month_str)
    except Exception:
        return None


def load_summary_dataframe(excel_path: Path) -> tuple[pd.DataFrame, dict]:
    """返回 (明细df, 附加sheet字典)。"""
    xl = pd.ExcelFile(excel_path)
    extra = {}
    for name in xl.sheet_names:
        if name != "汇总":
            try:
                extra[name] = pd.read_excel(excel_path, sheet_name=name)
            except Exception:
                pass

    sheet = "汇总" if "汇总" in xl.sheet_names else xl.sheet_names[0]
    raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)
    header_row = detect_header_row(raw)
    df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    if "来源Sheet" in df.columns and df["来源Sheet"].isna().all():
        df = df.drop(columns=["来源Sheet"])

    # 型号规格优先标准列
    if "标准型号规格" in df.columns:
        if "型号规格" in df.columns:
            df = df.drop(columns=["型号规格"])
        df = df.rename(columns={"标准型号规格": "型号规格"})
    if "标准机具种类" not in df.columns and "机具种类" in df.columns:
        df = df.rename(columns={"机具种类": "标准机具种类"})

    groups = {
        "区域": ["区域", "项目区域", "所在区域"],
        "标准机具种类": ["标准机具种类", "机具种类", "种类"],
        "月份": ["月份", "开始月份", "使用月份", "结算月份", "年月"],
        "台班费_泰铢": ["台班费_泰铢", "台班费", "单价"],
        "本次结算台班": ["本次结算台班", "结算台班", "台班数量", "数量"],
        "结算费用_泰铢": ["结算费用_泰铢", "本次费用合计", "费用合计"],
        "出租单位": ["出租单位", "供应商", "租赁单位"],
        "项目名称": ["项目名称", "项目"],
    }
    rename = {}
    for std, cands in groups.items():
        found = smart_find_column(df, cands, std)
        if found and found != std and std not in df.columns:
            rename[found] = std
    if rename:
        df = df.rename(columns=rename)

    if "区域" not in df.columns:
        df["区域"] = "未知区域"
        if "处理报告" in extra.columns if False else "处理报告" in extra:
            rep = extra["处理报告"]
            if "项目名称" in rep.columns and "区域" in rep.columns and "项目名称" in df.columns:
                mp = dict(zip(rep["项目名称"].astype(str), rep["区域"].astype(str)))
                df["区域"] = df["项目名称"].astype(str).map(mp).fillna("未知区域")

    # 金额列：优先本次费用合计
    if "结算费用_泰铢" not in df.columns and "本次费用合计" in df.columns:
        df["结算费用_泰铢"] = df["本次费用合计"]

    required = ["标准机具种类"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"缺少必要列 {missing}，当前列: {list(df.columns)}")

    if "月份" in df.columns:
        df["月份"] = df["月份"].apply(parse_month)
        df = df.dropna(subset=["月份"])
    else:
        df["月份"] = pd.Timestamp(datetime.now().year, datetime.now().month, 1)

    for col in ["台班费_泰铢", "本次结算台班", "结算费用_泰铢"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "型号规格" not in df.columns:
        df["型号规格"] = ""
    if "出租单位" not in df.columns:
        df["出租单位"] = "未知"
    if "项目名称" not in df.columns:
        df["项目名称"] = "未知项目"

    return df, extra


def build_figures(df: pd.DataFrame) -> list:
    figs = []
    amount_col = "结算费用_泰铢" if "结算费用_泰铢" in df.columns else None
    if amount_col is None:
        return figs

    # 1 按月趋势
    monthly = df.groupby(df["月份"].dt.to_period("M"))[amount_col].sum().reset_index()
    monthly["月份"] = monthly["月份"].astype(str)
    fig = px.bar(monthly, x="月份", y=amount_col, title="月度机具费用趋势（泰铢）",
                 color_discrete_sequence=["#1A6B7A"])
    fig.update_layout(template="plotly_white")
    figs.append(fig)

    # 2 机具种类 Top15
    kind = df.groupby("标准机具种类")[amount_col].sum().sort_values(ascending=False).head(15).reset_index()
    fig = px.bar(kind, x=amount_col, y="标准机具种类", orientation="h",
                 title="机具种类费用 Top15", color_discrete_sequence=["#0F3D4C"])
    fig.update_layout(template="plotly_white", yaxis={"categoryorder": "total ascending"})
    figs.append(fig)

    # 3 项目费用
    proj = df.groupby("项目名称")[amount_col].sum().sort_values(ascending=False).head(20).reset_index()
    fig = px.bar(proj, x="项目名称", y=amount_col, title="项目费用 Top20",
                 color_discrete_sequence=["#2B8A6E"])
    fig.update_layout(template="plotly_white", xaxis_tickangle=-35)
    figs.append(fig)

    # 4 区域饼图
    if "区域" in df.columns:
        reg = df.groupby("区域")[amount_col].sum().reset_index()
        reg = reg[reg["区域"].astype(str).str.strip() != ""]
        if len(reg):
            fig = px.pie(reg, names="区域", values=amount_col, title="区域费用占比")
            fig.update_layout(template="plotly_white")
            figs.append(fig)

    # 5 供应商 Top15
    if "出租单位" in df.columns:
        sup = df.groupby("出租单位")[amount_col].sum().sort_values(ascending=False).head(15).reset_index()
        fig = px.bar(sup, x=amount_col, y="出租单位", orientation="h", title="出租单位费用 Top15",
                     color_discrete_sequence=["#C4892A"])
        fig.update_layout(template="plotly_white", yaxis={"categoryorder": "total ascending"})
        figs.append(fig)

    return figs


def write_excel_summary(df: pd.DataFrame, path: Path, extra: dict):
    amount_col = "结算费用_泰铢"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # 汇总
        overview = pd.DataFrame({
            "指标": ["记录数", "项目数", "机具种类数", "出租单位数", "费用合计(泰铢)"],
            "值": [
                len(df),
                df["项目名称"].nunique() if "项目名称" in df.columns else 0,
                df["标准机具种类"].nunique(),
                df["出租单位"].nunique() if "出租单位" in df.columns else 0,
                float(df[amount_col].sum()) if amount_col in df.columns else 0,
            ],
        })
        overview.to_excel(writer, sheet_name="总览", index=False)

        if amount_col in df.columns:
            df.groupby("标准机具种类")[amount_col].sum().sort_values(ascending=False).reset_index().to_excel(
                writer, sheet_name="按机具种类", index=False
            )
            df.groupby("项目名称")[amount_col].sum().sort_values(ascending=False).reset_index().to_excel(
                writer, sheet_name="按项目", index=False
            )
            if "出租单位" in df.columns:
                df.groupby("出租单位")[amount_col].sum().sort_values(ascending=False).reset_index().to_excel(
                    writer, sheet_name="按出租单位", index=False
                )
            m = df.copy()
            m["月"] = m["月份"].dt.strftime("%Y-%m")
            m.groupby("月")[amount_col].sum().reset_index().to_excel(writer, sheet_name="按月", index=False)

        # 附带处理报告等
        for name, edf in extra.items():
            if name in ("汇总",):
                continue
            safe = str(name)[:31]
            try:
                edf.to_excel(writer, sheet_name=safe, index=False)
            except Exception:
                pass

        df.to_excel(writer, sheet_name="明细数据", index=False)


def run_report(excel_path: Path | None = None, output_dir: Path | None = None) -> tuple[Path, Path]:
    excel_path = Path(excel_path) if excel_path else find_latest_summary()
    output_dir = Path(output_dir or REPORT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"使用文件: {excel_path.name}")

    df, extra = load_summary_dataframe(excel_path)
    print(f"成功读取: {len(df)} 行, 列={list(df.columns)[:12]}...")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = output_dir / f"机具租赁费用分析报告_{ts}.html"
    xlsx_path = output_dir / f"机具租赁统计汇总_{ts}.xlsx"

    figs = build_figures(df)
    parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'/>",
        f"<title>机具租赁费用分析报告</title>",
        "<style>body{font-family:PingFang SC,Microsoft YaHei,sans-serif;margin:24px;background:#eef2f1;}",
        ".card{background:#fff;border-radius:12px;padding:16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,.06);}",
        "h1{color:#0f3d4c;} .meta{color:#667872;font-size:13px;}</style></head><body>",
        f"<div class='card'><h1>机具租赁费用分析报告</h1>",
        f"<p class='meta'>源文件: {excel_path.name} · 生成: {datetime.now():%Y-%m-%d %H:%M} · 记录数: {len(df)}</p></div>",
    ]
    for i, fig in enumerate(figs):
        parts.append(f"<div class='card'>{fig.to_html(full_html=False, include_plotlyjs=('cdn' if i==0 else False))}</div>")
    parts.append("</body></html>")
    html_path.write_text("\n".join(parts), encoding="utf-8")

    write_excel_summary(df, xlsx_path, extra)

    # 清理旧报告（保留最新）
    for pattern in ("机具租赁费用分析报告_*.html", "机具租赁统计汇总_*.xlsx"):
        files = sorted(glob.glob(str(output_dir / pattern)))
        for old in files[:-1]:
            try:
                os.remove(old)
                print(f"  已删除旧文件: {Path(old).name}")
            except Exception:
                pass

    print("=" * 60)
    print(f"HTML 报告: {html_path}")
    print(f"Excel 汇总: {xlsx_path}")
    print("=" * 60)
    return html_path, xlsx_path


def main(argv=None):
    run_report()


if __name__ == "__main__":
    main()
