# -*- coding: utf-8 -*-
"""机具费 HTML / Excel 分析报告（优化版）。"""
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

from machine.config import FULL_SUMMARY_FILENAME, LEGACY_SUM_DIR, REPORT_DIR, STD_SUM_DIR, SUM_DIR, ensure_data_dirs

warnings.filterwarnings("ignore")

# 配色
BRAND = "#0F3D4C"
ACCENT = "#1A6B7A"
GREEN = "#2B8A6E"
AMBER = "#C4892A"
CORAL = "#E8684A"
PURPLE = "#5B8FF9"


def find_latest_summary(sum_dir: Path | None = None) -> Path:
    ensure_data_dirs()
    dirs = []
    if sum_dir:
        dirs.append(Path(sum_dir))
    dirs.extend([STD_SUM_DIR, SUM_DIR, LEGACY_SUM_DIR])

    for d in dirs:
        if not d.exists():
            continue
        preferred = d / FULL_SUMMARY_FILENAME
        if preferred.exists():
            return preferred
        files = [
            Path(f)
            for f in glob.glob(str(d / "*全量标准化*汇总*.xlsx"))
            if not Path(f).name.startswith("~$")
        ]
        if not files:
            files = [
                Path(f)
                for f in glob.glob(str(d / "*汇总*.xlsx"))
                if not Path(f).name.startswith("~$") and "TEST_" not in Path(f).name
            ]
        if files:
            return max(files, key=lambda p: p.stat().st_mtime)
    raise FileNotFoundError(f"未找到标准汇总文件，已查: {dirs}")


def detect_header_row(df_raw: pd.DataFrame) -> int:
    for idx, row in df_raw.iterrows():
        row_str = "".join(row.astype(str)).lower()
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
    if re.match(r"20\d{2}-\d{2}$", month_str):
        try:
            return pd.to_datetime(month_str + "-01")
        except Exception:
            pass
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
    try:
        return pd.to_datetime(month_str)
    except Exception:
        return None


def load_summary_dataframe(excel_path: Path) -> tuple[pd.DataFrame, dict]:
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
    # 若首行已是表头（汇总置首后通常 header=0）
    if header_row == 0 and any("机具" in str(v) or "项目" in str(v) for v in raw.iloc[0].tolist()):
        df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
    else:
        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    if "标准机具种类" not in df.columns and "机具种类" in df.columns:
        df = df.rename(columns={"机具种类": "标准机具种类"})
    if "标准型号规格" in df.columns:
        if "型号规格" in df.columns:
            df = df.drop(columns=["型号规格"])
        df = df.rename(columns={"标准型号规格": "型号规格"})

    groups = {
        "区域": ["区域", "项目区域"],
        "标准机具种类": ["标准机具种类", "机具种类"],
        "月份": ["月份", "年月"],
        "台班费_泰铢": ["台班费_泰铢", "台班费"],
        "本次结算台班": ["本次结算台班", "结算台班"],
        "结算费用_泰铢": ["结算费用_泰铢", "本次费用合计", "费用合计"],
        "本次费用合计": ["本次费用合计"],
        "出租单位": ["出租单位", "供应商"],
        "项目名称": ["项目名称", "项目"],
        "项目编码": ["项目编码", "编码"],
    }
    rename = {}
    for std, cands in groups.items():
        found = smart_find_column(df, cands, std)
        if found and found != std and std not in df.columns:
            rename[found] = std
    if rename:
        df = df.rename(columns=rename)

    amount_col = "本次费用合计" if "本次费用合计" in df.columns else (
        "结算费用_泰铢" if "结算费用_泰铢" in df.columns else None
    )
    if amount_col is None:
        raise KeyError(f"缺少金额列，当前列: {list(df.columns)}")

    if "区域" not in df.columns:
        df["区域"] = "未知区域"
    if "型号规格" not in df.columns:
        df["型号规格"] = ""
    if "出租单位" not in df.columns:
        df["出租单位"] = "未知"
    if "项目名称" not in df.columns:
        df["项目名称"] = "未知项目"

    # 月份：优先年月
    if "年月" in df.columns and ("月份" not in df.columns or df["月份"].isna().all()):
        df["月份"] = df["年月"]
    if "月份" in df.columns:
        df["月份"] = df["月份"].apply(parse_month)
        df = df.dropna(subset=["月份"])
    else:
        df["月份"] = pd.Timestamp(datetime.now().year, datetime.now().month, 1)

    for col in ["台班费_泰铢", "本次结算台班", "结算费用_泰铢", "本次费用合计"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["_金额"] = df[amount_col]
    return df, extra


def kpi_cards(df: pd.DataFrame) -> dict:
    return {
        "记录数": len(df),
        "项目数": int(df["项目名称"].nunique()) if "项目名称" in df.columns else 0,
        "机具种类数": int(df["标准机具种类"].nunique()),
        "出租单位数": int(df["出租单位"].nunique()) if "出租单位" in df.columns else 0,
        "费用合计_泰铢": float(df["_金额"].sum()),
        "月数": int(df["月份"].dt.to_period("M").nunique()) if "月份" in df.columns else 0,
    }


def build_figures(df: pd.DataFrame) -> list:
    figs = []
    # 1 月度趋势
    monthly = (
        df.groupby(df["月份"].dt.to_period("M"))["_金额"]
        .sum()
        .reset_index()
    )
    monthly["月份"] = monthly["月份"].astype(str)
    fig = px.area(
        monthly, x="月份", y="_金额",
        title="① 月度机具费用趋势",
        color_discrete_sequence=[ACCENT],
    )
    fig.update_traces(line_color=ACCENT, fillcolor="rgba(26,107,122,0.25)")
    fig.update_layout(
        template="plotly_white",
        yaxis_title="金额（泰铢）",
        margin=dict(l=50, r=20, t=50, b=40),
        height=360,
    )
    figs.append(("月度趋势", fig))

    # 2 项目 Top15
    proj = (
        df.groupby("项目名称")["_金额"]
        .sum()
        .sort_values(ascending=False)
        .head(15)
        .reset_index()
    )
    fig = px.bar(
        proj, x="_金额", y="项目名称", orientation="h",
        title="② 项目费用 Top15",
        color="_金额",
        color_continuous_scale=["#E8F5F3", GREEN],
    )
    fig.update_layout(
        template="plotly_white",
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        coloraxis_showscale=False,
        margin=dict(l=120, r=20, t=50, b=40),
        height=420,
    )
    figs.append(("项目Top", fig))

    # 3 机具种类 Top15
    kind = (
        df.groupby("标准机具种类")["_金额"]
        .sum()
        .sort_values(ascending=False)
        .head(15)
        .reset_index()
    )
    fig = px.bar(
        kind, x="_金额", y="标准机具种类", orientation="h",
        title="③ 机具种类费用 Top15",
        color_discrete_sequence=[BRAND],
    )
    fig.update_layout(
        template="plotly_white",
        yaxis={"categoryorder": "total ascending"},
        margin=dict(l=120, r=20, t=50, b=40),
        height=420,
    )
    figs.append(("机具种类", fig))

    # 4 区域
    if "区域" in df.columns:
        reg = df.groupby("区域")["_金额"].sum().reset_index()
        reg = reg[reg["区域"].astype(str).str.strip().isin(["", "未知区域", "未知"]) == False]
        if len(reg):
            fig = px.pie(
                reg, names="区域", values="_金额", title="④ 区域费用占比",
                color_discrete_sequence=px.colors.qualitative.Set2,
                hole=0.4,
            )
            fig.update_layout(template="plotly_white", height=380, margin=dict(t=50, b=20))
            figs.append(("区域", fig))

    # 5 出租单位 Top12
    if "出租单位" in df.columns:
        sup = (
            df.groupby("出租单位")["_金额"]
            .sum()
            .sort_values(ascending=False)
            .head(12)
            .reset_index()
        )
        fig = px.bar(
            sup, x="_金额", y="出租单位", orientation="h",
            title="⑤ 出租单位费用 Top12",
            color_discrete_sequence=[AMBER],
        )
        fig.update_layout(
            template="plotly_white",
            yaxis={"categoryorder": "total ascending"},
            margin=dict(l=140, r=20, t=50, b=40),
            height=400,
        )
        figs.append(("出租单位", fig))

    # 6 项目×月份热力（Top 项目）
    top_projects = (
        df.groupby("项目名称")["_金额"].sum().sort_values(ascending=False).head(12).index.tolist()
    )
    heat_src = df[df["项目名称"].isin(top_projects)].copy()
    if len(heat_src):
        heat_src["月"] = heat_src["月份"].dt.strftime("%Y-%m")
        pivot = heat_src.pivot_table(
            index="项目名称", columns="月", values="_金额", aggfunc="sum", fill_value=0
        )
        pivot = pivot.reindex(columns=sorted(pivot.columns))
        fig = go.Figure(
            data=go.Heatmap(
                z=pivot.values,
                x=list(pivot.columns),
                y=list(pivot.index),
                colorscale="Teal",
                colorbar=dict(title="泰铢"),
            )
        )
        fig.update_layout(
            title="⑥ 重点项目 × 月份费用热力图",
            template="plotly_white",
            height=420,
            margin=dict(l=120, r=20, t=50, b=40),
        )
        figs.append(("热力图", fig))

    return figs


def write_excel_summary(df: pd.DataFrame, path: Path, extra: dict, kpis: dict):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # 总览置首
        overview = pd.DataFrame({
            "指标": list(kpis.keys()),
            "值": list(kpis.values()),
        })
        overview.to_excel(writer, sheet_name="总览", index=False)

        df.groupby("标准机具种类")["_金额"].sum().sort_values(ascending=False).reset_index().rename(
            columns={"_金额": "费用合计(泰铢)"}
        ).to_excel(writer, sheet_name="按机具种类", index=False)

        df.groupby(["项目编码", "项目名称"] if "项目编码" in df.columns else ["项目名称"])["_金额"].sum().sort_values(
            ascending=False
        ).reset_index().rename(columns={"_金额": "费用合计(泰铢)"}).to_excel(
            writer, sheet_name="按项目", index=False
        )

        if "出租单位" in df.columns:
            df.groupby("出租单位")["_金额"].sum().sort_values(ascending=False).reset_index().rename(
                columns={"_金额": "费用合计(泰铢)"}
            ).to_excel(writer, sheet_name="按出租单位", index=False)

        m = df.copy()
        m["月"] = m["月份"].dt.strftime("%Y-%m")
        m.groupby("月")["_金额"].sum().reset_index().rename(
            columns={"_金额": "费用合计(泰铢)"}
        ).to_excel(writer, sheet_name="按月", index=False)

        for name, edf in extra.items():
            if name in ("汇总",):
                continue
            try:
                edf.to_excel(writer, sheet_name=str(name)[:31], index=False)
            except Exception:
                pass

        export = df.drop(columns=["_金额"], errors="ignore")
        export.to_excel(writer, sheet_name="明细数据", index=False)


def run_report(excel_path: Path | None = None, output_dir: Path | None = None) -> tuple[Path, Path]:
    ensure_data_dirs()
    excel_path = Path(excel_path) if excel_path else find_latest_summary()
    output_dir = Path(output_dir or REPORT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"使用标准汇总: {excel_path}")

    df, extra = load_summary_dataframe(excel_path)
    kpis = kpi_cards(df)
    print(f"读取 {len(df)} 行 | 费用合计 {kpis['费用合计_泰铢']:,.0f} 泰铢 | 项目 {kpis['项目数']}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = output_dir / f"机具租赁费用分析报告_{ts}.html"
    xlsx_path = output_dir / f"机具租赁统计汇总_{ts}.xlsx"

    figs = build_figures(df)
    total_wan = kpis["费用合计_泰铢"] / 10000

    cards_html = f"""
    <div class="kpis">
      <div class="kpi"><div class="l">费用合计</div><div class="v">{total_wan:,.1f}</div><div class="u">万泰铢</div></div>
      <div class="kpi"><div class="l">明细记录</div><div class="v">{kpis['记录数']:,}</div><div class="u">行</div></div>
      <div class="kpi"><div class="l">项目数</div><div class="v">{kpis['项目数']}</div><div class="u">个</div></div>
      <div class="kpi"><div class="l">机具种类</div><div class="v">{kpis['机具种类数']}</div><div class="u">种</div></div>
      <div class="kpi"><div class="l">出租单位</div><div class="v">{kpis['出租单位数']}</div><div class="u">家</div></div>
      <div class="kpi"><div class="l">覆盖月份</div><div class="v">{kpis['月数']}</div><div class="u">月</div></div>
    </div>
    """

    parts = [
        "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'/>",
        "<meta name='viewport' content='width=device-width, initial-scale=1'/>",
        "<title>机具租赁费用分析报告</title>",
        """<style>
:root { --bg:#eef2f1; --card:#fff; --text:#15201c; --muted:#667872; --line:#dde5e1;
  --brand:#0f3d4c; --brand2:#1a6b7a; }
* { box-sizing:border-box; }
body { margin:0; font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;
  background:var(--bg); color:var(--text); }
.wrap { max-width:1280px; margin:0 auto; padding:20px 18px 48px; }
.hero { background:linear-gradient(125deg,var(--brand),var(--brand2) 55%,#1f8a65);
  color:#fff; border-radius:18px; padding:22px 26px; margin-bottom:14px;
  box-shadow:0 10px 28px rgba(15,61,76,.18); }
.hero h1 { margin:0 0 8px; font-size:22px; letter-spacing:.02em; }
.hero p { margin:4px 0; font-size:13px; opacity:.93; }
.kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:10px; margin-bottom:14px; }
.kpi { background:var(--card); border:1px solid var(--line); border-radius:14px; padding:12px 14px; }
.kpi .l { font-size:11px; color:var(--muted); }
.kpi .v { font-size:20px; font-weight:700; margin-top:4px; color:var(--brand); }
.kpi .u { font-size:11px; color:var(--muted); margin-top:2px; }
.panel { background:var(--card); border:1px solid var(--line); border-radius:14px;
  padding:14px 16px; margin-bottom:12px; box-shadow:0 2px 8px rgba(0,0,0,.04); }
.panel h2 { margin:0 0 8px; font-size:15px; color:var(--brand); }
.grid2 { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
@media (max-width:900px) { .grid2 { grid-template-columns:1fr; } }
.meta { font-size:12px; color:var(--muted); }
footer { margin-top:18px; font-size:12px; color:var(--muted); line-height:1.55; }
</style></head><body><div class="wrap">""",
        f"""<div class="hero">
  <h1>机具租赁费用分析报告</h1>
  <p>数据源：{excel_path.name}</p>
  <p>生成时间：{datetime.now():%Y-%m-%d %H:%M} · 单位：泰铢（KPI 显示万泰铢）</p>
</div>""",
        cards_html,
    ]

    # 图表：趋势/热力全宽，其余两两并排
    pair_buf = []
    for i, (title, fig) in enumerate(figs):
        include_js = "cdn" if i == 0 else False
        chart_html = fig.to_html(full_html=False, include_plotlyjs=include_js)
        full_width = title in ("月度趋势", "热力图") or i == 0
        if full_width:
            if pair_buf:
                parts.append("<div class='grid2'>" + "".join(pair_buf) + "</div>")
                pair_buf = []
            parts.append(f"<div class='panel'>{chart_html}</div>")
        else:
            pair_buf.append(f"<div class='panel'>{chart_html}</div>")
            if len(pair_buf) == 2:
                parts.append("<div class='grid2'>" + "".join(pair_buf) + "</div>")
                pair_buf = []
    if pair_buf:
        parts.append("<div class='grid2'>" + "".join(pair_buf) + "</div>")

    parts.append(
        f"<footer>源文件：{excel_path}<br/>"
        f"流水线：B_raw_sum → C_raw_split → D_std_sum → Z_report/机具费分析 · "
        f"Excel 统计见同目录 {xlsx_path.name}</footer>"
    )
    parts.append("</div></body></html>")
    html_path.write_text("\n".join(parts), encoding="utf-8")

    write_excel_summary(df, xlsx_path, extra, kpis)

    for pattern in ("机具租赁费用分析报告_*.html", "机具租赁统计汇总_*.xlsx"):
        files = sorted(glob.glob(str(output_dir / pattern)))
        for old in files[:-1]:
            try:
                os.remove(old)
            except Exception:
                pass

    print("=" * 60)
    print(f"HTML 报告: {html_path}")
    print(f"Excel 统计: {xlsx_path}")
    print("=" * 60)
    return html_path, xlsx_path


def main(argv=None):
    run_report()


if __name__ == "__main__":
    main()
