from __future__ import annotations

import argparse
import importlib.util
import math
import re
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

SCRIPT_DIR = Path(__file__).resolve().parent
_PROJECT_INFO_UTIL = SCRIPT_DIR.parents[2] / "0_项目基本信息" / "excel_formatting_utils.py"


def _load_module_from_path(module_name: str, module_path: Path):
  spec = importlib.util.spec_from_file_location(module_name, str(module_path))
  if spec is None or spec.loader is None:
    raise ImportError(f"无法加载模块: {module_path}")
  module = importlib.util.module_from_spec(spec)
  spec.loader.exec_module(module)
  return module


_excel_formatting_utils = _load_module_from_path("excel_formatting_utils", _PROJECT_INFO_UTIL)
apply_common_excel_layout = _excel_formatting_utils.apply_common_excel_layout


import sys as _sys

_PKG = Path(__file__).resolve().parent.parent
_CODE = _PKG.parent
for _p in (str(_PKG), str(_CODE)):
    if _p not in _sys.path:
        _sys.path.insert(0, _p)
try:
    from _采购路径 import ANALYSIS_DIR, MERGED_STD_NAME, SUM_DIR

    DEFAULT_SUMMARY_FILE = SUM_DIR / MERGED_STD_NAME
    SUMMARY_INPUT_ROOTS = [SUM_DIR]
    ANALYSIS_OUTPUT_ROOT = ANALYSIS_DIR
except Exception:
    _PROJECT_ROOT = Path.home() / "Desktop" / "邯郸中材" / "01成本统计"
    _DATA_ROOT = _PROJECT_ROOT / "9_采购台账📦"
    DEFAULT_SUMMARY_FILE = _DATA_ROOT / "D_采购部采购台账_sum" / "采购部采购台账_std_merged.xlsx"
    SUMMARY_INPUT_ROOTS = [
        _DATA_ROOT / "D_采购部采购台账_sum",
        _DATA_ROOT / "B_采购部采购台账_sum",
    ]
    ANALYSIS_OUTPUT_ROOT = _PROJECT_ROOT / "Z_report" / "C_采购数据分析"
INPUT_GLOB = "*std_merged*.xlsx"
DEFAULT_SHEET_NAME = "合并数据"
OUTPUT_SUFFIX = "_桩成本分析看板"

PILE_KEYWORD = "桩"
PILE_COST_PREFIX = "（打桩费）"
DRIVING_KEYWORDS = ("打桩", "施工")
SERVICE_EXCLUDE_KEYWORDS = ("运费", "引孔", "设备", "进场", "费用", "检测", "应变", "维修")

SPEC_PATTERN = re.compile(r"^(?P<width>\d+(?:\.\d+)?)x(?P<height>\d+(?:\.\d+)?)x(?P<length>\d+(?:\.\d+)?)m$", re.IGNORECASE)

COLOR_PILE = "#D97706"
COLOR_DRIVING = "#0F766E"
COLOR_MIN = "#1D4ED8"
COLOR_MAX = "#DC2626"
COLOR_TOTAL = "#7C3AED"

pio.templates.default = "simple_white"


def parse_args():
    parser = argparse.ArgumentParser(description="生成桩材与打桩费成本分析 HTML 看板")
    parser.add_argument(
        "input_file",
        nargs="?",
      default="",
      help="输入 Excel 文件路径；不传时自动读取 C_数据分析 下最新 *_材料采购_merged_标准化汇总.xlsx",
    )
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help="分析的工作表名称，默认是 合并数据",
    )
    parser.add_argument(
        "--output-file",
        default="",
        help="输出 HTML 文件路径；不传时自动写到输入文件同目录",
    )
    return parser.parse_args()


def build_output_path(input_file: Path, custom_output: str) -> tuple[Path, Path]:
    """返回 (html_path, excel_path)。"""
    if custom_output:
        html_p = Path(custom_output).expanduser().resolve()
        excel_p = html_p.with_suffix(".xlsx")
        return html_p, excel_p
    html_dir = ANALYSIS_OUTPUT_ROOT / "HTML"
    excel_dir = ANALYSIS_OUTPUT_ROOT / "Excel"
    html_dir.mkdir(parents=True, exist_ok=True)
    excel_dir.mkdir(parents=True, exist_ok=True)
    html_p = html_dir / f"{input_file.stem}{OUTPUT_SUFFIX}.html"
    excel_p = excel_dir / f"{input_file.stem}{OUTPUT_SUFFIX}.xlsx"
    return html_p, excel_p


def extract_year_month_from_name(path: Path) -> tuple[int, int] | None:
    m = re.search(r"(?<!\d)((?:19|20)\d{2})年(0?[1-9]|1[0-2])月", path.stem)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def find_latest_summary_file() -> Path | None:
    if DEFAULT_SUMMARY_FILE.is_file():
        return DEFAULT_SUMMARY_FILE

    for summary_root in SUMMARY_INPUT_ROOTS:
        if not summary_root.is_dir():
            continue
        candidates = [p for p in summary_root.glob(INPUT_GLOB) if p.is_file()]
        if not candidates:
            continue

        def sort_key(p: Path):
            ym = extract_year_month_from_name(p)
            if ym is not None:
                return (1, ym[0], ym[1], p.stat().st_mtime)
            return (0, 0, 0, p.stat().st_mtime)

        return max(candidates, key=sort_key)

    return None


def safe_float(value) -> float | None:
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(number):
        return None
    return float(number)


def format_number(value: float | int | None, digits: int = 2) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    if digits == 0:
        return f"{value:,.0f}"
    return f"{value:,.{digits}f}"


def format_dimension(value: float) -> str:
  rounded = round(value)
  if math.isclose(value, rounded, abs_tol=1e-9):
    return str(int(rounded))
  return f"{value:.3f}".rstrip("0").rstrip(".")


def classify_item(item_name: object) -> str | None:
    if item_name is None:
        return None

    item_text = str(item_name).strip()
    if not item_text or PILE_KEYWORD not in item_text:
        return None
    if item_text.startswith(PILE_COST_PREFIX) or any(keyword in item_text for keyword in DRIVING_KEYWORDS):
        return "打桩费"
    if any(keyword in item_text for keyword in SERVICE_EXCLUDE_KEYWORDS):
        return None
    if "费" in item_text:
        return None
    return "桩材"


def parse_spec(spec_text: object) -> dict[str, float] | None:
    if spec_text is None:
        return None

    text = str(spec_text).strip()
    if not text or text.lower() == "nan":
        return None

    match = SPEC_PATTERN.match(text)
    if not match:
        return None

    return {
        "宽mm": float(match.group("width")),
        "高mm": float(match.group("height")),
        "长度m": float(match.group("length")),
    }


def spec_sort_key(spec_text: str) -> tuple[float, float, float]:
    parsed = parse_spec(spec_text)
    if not parsed:
        return (float("inf"), float("inf"), float("inf"))
    return (parsed["宽mm"], parsed["高mm"], parsed["长度m"])


def section_sort_key(section_text: str) -> tuple[float, float]:
  if not isinstance(section_text, str) or "x" not in section_text:
    return (float("inf"), float("inf"))
  parts = section_text.split("x")
  if len(parts) != 2:
    return (float("inf"), float("inf"))
  try:
    return (float(parts[0]), float(parts[1]))
  except ValueError:
    return (float("inf"), float("inf"))


def weighted_average(frame: pd.DataFrame, value_col: str, weight_col: str) -> float | None:
    valid = frame[[value_col, weight_col]].dropna()
    valid = valid[valid[weight_col] > 0]
    if valid.empty:
        return None
    total_weight = valid[weight_col].sum()
    if total_weight == 0:
        return None
    return float((valid[value_col] * valid[weight_col]).sum() / total_weight)


def load_source_data(input_file: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # 兼容不同版本采购台账字段名
    if "标准项目名称" not in df.columns and "项目名称" in df.columns:
        df["标准项目名称"] = df["项目名称"]
    if "标准单位" not in df.columns and "单位" in df.columns:
        df["标准单位"] = df["单位"]

    required_columns = ["标准项目名称", "物品名称", "规格", "标准单位", "数量", "单价", "不含税金额", "供货商名称"]
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        raise ValueError(f"缺少必要列: {', '.join(missing_columns)}")
    return df


def prepare_analysis_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    pile_df = df[df["物品名称"].astype(str).str.contains(PILE_KEYWORD, na=False)].copy()
    pile_df["分类"] = pile_df["物品名称"].apply(classify_item)
    pile_df["标准单位"] = pile_df["标准单位"].astype(str).str.strip()
    pile_df["数量"] = pd.to_numeric(pile_df["数量"], errors="coerce")
    pile_df["单价"] = pd.to_numeric(pile_df["单价"], errors="coerce")
    pile_df["不含税金额"] = pd.to_numeric(pile_df["不含税金额"], errors="coerce")

    parsed_specs = pile_df["规格"].apply(parse_spec)
    pile_df["宽mm"] = parsed_specs.apply(lambda value: value["宽mm"] if value else None)
    pile_df["高mm"] = parsed_specs.apply(lambda value: value["高mm"] if value else None)
    pile_df["长度m"] = parsed_specs.apply(lambda value: value["长度m"] if value else None)

    stats = {
        "桩相关总行数": int(len(pile_df)),
        "标准单位为根": int((pile_df["标准单位"] == "根").sum()),
        "分类成功": int(pile_df["分类"].notna().sum()),
        "规格可解析": int(pile_df["长度m"].notna().sum()),
    }

    analysis_df = pile_df[
        pile_df["分类"].notna()
        & (pile_df["标准单位"] == "根")
        & pile_df["长度m"].notna()
        & pile_df["单价"].notna()
        & (pile_df["单价"] > 0)
        & pile_df["数量"].notna()
        & (pile_df["数量"] > 0)
    ].copy()

    analysis_df["每米成本"] = analysis_df["单价"] / analysis_df["长度m"]
    analysis_df["不含税金额"] = analysis_df["不含税金额"].fillna(analysis_df["单价"] * analysis_df["数量"])
    analysis_df["总米数"] = analysis_df["数量"] * analysis_df["长度m"]
    analysis_df["截面规格"] = analysis_df.apply(
      lambda row: f"{format_dimension(float(row['宽mm']))}x{format_dimension(float(row['高mm']))}", axis=1
    )

    stats["最终分析行数"] = int(len(analysis_df))
    stats["桩材分析行数"] = int((analysis_df["分类"] == "桩材").sum())
    stats["打桩费分析行数"] = int((analysis_df["分类"] == "打桩费").sum())
    stats["截面规格数"] = int(analysis_df["截面规格"].nunique())
    return analysis_df, stats


def summarize_by_spec(df: pd.DataFrame, category: str) -> pd.DataFrame:
    category_df = df[df["分类"] == category].copy()
    if category_df.empty:
        return pd.DataFrame(
            columns=[
                "规格",
                "分类",
                "样本数",
                "项目数",
                "供应商数",
                "总根数",
                "总金额",
                "加权平均单价",
                "最低单价",
                "最高单价",
                "加权平均每米成本",
                "最低每米成本",
                "最高每米成本",
                "宽mm",
                "高mm",
                "长度m",
            ]
        )

    summary_rows = []
    for spec, group in category_df.groupby("规格", sort=False):
        total_quantity = float(group["数量"].sum())
        weighted_unit_price = weighted_average(group, "单价", "数量")
        weighted_per_meter = weighted_average(group, "每米成本", "数量")
        summary_rows.append(
            {
                "规格": spec,
                "分类": category,
                "样本数": int(len(group)),
                "项目数": int(group["标准项目名称"].nunique()),
                "供应商数": int(group["供货商名称"].fillna("未标注").nunique()),
                "总根数": total_quantity,
                "总金额": float(group["不含税金额"].sum()),
                "加权平均单价": weighted_unit_price,
                "最低单价": float(group["单价"].min()),
                "最高单价": float(group["单价"].max()),
                "加权平均每米成本": weighted_per_meter,
                "最低每米成本": float(group["每米成本"].min()),
                "最高每米成本": float(group["每米成本"].max()),
                "宽mm": float(group["宽mm"].iloc[0]),
                "高mm": float(group["高mm"].iloc[0]),
                "长度m": float(group["长度m"].iloc[0]),
            }
        )

    summary_df = pd.DataFrame(summary_rows)
    summary_df = summary_df.sort_values(by=["宽mm", "高mm", "长度m", "规格"]).reset_index(drop=True)
    return summary_df


def summarize_by_section(df: pd.DataFrame, category: str) -> pd.DataFrame:
    category_df = df[df["分类"] == category].copy()
    if category_df.empty:
        return pd.DataFrame(
            columns=[
                "截面规格",
                "分类",
                "样本数",
                "规格数",
                "项目数",
                "供应商数",
                "总根数",
                "总米数",
                "总金额",
                "加权平均长度m",
                "最短长度m",
                "最长长度m",
                "加权平均单价",
                "最低单价",
                "最高单价",
                "汇总每米成本",
                "宽mm",
                "高mm",
            ]
        )

    summary_rows = []
    for section, group in category_df.groupby("截面规格", sort=False):
        total_quantity = float(group["数量"].sum())
        total_meter = float(group["总米数"].sum())
        total_amount = float(group["不含税金额"].sum())
        merged_per_meter = (total_amount / total_meter) if total_meter > 0 else None
        summary_rows.append(
            {
                "截面规格": section,
                "分类": category,
                "样本数": int(len(group)),
                "规格数": int(group["规格"].nunique()),
                "项目数": int(group["标准项目名称"].nunique()),
                "供应商数": int(group["供货商名称"].fillna("未标注").nunique()),
                "总根数": total_quantity,
                "总米数": total_meter,
                "总金额": total_amount,
                "加权平均长度m": weighted_average(group, "长度m", "数量"),
                "最短长度m": float(group["长度m"].min()),
                "最长长度m": float(group["长度m"].max()),
                "加权平均单价": weighted_average(group, "单价", "数量"),
                "最低单价": float(group["单价"].min()),
                "最高单价": float(group["单价"].max()),
                "汇总每米成本": merged_per_meter,
                "宽mm": float(group["宽mm"].iloc[0]),
                "高mm": float(group["高mm"].iloc[0]),
            }
        )

    summary_df = pd.DataFrame(summary_rows)
    summary_df = summary_df.sort_values(by=["宽mm", "高mm", "截面规格"]).reset_index(drop=True)
    return summary_df


def build_overview_cards(analysis_df: pd.DataFrame, pile_summary: pd.DataFrame, driving_summary: pd.DataFrame) -> list[dict[str, str]]:
    pile_df = analysis_df[analysis_df["分类"] == "桩材"].copy()
    driving_df = analysis_df[analysis_df["分类"] == "打桩费"].copy()

    common_specs = set(pile_summary["规格"]).intersection(set(driving_summary["规格"]))
    combined_summary = pile_summary[["规格", "加权平均每米成本"]].merge(
        driving_summary[["规格", "加权平均每米成本"]],
        on="规格",
        how="inner",
        suffixes=("_桩材", "_打桩费"),
    )
    if not combined_summary.empty:
        combined_summary["综合每米成本"] = (
            combined_summary["加权平均每米成本_桩材"] + combined_summary["加权平均每米成本_打桩费"]
        )
        top_combined = combined_summary.sort_values("综合每米成本", ascending=False).iloc[0]
        top_combined_value = format_number(top_combined["综合每米成本"])
        top_combined_text = f"{top_combined['规格']}，可对比规格 {format_number(len(common_specs), 0)} 种"
    else:
        top_combined_value = "-"
        top_combined_text = "-"

    cards = [
        {
            "label": "分析样本",
            "value": format_number(len(analysis_df), 0),
            "hint": f"覆盖项目 {analysis_df['标准项目名称'].nunique()} 个，规格 {analysis_df['规格'].nunique()} 种",
            "tone": "slate",
        },
        {
            "label": "桩材加权均价",
            "value": format_number(weighted_average(pile_df, "单价", "数量")),
            "hint": f"总根数 {format_number(pile_df['数量'].sum(), 0)}，每米均价 {format_number(weighted_average(pile_df, '每米成本', '数量'))}",
            "tone": "amber",
        },
        {
            "label": "打桩费加权均价",
            "value": format_number(weighted_average(driving_df, "单价", "数量")),
            "hint": f"总根数 {format_number(driving_df['数量'].sum(), 0)}，每米均价 {format_number(weighted_average(driving_df, '每米成本', '数量'))}",
            "tone": "teal",
        },
        {
          "label": "综合每米最高价",
          "value": top_combined_value,
          "hint": top_combined_text,
            "tone": "violet",
        },
    ]
    return cards


def build_price_chart(summary_df: pd.DataFrame, title: str, avg_color: str, div_id: str) -> str:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=summary_df["规格"],
            y=summary_df["最低单价"],
            name="最低价",
            marker_color=COLOR_MIN,
            hovertemplate="规格=%{x}<br>最低价=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=summary_df["规格"],
            y=summary_df["加权平均单价"],
            name="平均价",
            marker_color=avg_color,
            hovertemplate="规格=%{x}<br>平均价=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=summary_df["规格"],
            y=summary_df["最高单价"],
            name="最高价",
            marker_color=COLOR_MAX,
            hovertemplate="规格=%{x}<br>最高价=%{y:,.2f}<extra></extra>",
        )
    )

    fig.update_layout(
        title=title,
        barmode="group",
        height=max(420, len(summary_df) * 22),
        margin=dict(l=40, r=20, t=70, b=120),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(title="规格", tickangle=-40),
        yaxis=dict(title="单价"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white",
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)


def build_per_meter_chart(pile_summary: pd.DataFrame, driving_summary: pd.DataFrame, div_id: str) -> str:
    merged = pile_summary[["规格", "加权平均每米成本"]].merge(
        driving_summary[["规格", "加权平均每米成本"]],
        on="规格",
        how="outer",
        suffixes=("_桩材", "_打桩费"),
    )
    merged = merged.sort_values(by="规格", key=lambda series: series.map(spec_sort_key)).reset_index(drop=True)
    merged["综合每米成本"] = merged["加权平均每米成本_桩材"].fillna(0) + merged["加权平均每米成本_打桩费"].fillna(0)
    merged.loc[
        merged[["加权平均每米成本_桩材", "加权平均每米成本_打桩费"]].isna().any(axis=1),
        "综合每米成本",
    ] = None

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=merged["规格"],
            y=merged["加权平均每米成本_桩材"],
            name="桩材每米成本",
            marker_color=COLOR_PILE,
            hovertemplate="规格=%{x}<br>桩材每米成本=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=merged["规格"],
            y=merged["加权平均每米成本_打桩费"],
            name="打桩费每米成本",
            marker_color=COLOR_DRIVING,
            hovertemplate="规格=%{x}<br>打桩费每米成本=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=merged["规格"],
            y=merged["综合每米成本"],
            name="综合每米成本",
            marker_color=COLOR_TOTAL,
            hovertemplate="规格=%{x}<br>综合每米成本=%{y:,.2f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="不同规格的每米成本对比",
        barmode="group",
        height=max(420, len(merged) * 22),
        margin=dict(l=40, r=20, t=70, b=120),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(title="规格", tickangle=-40),
        yaxis=dict(title="每米成本"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white",
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)


def build_section_per_meter_chart(pile_section: pd.DataFrame, driving_section: pd.DataFrame, div_id: str) -> str:
    merged = pile_section[["截面规格", "汇总每米成本", "加权平均长度m"]].merge(
        driving_section[["截面规格", "汇总每米成本", "加权平均长度m"]],
        on="截面规格",
        how="outer",
        suffixes=("_桩材", "_打桩费"),
    )
    merged = merged.sort_values(by="截面规格", key=lambda series: series.map(section_sort_key)).reset_index(drop=True)
    merged["综合每米成本"] = merged["汇总每米成本_桩材"].fillna(0) + merged["汇总每米成本_打桩费"].fillna(0)
    merged.loc[
        merged[["汇总每米成本_桩材", "汇总每米成本_打桩费"]].isna().all(axis=1),
        "综合每米成本",
    ] = None

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=merged["截面规格"],
            y=merged["汇总每米成本_桩材"],
            name="桩材每米成本",
            marker_color=COLOR_PILE,
            hovertemplate="截面=%{x}<br>桩材每米成本=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=merged["截面规格"],
            y=merged["汇总每米成本_打桩费"],
            name="打桩费每米成本",
            marker_color=COLOR_DRIVING,
            hovertemplate="截面=%{x}<br>打桩费每米成本=%{y:,.2f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=merged["截面规格"],
            y=merged["综合每米成本"],
            name="综合每米成本",
            marker_color=COLOR_TOTAL,
            hovertemplate="截面=%{x}<br>综合每米成本=%{y:,.2f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="按截面规格汇总的每米成本（例：400x400）",
        barmode="group",
        height=max(420, len(merged) * 22),
        margin=dict(l=40, r=20, t=70, b=120),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(title="截面规格", tickangle=-35),
        yaxis=dict(title="每米成本"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white",
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)


def format_summary_table(summary_df: pd.DataFrame, title: str) -> str:
  display_df = summary_df.copy()
  for column in ["总根数", "加权平均单价", "最低单价", "最高单价", "加权平均每米成本", "最低每米成本", "最高每米成本"]:
    if column in display_df.columns:
      digits = 0 if column == "总根数" else 2
      display_df[column] = display_df[column].apply(lambda value: format_number(value, digits))

  if "长度m" in display_df.columns:
    display_df["长度m"] = display_df["长度m"].apply(lambda value: format_number(value, 1))
  if "宽mm" in display_df.columns:
    display_df["宽mm"] = display_df["宽mm"].apply(lambda value: format_number(value, 0))
  if "高mm" in display_df.columns:
    display_df["高mm"] = display_df["高mm"].apply(lambda value: format_number(value, 0))

  display_df = display_df[
    [
      "规格",
      "样本数",
      "项目数",
      "供应商数",
      "总根数",
      "加权平均单价",
      "最低单价",
      "最高单价",
      "加权平均每米成本",
      "最低每米成本",
      "最高每米成本",
    ]
  ]

  return f"""
  <section class=\"table-card\">
    <div class=\"section-head\">
    <h3>{title}</h3>
    <p>按规格聚合，均价采用数量加权。</p>
    </div>
    <div class=\"table-wrap\">{display_df.to_html(index=False, classes='data-table', border=0)}</div>
  </section>
  """


def format_section_summary_table(summary_df: pd.DataFrame, title: str) -> str:
  display_df = summary_df.copy()
  number_columns_digits = {
    "总根数": 0,
    "总米数": 1,
    "加权平均长度m": 2,
    "最短长度m": 1,
    "最长长度m": 1,
    "加权平均单价": 2,
    "最低单价": 2,
    "最高单价": 2,
    "汇总每米成本": 2,
  }
  for column, digits in number_columns_digits.items():
    if column in display_df.columns:
      display_df[column] = display_df[column].apply(lambda value: format_number(value, digits))

  display_df = display_df[
    [
      "截面规格",
      "样本数",
      "规格数",
      "项目数",
      "供应商数",
      "总根数",
      "加权平均长度m",
      "最短长度m",
      "最长长度m",
      "汇总每米成本",
      "加权平均单价",
      "最低单价",
      "最高单价",
      "总米数",
    ]
  ]

  return f"""
  <section class=\"table-card\">
    <div class=\"section-head\">
    <h3>{title}</h3>
    <p>把 400x400x9m、400x400x10m 等规格按截面 400x400 汇总，长度按数量加权平均。</p>
    </div>
    <div class=\"table-wrap\">{display_df.to_html(index=False, classes='data-table', border=0)}</div>
  </section>
  """


def build_dashboard_html(
    source_file: Path,
    sheet_name: str,
    analysis_df: pd.DataFrame,
    stats: dict[str, int],
    pile_summary: pd.DataFrame,
    driving_summary: pd.DataFrame,
    pile_section_summary: pd.DataFrame,
    driving_section_summary: pd.DataFrame,
    output_file: Path,
) -> str:
    cards = build_overview_cards(analysis_df, pile_summary, driving_summary)
    card_html = "".join(
        f"""
        <div class=\"metric-card tone-{card['tone']}\">
          <div class=\"metric-label\">{card['label']}</div>
          <div class=\"metric-value\">{card['value']}</div>
          <div class=\"metric-hint\">{card['hint']}</div>
        </div>
        """
        for card in cards
    )

    pile_chart_html = build_price_chart(pile_summary, "桩材价格分布：平均价 / 最高价 / 最低价", COLOR_PILE, "pile-price-chart")
    driving_chart_html = build_price_chart(driving_summary, "打桩费价格分布：平均价 / 最高价 / 最低价", COLOR_DRIVING, "driving-price-chart")
    per_meter_chart_html = build_per_meter_chart(pile_summary, driving_summary, "per-meter-chart")
    section_per_meter_chart_html = build_section_per_meter_chart(
        pile_section_summary,
        driving_section_summary,
        "section-per-meter-chart",
    )

    note_items = [
        f"原始桩相关记录 {stats['桩相关总行数']} 行",
        f"标准单位=根 {stats['标准单位为根']} 行",
        f"规格可解析 {stats['规格可解析']} 行",
        f"最终纳入分析 {stats['最终分析行数']} 行",
        f"截面规格共 {stats['截面规格数']} 种（如 400x400）",
        "桩材类剔除了运费、引孔、设备、进场等服务性条目",
        "单价均采用原表单价列，均价按数量加权计算",
    ]
    note_html = "".join(f"<li>{item}</li>" for item in note_items)

    pile_table_html = format_summary_table(pile_summary, "桩材规格明细")
    driving_table_html = format_summary_table(driving_summary, "打桩费规格明细")
    pile_section_table_html = format_section_summary_table(pile_section_summary, "桩材截面汇总明细")
    driving_section_table_html = format_section_summary_table(driving_section_summary, "打桩费截面汇总明细")

    html = f"""<!DOCTYPE html>
<html lang=\"zh-CN\">
<head>
  <meta charset=\"UTF-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
  <title>桩成本分析看板</title>
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <style>
    :root {{
      --bg: #f4efe6;
      --ink: #1f2937;
      --muted: #667085;
      --card: rgba(255,255,255,0.88);
      --line: rgba(15, 23, 42, 0.08);
      --shadow: 0 24px 80px rgba(15, 23, 42, 0.12);
      --amber: #d97706;
      --teal: #0f766e;
      --violet: #7c3aed;
      --slate: #334155;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Avenir Next", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(217, 119, 6, 0.18), transparent 26%),
        radial-gradient(circle at top right, rgba(15, 118, 110, 0.14), transparent 24%),
        linear-gradient(180deg, #fbf7f1 0%, var(--bg) 100%);
    }}
    .page {{
      max-width: 1480px;
      margin: 0 auto;
      padding: 36px 24px 48px;
    }}
    .hero {{
      position: relative;
      overflow: hidden;
      padding: 28px 30px;
      border-radius: 28px;
      background: linear-gradient(135deg, rgba(255,255,255,0.96), rgba(252, 244, 229, 0.88));
      border: 1px solid rgba(217, 119, 6, 0.12);
      box-shadow: var(--shadow);
    }}
    .hero::after {{
      content: "";
      position: absolute;
      right: -80px;
      top: -70px;
      width: 260px;
      height: 260px;
      background: radial-gradient(circle, rgba(15, 118, 110, 0.16), transparent 62%);
      pointer-events: none;
    }}
    .eyebrow {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 12px;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--amber);
      font-weight: 700;
    }}
    h1 {{
      margin: 12px 0 10px;
      font-size: clamp(28px, 4vw, 44px);
      line-height: 1.1;
    }}
    .hero p {{
      max-width: 980px;
      margin: 0;
      font-size: 15px;
      line-height: 1.8;
      color: var(--muted);
    }}
    .metric-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 16px;
      margin: 24px 0 10px;
    }}
    .metric-card {{
      position: relative;
      overflow: hidden;
      padding: 20px 20px 18px;
      border-radius: 22px;
      background: var(--card);
      border: 1px solid var(--line);
      box-shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
    }}
    .metric-card::before {{
      content: "";
      position: absolute;
      inset: 0 auto auto 0;
      width: 100%;
      height: 4px;
      opacity: 0.95;
    }}
    .tone-amber::before {{ background: linear-gradient(90deg, #f59e0b, #d97706); }}
    .tone-teal::before {{ background: linear-gradient(90deg, #14b8a6, #0f766e); }}
    .tone-violet::before {{ background: linear-gradient(90deg, #8b5cf6, #7c3aed); }}
    .tone-slate::before {{ background: linear-gradient(90deg, #64748b, #334155); }}
    .metric-label {{
      font-size: 13px;
      color: var(--muted);
    }}
    .metric-value {{
      margin: 10px 0 8px;
      font-size: 34px;
      font-weight: 800;
      letter-spacing: -0.03em;
    }}
    .metric-hint {{
      font-size: 13px;
      line-height: 1.6;
      color: var(--muted);
    }}
    .chart-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
      margin-top: 18px;
    }}
    .chart-card, .table-card, .note-card {{
      background: rgba(255,255,255,0.92);
      border-radius: 24px;
      border: 1px solid var(--line);
      box-shadow: 0 18px 48px rgba(15, 23, 42, 0.08);
    }}
    .chart-card {{ padding: 18px 18px 10px; }}
    .chart-card.full {{ grid-column: 1 / -1; }}
    .section-head {{
      padding: 18px 20px 0;
    }}
    .section-head h2, .section-head h3 {{
      margin: 0;
      font-size: 22px;
    }}
    .section-head p {{
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.7;
    }}
    .note-card {{
      padding: 12px 22px 18px;
      margin-top: 18px;
    }}
    .note-card ul {{
      margin: 14px 0 0;
      padding-left: 18px;
      color: var(--muted);
      line-height: 1.9;
    }}
    .table-card {{
      margin-top: 18px;
      overflow: hidden;
    }}
    .table-wrap {{
      padding: 12px 18px 20px;
      overflow: auto;
    }}
    table.data-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      white-space: nowrap;
    }}
    table.data-table thead th {{
      position: sticky;
      top: 0;
      background: #f8fafc;
      z-index: 1;
      color: #0f172a;
      font-weight: 700;
    }}
    table.data-table th,
    table.data-table td {{
      padding: 10px 12px;
      border-bottom: 1px solid rgba(148, 163, 184, 0.16);
      text-align: right;
    }}
    table.data-table th:first-child,
    table.data-table td:first-child {{
      text-align: left;
    }}
    table.data-table tbody tr:hover {{
      background: rgba(248, 250, 252, 0.84);
    }}
    @media (max-width: 980px) {{
      .chart-grid {{ grid-template-columns: 1fr; }}
      .chart-card.full {{ grid-column: auto; }}
      .page {{ padding: 20px 14px 34px; }}
      .hero {{ padding: 22px 18px; }}
    }}
  </style>
</head>
<body>
  <div class=\"page\">
    <section class=\"hero\">
      <div class=\"eyebrow\">Pile Cost Dashboard</div>
      <h1>桩材与打桩费成本分析看板</h1>
      <p>围绕 物品名称、规格、标准单位=根、单价 四个核心字段，拆分出 桩材价格 和 打桩费价格 两类分析，并进一步计算不同规格的每米成本。图表中的平均价均采用数量加权，适合直接用于快速查看规格差异和异常价格。</p>
    </section>

    <section class=\"metric-grid\">{card_html}</section>

    <section class=\"note-card\">
      <div class=\"section-head\">
        <h3>分析说明</h3>
        <p>先对桩相关记录做口径清洗，再分别输出桩材价格与打桩费价格的统计结果。</p>
      </div>
      <ul>{note_html}</ul>
    </section>

    <section class=\"chart-grid\">
      <section class="chart-card full">
        <div class="section-head">
          <h2>截面汇总每米成本</h2>
          <p>把不同长度先合并到同一截面（如 400x400），并输出该截面的平均长度与每米成本汇总。</p>
        </div>
        {section_per_meter_chart_html}
      </section>
    </section>

    {pile_section_table_html}
    {driving_section_table_html}

    <section class="chart-grid">
      <section class="chart-card full">
        <div class="section-head">
          <h2>每米成本对比</h2>
          <p>同一规格下，分别展示桩材每米成本、打桩费每米成本，以及可同时匹配时的综合每米成本。</p>
        </div>
        {per_meter_chart_html}
      </section>
      <section class=\"chart-card\">
        <div class=\"section-head\">
          <h2>桩材价格分析</h2>
          <p>对比各规格桩材的平均价、最高价、最低价。</p>
        </div>
        {pile_chart_html}
      </section>
      <section class=\"chart-card\">
        <div class=\"section-head\">
          <h2>打桩费价格分析</h2>
          <p>对比各规格打桩费的平均价、最高价、最低价。</p>
        </div>
        {driving_chart_html}
      </section>
    </section>

    {pile_table_html}
    {driving_table_html}
  </div>
</body>
</html>
"""
    return html


def main():
    args = parse_args()
    if args.input_file:
        input_file = Path(args.input_file).expanduser().resolve()
    else:
        latest = find_latest_summary_file()
        if latest is None:
            raise FileNotFoundError(f"未在 {SUMMARY_INPUT_ROOTS} 找到 {INPUT_GLOB}")
        input_file = latest.resolve()

    output_file, excel_out = build_output_path(input_file, args.output_file)

    if not input_file.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    print(f"读取文件: {input_file}")
    source_df = load_source_data(input_file, args.sheet_name)
    analysis_df, stats = prepare_analysis_data(source_df)

    if analysis_df.empty:
        raise ValueError("清洗后没有可用于分析的桩数据，请检查规格、标准单位或单价字段")

    pile_summary = summarize_by_spec(analysis_df, "桩材")
    driving_summary = summarize_by_spec(analysis_df, "打桩费")
    pile_section_summary = summarize_by_section(analysis_df, "桩材")
    driving_section_summary = summarize_by_section(analysis_df, "打桩费")

    html = build_dashboard_html(
        input_file,
        args.sheet_name,
        analysis_df,
        stats,
        pile_summary,
        driving_summary,
        pile_section_summary,
        driving_section_summary,
        output_file,
    )
    output_file.write_text(html, encoding="utf-8")

    # --- Excel 输出 ---
    with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
        analysis_df.to_excel(writer, sheet_name="明细数据", index=False)
        pile_summary.to_excel(writer, sheet_name="桩材规格汇总", index=False)
        driving_summary.to_excel(writer, sheet_name="打桩费规格汇总", index=False)
        pile_section_summary.to_excel(writer, sheet_name="桩材截面汇总", index=False)
        driving_section_summary.to_excel(writer, sheet_name="打桩费截面汇总", index=False)
        wb = writer.book
        for sname in ["明细数据", "桩材规格汇总", "打桩费规格汇总", "桩材截面汇总", "打桩费截面汇总"]:
            apply_common_excel_layout(wb[sname], numeric_start_col=1)

    print(f"分析完成，HTML 已输出: {output_file}")
    print(f"Excel 已输出: {excel_out}")
    print(f"纳入分析记录: {stats['最终分析行数']} 行")
    print(f"桩材记录: {stats['桩材分析行数']} 行，规格数: {pile_summary['规格'].nunique()}")
    print(f"打桩费记录: {stats['打桩费分析行数']} 行，规格数: {driving_summary['规格'].nunique()}")
    print(f"截面规格数: {stats['截面规格数']}（桩材 {pile_section_summary['截面规格'].nunique()}，打桩费 {driving_section_summary['截面规格'].nunique()}）")


if __name__ == "__main__":
    main()