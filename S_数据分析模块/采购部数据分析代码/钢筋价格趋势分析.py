from __future__ import annotations

import argparse
from datetime import date, timedelta
from html import escape
import json
import math
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote
from urllib.request import Request, urlopen

import numpy as np
import pandas as pd
import plotly.graph_objects as go


SCRIPT_DIR = Path(__file__).resolve().parent
# 脚本位于 19_代码/S_数据分析模块/采购部数据分析代码/，需上三级到达 01成本统计。
# 注意 SCRIPT_DIR 是目录路径，parents[2] 才是 01成本统计。
_PROJECT_ROOT = SCRIPT_DIR.parents[2]
_DATA_ROOT = _PROJECT_ROOT / "9_采购部采购台账"
SUMMARY_INPUT_ROOT = _DATA_ROOT / "B_汇总数据" / "B_标准化汇总"
ANALYSIS_OUTPUT_ROOT = _DATA_ROOT / "C_数据分析"
INPUT_GLOB = "*_材料采购_merged.xlsx"
DEFAULT_SHEET_NAME = "合并数据"
OUTPUT_EXCEL_SUFFIX = "_钢筋价格趋势分析"
OUTPUT_HTML_SUFFIX = "_钢筋价格趋势分析"

STEEL_TOOL_EXCLUDE = (
    "运费",
    "搬运",
    "人工",
    "施工",
    "加工",
    "切割",
    "焊接",
    "扳手",
    "扎钩",
    "钢筋棚",
    "钢筋剪",
    "钢筋钩",
)

DIA_PATTERN = re.compile(r"[Φφ]\s*(\d+(?:\.\d+)?)")
MESH_CANON_PATTERN = re.compile(
    r"[Φφ]?(\d+(?:\.\d+)?)\s*@\s*(\d{2,3})\s*[xX]\s*(\d{2,3})(?:\s*-\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*m)?",
    re.IGNORECASE,
)
MESH_RAW_PATTERN = re.compile(
    r"(\d+(?:\.\d+)?)\s*(?:mm|毫米)?\s*[xX×*＊]\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米|mm|毫米)?\s*[xX×*＊]\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米|mm|毫米)?\s*[xX×*＊]\s*(\d+(?:\.\d+)?)\s*[xX×*＊]\s*(\d+(?:\.\d+)?)\s*(?:m|米)?",
    re.IGNORECASE,
)


def normalize_col_name(name: str) -> str:
    s = str(name).strip().lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s_\-]+", "", s)
    return s


def is_blank(val: object) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    return str(val).strip().lower() in {"", "nan", "none", "null"}


def find_column(columns: list[str], aliases: list[str]) -> str | None:
    norm_map = {normalize_col_name(c): c for c in columns}
    for alias in aliases:
        key = normalize_col_name(alias)
        if key in norm_map:
            return norm_map[key]

    for col in columns:
        col_key = normalize_col_name(col)
        for alias in aliases:
            alias_key = normalize_col_name(alias)
            if alias_key and alias_key in col_key:
                return col
    return None


def normalize_unit(val: object) -> str:
    if is_blank(val):
        return ""
    s = str(val).strip().lower()
    s = s.replace(" ", "")
    return s


def is_kg_unit(val: object) -> bool:
    u = normalize_unit(val)
    return u in {"kg", "kgs", "公斤", "千克", "kg.", "kg/"}


def month_text(year_val: object, month_val: object) -> str | None:
    try:
        y = int(float(str(year_val).strip()))
        m = int(float(str(month_val).strip()))
        if y < 2000 or y > 2100 or m < 1 or m > 12:
            return None
        return f"{y}-{m:02d}"
    except Exception:
        return None


def parse_year_month_text(val: object) -> str | None:
    if is_blank(val):
        return None
    txt = str(val).strip()

    # 兼容格式：2025-11 / 2025/11 / 2025年11月 / 202511
    m = re.search(r"((?:19|20)\d{2})\D{0,2}(0?[1-9]|1[0-2])", txt)
    if m:
        y = int(m.group(1))
        mo = int(m.group(2))
        return f"{y}-{mo:02d}"

    m2 = re.fullmatch(r"((?:19|20)\d{2})(0[1-9]|1[0-2])", re.sub(r"\D", "", txt))
    if m2:
        y = int(m2.group(1))
        mo = int(m2.group(2))
        return f"{y}-{mo:02d}"

    return None


def detect_material(item_name: object) -> str | None:
    if is_blank(item_name):
        return None
    txt = str(item_name).strip()
    if not txt:
        return None

    if any(k in txt for k in STEEL_TOOL_EXCLUDE):
        return None

    if "网片" in txt:
        return "钢筋网片"
    if "螺纹钢" in txt:
        return "螺纹钢"
    if "钢筋" in txt:
        return "钢筋"
    return None


def detect_material_from_class(material_class: object, item_name: object) -> str | None:
    """融合“材料归类 + 物品名称”，优先采用更具体的类型（网片/螺纹钢优先于钢筋）。"""
    item_type = detect_material(item_name)

    class_type = None
    if not is_blank(material_class):
        cls = str(material_class).strip()
        if "网片" in cls:
            class_type = "钢筋网片"
        elif "螺纹钢" in cls:
            class_type = "螺纹钢"
        elif "钢筋" in cls or "钢材" in cls:
            class_type = "钢筋"

    # 具体程度排序：钢筋网片 > 螺纹钢 > 钢筋
    rank = {"钢筋": 1, "螺纹钢": 2, "钢筋网片": 3}
    candidates = [t for t in [item_type, class_type] if t is not None]
    if not candidates:
        return None

    return max(candidates, key=lambda t: rank.get(t, 0))


def fmt_num(v: float) -> str:
    r = round(v)
    if math.isclose(v, r, abs_tol=1e-9):
        return str(int(r))
    return f"{v:.6f}".rstrip("0").rstrip(".")


def canonical_diameter(spec: object) -> str | None:
    if is_blank(spec):
        return None
    txt = str(spec)
    m = DIA_PATTERN.search(txt)
    if not m:
        return None
    d = float(m.group(1))
    if d <= 0 or d > 100:
        return None
    return f"Φ{fmt_num(d)}"


def canonical_mesh_spec(spec: object) -> str | None:
    if is_blank(spec):
        return None
    txt = str(spec).strip().replace("×", "x").replace("＊", "x").replace("*", "x")

    m = MESH_CANON_PATTERN.search(txt)
    if m:
        dia = float(m.group(1))
        s1 = int(m.group(2))
        s2 = int(m.group(3))
        if m.group(4) and m.group(5):
            w = float(m.group(4))
            h = float(m.group(5))
            return f"Φ{fmt_num(dia)}@{s1}x{s2}-{fmt_num(w)}x{fmt_num(h)}m"
        return f"Φ{fmt_num(dia)}@{s1}x{s2}"

    m2 = MESH_RAW_PATTERN.search(txt)
    if m2:
        dia = float(m2.group(1))
        p1 = float(m2.group(2))
        p2 = float(m2.group(3))
        w = float(m2.group(4))
        h = float(m2.group(5))

        # 20cm -> 200mm；若已是毫米数量级（>=50）则按毫米直接使用
        p1_mm = int(round(p1 * 10 if p1 < 50 else p1))
        p2_mm = int(round(p2 * 10 if p2 < 50 else p2))
        return f"Φ{fmt_num(dia)}@{p1_mm}x{p2_mm}-{fmt_num(w)}x{fmt_num(h)}m"

    # 至少保留线径，不强制必须有片尺寸
    dia = canonical_diameter(txt)
    return dia


def outlier_filter_series(
    series: pd.Series,
    mag_factor: float = 10.0,
    iqr_k: float = 2.5,
    iqr_min_n: int = 10,
) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    valid = s.notna() & (s > 0)
    keep = valid.copy()
    if valid.sum() == 0:
        return keep

    vals = s[valid]
    n = int(vals.shape[0])
    median = float(vals.median())

    # 数量级过滤（放宽）：避免轻微偏离即被剔除
    mag_low = median / mag_factor if median > 0 else 0
    mag_high = median * mag_factor if median > 0 else np.inf
    keep &= (s >= mag_low) & (s <= mag_high)

    # IQR过滤（放宽）：样本更充足时才启用，边界由1.5倍放宽到2.0倍
    if n >= iqr_min_n:
        q1 = float(vals.quantile(0.25))
        q3 = float(vals.quantile(0.75))
        iqr = q3 - q1
        low = q1 - iqr_k * iqr
        high = q3 + iqr_k * iqr
        keep &= (s >= low) & (s <= high)

    return keep.fillna(False)


def build_output_paths(input_file: Path, excel_suffix: str, html_suffix: str) -> tuple[Path, Path]:
    excel_dir = ANALYSIS_OUTPUT_ROOT / "Excel"
    html_dir = ANALYSIS_OUTPUT_ROOT / "HTML"
    excel_dir.mkdir(parents=True, exist_ok=True)
    html_dir.mkdir(parents=True, exist_ok=True)
    excel_path = excel_dir / f"{input_file.stem}{excel_suffix}{input_file.suffix}"
    html_path = html_dir / f"{input_file.stem}{html_suffix}.html"
    return excel_path, html_path


def extract_year_month_from_name(path: Path) -> tuple[int, int] | None:
    m = re.search(r"(?<!\d)((?:19|20)\d{2})年(0?[1-9]|1[0-2])月", path.stem)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def find_latest_summary_file() -> Path | None:
    if not SUMMARY_INPUT_ROOT.is_dir():
        return None
    candidates = [p for p in SUMMARY_INPUT_ROOT.glob(INPUT_GLOB) if p.is_file()]
    if not candidates:
        return None

    def sort_key(p: Path):
        ym = extract_year_month_from_name(p)
        if ym is not None:
            return (1, ym[0], ym[1], p.stat().st_mtime)
        return (0, 0, 0, p.stat().st_mtime)

    return max(candidates, key=sort_key)


def parse_args():
    parser = argparse.ArgumentParser(description="按规格分析钢筋/螺纹钢/钢筋网片月度价格趋势（kg口径）")
    parser.add_argument(
        "input_file",
        nargs="?",
        default="",
        help="输入 Excel 文件；不传时自动读取 C_数据分析 下最新 *_材料采购_merged_标准化汇总.xlsx",
    )
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help="工作表名，默认 合并数据")
    parser.add_argument("--excel-suffix", default=OUTPUT_EXCEL_SUFFIX, help="输出 Excel 后缀")
    parser.add_argument("--html-suffix", default=OUTPUT_HTML_SUFFIX, help="输出 HTML 后缀")
    parser.add_argument("--outlier-mag-factor", type=float, default=10.0, help="数量级过滤倍数，默认10")
    parser.add_argument("--outlier-iqr-k", type=float, default=2.5, help="IQR倍数，默认2.5")
    parser.add_argument("--outlier-iqr-min-n", type=int, default=10, help="启用IQR所需最小样本数，默认10")
    return parser.parse_args()


def fmt_price(v: Any) -> str:
    try:
        n = float(v)
        return f"{n:,.2f}"
    except Exception:
        return "-"


def build_latest_price_table(summary: pd.DataFrame) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    if summary.empty:
        return pd.DataFrame(
            columns=["材料类型", "规格组", "最新月份", "最新均价", "上期均价", "环比变化", "环比变化%", "最新样本数"]
        )

    for (mat, spec), part in summary.groupby(["材料类型", "规格组"], sort=False):
        d = part.sort_values("年-月")
        last = d.iloc[-1]
        prev = d.iloc[-2] if len(d) >= 2 else None

        last_avg = float(last["平均单价"])
        prev_avg = float(prev["平均单价"]) if prev is not None else np.nan
        diff = last_avg - prev_avg if prev is not None else np.nan
        pct = (diff / prev_avg * 100) if prev is not None and prev_avg != 0 else np.nan

        records.append(
            {
                "材料类型": mat,
                "规格组": spec,
                "最新月份": str(last["年-月"]),
                "最新均价": last_avg,
                "上期均价": prev_avg,
                "环比变化": diff,
                "环比变化%": pct,
                "最新样本数": int(last["样本数"]),
            }
        )

    out = pd.DataFrame(records)
    if out.empty:
        return out
    return out.sort_values(["材料类型", "最新月份", "最新均价"], ascending=[True, False, False]).reset_index(drop=True)


# 东方财富secid映射：SHFE前缀113，DCE前缀114
_EASTMONEY_SECIDS: dict[str, str] = {
    "螺纹钢主力": "113.RBM",
    "热卷主力":   "113.HCM",
    "铁矿石主力": "114.IM",
}


def fetch_online_steel_history(
    months_back: int = 24,
    timeout_sec: int = 12,
) -> tuple[dict[str, pd.DataFrame], str | None]:
    """从东方财富抓取主力期货历史日K数据，返回 {品种名: DataFrame(date, close)}。"""
    end_date = date.today().strftime("%Y%m%d")
    start_date = (date.today() - timedelta(days=months_back * 31)).strftime("%Y%m%d")

    results: dict[str, pd.DataFrame] = {}
    errors: list[str] = []

    for name, secid in _EASTMONEY_SECIDS.items():
        url = (
            f"https://push2his.eastmoney.com/api/qt/stock/kline/get"
            f"?secid={secid}"
            f"&fields1=f1,f2,f3,f4,f5,f6"
            f"&fields2=f51,f52,f53,f54,f55,f56,f57,f58"
            f"&klt=101&fqt=1&beg={start_date}&end={end_date}&lmt=600"
        )
        req = Request(
            url,
            headers={
                "Referer": "https://quote.eastmoney.com",
                "User-Agent": "Mozilla/5.0",
            },
        )
        try:
            with urlopen(req, timeout=timeout_sec) as resp:
                data = json.loads(resp.read().decode("utf-8", errors="ignore"))
            klines = (data.get("data") or {}).get("klines") or []
            rows: list[dict[str, Any]] = []
            for k in klines:
                parts = k.split(",")
                if len(parts) >= 3:
                    try:
                        rows.append({"date": parts[0], "close": float(parts[2])})
                    except ValueError:
                        pass
            if rows:
                results[name] = pd.DataFrame(rows)
            else:
                errors.append(f"{name}: 返回数据为空")
        except Exception as e:
            errors.append(f"{name}: {e}")

    err_msg = "; ".join(errors) if errors and not results else None
    return results, err_msg


def build_online_history_fig(history_data: dict[str, pd.DataFrame]) -> go.Figure | None:
    """把三个主力合约历史收盘价合并绘制到一张图上。"""
    if not history_data:
        return None

    fig = go.Figure()
    colors = {"螺纹钢主力": "#e05c2d", "热卷主力": "#2d7ae0", "铁矿石主力": "#c0a020"}
    for name, df in history_data.items():
        if df.empty:
            continue
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["close"],
                mode="lines",
                name=name,
                line=dict(width=2, color=colors.get(name)),
                hovertemplate=f"{name}<br>日期: %{{x}}<br>收盘价: %{{y:,.0f}} 元/吨<extra></extra>",
            )
        )

    if not fig.data:
        return None

    fig.update_layout(
        height=380,
        title="钢材期货主力合约历史价格走势（东方财富，元/吨）",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=35, r=20, t=70, b=45),
        hovermode="x unified",
        template="plotly_white",
    )
    fig.update_yaxes(title_text="收盘价 (元/吨)")
    fig.update_xaxes(title_text="日期")
    return fig


def _extract_floats_from_csv_parts(parts: list[str]) -> list[float]:
    nums: list[float] = []
    for p in parts:
        t = p.strip()
        if not t:
            continue
        try:
            nums.append(float(t))
        except Exception:
            continue
    return nums


def fetch_online_steel_snapshot(timeout_sec: int = 8) -> tuple[list[dict[str, Any]], str | None]:
    """抓取公开在线行情快照；失败时返回错误信息，不中断主流程。"""
    symbols = {
        "nf_RB0": "螺纹钢主力",
        "nf_HC0": "热卷主力",
        "nf_I0": "铁矿石主力",
    }
    src = f"https://hq.sinajs.cn/list={','.join(symbols.keys())}"
    url = quote(src, safe=':/?&=,_')
    req = Request(url, headers={"Referer": "https://finance.sina.com.cn", "User-Agent": "Mozilla/5.0"})

    try:
        with urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read()
        text = raw.decode("gbk", errors="ignore")
    except Exception as e:
        return [], f"在线行情获取失败: {e}"

    line_re = re.compile(r'var\s+hq_str_(\w+)="([^"]*)";')
    rows: list[dict[str, Any]] = []
    capture_time = ""
    for code, payload in line_re.findall(text):
        if code not in symbols:
            continue
        parts = [x.strip() for x in payload.split(",")]
        if len(parts) <= 1:
            continue

        name = parts[0] or symbols[code]
        nums = _extract_floats_from_csv_parts(parts[1:12])
        if not nums:
            continue

        latest = nums[0]
        day_high = max(nums[:5]) if len(nums) >= 2 else latest
        day_low = min(nums[:5]) if len(nums) >= 2 else latest
        amplitude = ((day_high - day_low) / latest * 100) if latest else np.nan

        change_pct = np.nan
        if len(nums) >= 2 and nums[1] != 0:
            candidate = (latest - nums[1]) / nums[1] * 100
            if abs(candidate) <= 30:
                change_pct = candidate

        ts_candidates = [p for p in reversed(parts) if re.search(r"\d{2}:\d{2}", p)]
        if ts_candidates and not capture_time:
            capture_time = ts_candidates[0]

        rows.append(
            {
                "品种": name,
                "最新价": latest,
                "日内最高": day_high,
                "日内最低": day_low,
                "日内振幅%": amplitude,
                "涨跌幅%": change_pct,
                "数据源": "新浪财经",
            }
        )

    if not rows:
        return [], "在线行情返回为空或格式变化"

    if capture_time:
        for r in rows:
            r["时间"] = capture_time
    return rows, None


def build_html_dashboard(
    html_out: Path,
    material_fig_html: dict[str, str],
    clean_df: pd.DataFrame,
    outlier_df: pd.DataFrame,
    summary: pd.DataFrame,
    overview: pd.DataFrame,
    online_history_fig_html: str = "",
) -> None:
    latest_df = build_latest_price_table(summary)
    online_rows, online_error = fetch_online_steel_snapshot()

    total_hit = len(clean_df) + len(outlier_df)
    clean_count = len(clean_df)
    outlier_count = len(outlier_df)
    outlier_ratio = (outlier_count / total_hit * 100) if total_hit else 0.0
    spec_count = int(clean_df[["材料类型", "规格组"]].drop_duplicates().shape[0]) if clean_count else 0
    latest_month = "-" if summary.empty else str(summary["年-月"].max())

    history_chart_block = (
        f"<div class='online-history'>{online_history_fig_html}</div>"
        if online_history_fig_html
        else "<p class='empty-tip' style='margin-top:12px'>历史价格曲线获取失败，不影响本地分析。</p>"
    )

    if online_rows:
        chips: list[str] = []
        for row in online_rows:
            chg = row.get("涨跌幅%", np.nan)
            amp = row.get("日内振幅%", np.nan)
            chg_text = "-" if pd.isna(chg) else f"{float(chg):+.2f}%"
            amp_text = "-" if pd.isna(amp) else f"{float(amp):.2f}%"
            tone = "up" if (not pd.isna(chg) and float(chg) > 0) else "down" if (not pd.isna(chg) and float(chg) < 0) else "flat"
            chips.append(
                "<div class='online-chip'>"
                f"<div class='name'>{escape(str(row.get('品种', '-')))}</div>"
                f"<div class='price'>¥{fmt_price(row.get('最新价'))}</div>"
                f"<div class='meta'><span class='{tone}'>涨跌 {escape(chg_text)}</span> | 振幅 {escape(amp_text)}</div>"
                "</div>"
            )
        online_capture_time = str(online_rows[0].get("时间", "")) if online_rows else ""
        online_html = (
            "<section class='online-box'>"
            "<div class='online-header'>"
            "<div class='online-title'>在线钢材行情快照</div>"
            f"<div class='online-sub'>来源：新浪财经 / 东方财富{' | 更新时间：' + escape(online_capture_time) if online_capture_time else ''}</div>"
            "</div>"
            "<div class='online-grid'>"
            f"{''.join(chips)}"
            "</div>"
            f"{history_chart_block}"
            "</section>"
        )
    else:
        online_html = (
            "<section class='online-box'>"
            "<div class='online-header'>"
            "<div class='online-title'>在线钢材行情快照</div>"
            f"<div class='online-sub'>实时快照暂未拉取成功：{escape(online_error or '未知原因')}，不影响本地分析结果。</div>"
            "</div>"
            f"{history_chart_block}"
            "</section>"
        )

    overview_show = overview.copy()
    if not overview_show.empty:
        overview_show["总均价"] = overview_show["总均价"].map(fmt_price)
    overview_table_html = (
        "<p class='empty-tip'>暂无规格总览数据</p>"
        if overview_show.empty
        else overview_show.to_html(index=False, classes="data-table", border=0)
    )

    latest_show = latest_df.copy()
    if not latest_show.empty:
        for c in ["最新均价", "上期均价", "环比变化"]:
            latest_show[c] = latest_show[c].map(fmt_price)
        latest_show["环比变化%"] = latest_show["环比变化%"].map(
            lambda x: "-" if pd.isna(x) else f"{x:+.2f}%"
        )
    latest_table_html = (
        "<p class='empty-tip'>暂无最新价格列表数据</p>"
        if latest_show.empty
        else latest_show.to_html(index=False, classes="data-table", border=0)
    )

    top_list_blocks: list[str] = []
    for mat in ["钢筋", "螺纹钢", "钢筋网片"]:
        part = latest_df[latest_df["材料类型"] == mat].head(6)
        if part.empty:
            items = "<li>暂无数据</li>"
        else:
            lines = []
            for _, row in part.iterrows():
                pct_text = "-"
                if not pd.isna(row["环比变化%"]):
                    pct_text = f"{float(row['环比变化%']):+.2f}%"
                lines.append(
                    "<li>"
                    f"<span class='spec'>{escape(str(row['规格组']))}</span>"
                    f"<span class='price'>¥{fmt_price(row['最新均价'])}</span>"
                    f"<span class='meta'>{escape(str(row['最新月份']))} | 环比 {escape(pct_text)} | 样本 {int(row['最新样本数'])}</span>"
                    "</li>"
                )
            items = "".join(lines)

        block = (
            "<section class='list-card'>"
            f"<h3>{escape(mat)} 最新规格报价</h3>"
            "<ul class='spec-list'>"
            f"{items}"
            "</ul>"
            "</section>"
        )
        top_list_blocks.append(block)

    html = f"""
<!doctype html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>钢筋价格趋势分析看板</title>
    <style>
        :root {{
            --bg: #f4f7f9;
            --panel: #ffffff;
            --ink: #1e2933;
            --muted: #61707f;
            --brand: #0f766e;
            --brand-soft: #d6f0ed;
            --line: #dce5eb;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            font-family: "PingFang SC", "Microsoft YaHei", "Noto Sans SC", sans-serif;
            color: var(--ink);
            background:
                radial-gradient(circle at 0% 0%, #e3f4f1 0%, transparent 35%),
                radial-gradient(circle at 100% 20%, #e8eef8 0%, transparent 30%),
                var(--bg);
        }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 24px 20px 36px; }}
        .hero {{
            background: linear-gradient(120deg, #0f766e, #155e75 60%, #1d4e89);
            color: #fff;
            border-radius: 16px;
            padding: 22px 24px;
            box-shadow: 0 12px 28px rgba(16, 40, 60, 0.18);
        }}
        .hero h1 {{ margin: 0; font-size: 26px; letter-spacing: 0.2px; }}
        .hero p {{ margin: 10px 0 0; opacity: 0.92; font-size: 14px; }}
        .kpis {{
            margin-top: 16px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
            gap: 12px;
        }}
        .online-box {{
            margin-top: 16px;
            background: var(--panel);
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: 14px;
            box-shadow: 0 8px 20px rgba(31, 54, 78, 0.08);
        }}
        .online-header {{ margin-bottom: 8px; }}
        .online-title {{ font-size: 18px; font-weight: 700; color: #18455a; }}
        .online-sub {{ margin-top: 4px; color: var(--muted); font-size: 12px; }}
        .online-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 10px;
            margin-bottom: 8px;
        }}
        .online-history {{ margin-top: 6px; }}
        .online-chip {{
            border: 1px solid var(--line);
            border-radius: 10px;
            padding: 10px;
            background: linear-gradient(135deg, #fbfdfd, #f5faf9);
        }}
        .online-chip .name {{ font-size: 13px; color: #35566a; }}
        .online-chip .price {{ font-size: 24px; font-weight: 700; margin-top: 4px; color: #12303f; }}
        .online-chip .meta {{ margin-top: 4px; font-size: 12px; color: var(--muted); }}
        .online-chip .up {{ color: #b91c1c; font-weight: 700; }}
        .online-chip .down {{ color: #0f766e; font-weight: 700; }}
        .online-chip .flat {{ color: #475569; font-weight: 700; }}
        .kpi {{
            background: var(--panel);
            border: 1px solid var(--line);
            border-radius: 12px;
            padding: 14px 16px;
            box-shadow: 0 8px 18px rgba(25, 45, 65, 0.06);
        }}
        .kpi .label {{ color: var(--muted); font-size: 13px; }}
        .kpi .value {{ margin-top: 6px; font-size: 26px; font-weight: 700; color: #12303f; }}
        .panel {{
            margin-top: 16px;
            background: var(--panel);
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: 14px;
            box-shadow: 0 8px 20px rgba(31, 54, 78, 0.08);
        }}
        .panel h2 {{ margin: 4px 0 12px; font-size: 18px; }}
        .list-grid {{
            margin-top: 16px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 12px;
        }}
        .list-card {{
            background: var(--panel);
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: 14px;
        }}
        .list-card h3 {{ margin: 0 0 10px; font-size: 16px; color: #18455a; }}
        .spec-list {{ list-style: none; margin: 0; padding: 0; display: grid; gap: 8px; }}
        .spec-list li {{
            border: 1px solid var(--line);
            border-radius: 10px;
            padding: 10px;
            background: linear-gradient(135deg, #fbfdfd, #f5faf9);
            display: grid;
            gap: 3px;
        }}
        .spec {{ font-weight: 700; font-size: 14px; }}
        .price {{ color: var(--brand); font-weight: 700; }}
        .meta {{ color: var(--muted); font-size: 12px; }}
        .table-block {{ margin-top: 16px; }}
        .table-head {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
            gap: 10px;
        }}
        .table-head h3 {{ margin: 0; font-size: 16px; }}
        .table-wrap {{ overflow: auto; border: 1px solid var(--line); border-radius: 10px; }}
        table.data-table {{ width: 100%; border-collapse: collapse; min-width: 760px; }}
        table.data-table thead th {{
            position: sticky; top: 0;
            background: #eaf3f8;
            color: #1a3d53;
            z-index: 2;
            font-weight: 700;
        }}
        table.data-table th, table.data-table td {{
            border-bottom: 1px solid var(--line);
            padding: 9px 10px;
            text-align: left;
            white-space: nowrap;
            font-size: 13px;
        }}
        table.data-table tbody tr:nth-child(even) {{ background: #f8fbfd; }}
        .empty-tip {{ color: var(--muted); margin: 8px 0 0; }}
        @media (max-width: 780px) {{
            .container {{ padding: 14px 12px 24px; }}
            .hero h1 {{ font-size: 20px; }}
            .kpi .value {{ font-size: 22px; }}
            .panel {{ padding: 10px; }}
            table.data-table {{ min-width: 640px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <section class="hero">
            <h1>钢筋价格趋势分析看板</h1>
            <p>口径：仅统计 kg 单位，按规格聚合月均价，并自动过滤离散值（数量级 + IQR）。</p>
        </section>

        {online_html}

        <section class="kpis">
            <div class="kpi"><div class="label">最新月份</div><div class="value">{escape(str(latest_month))}</div></div>
            <div class="kpi"><div class="label">清洗后有效行</div><div class="value">{clean_count:,}</div></div>
            <div class="kpi"><div class="label">离散值行数</div><div class="value">{outlier_count:,}</div></div>
            <div class="kpi"><div class="label">离散值占比</div><div class="value">{outlier_ratio:.1f}%</div></div>
            <div class="kpi"><div class="label">识别规格数</div><div class="value">{spec_count:,}</div></div>
        </section>

        <section class="panel">
            <h2>钢筋 月度价格趋势图</h2>
            {material_fig_html.get('钢筋', "<p class='empty-tip'>暂无钢筋数据</p>")}
        </section>

        <section class="panel">
            <h2>螺纹钢 月度价格趋势图</h2>
            {material_fig_html.get('螺纹钢', "<p class='empty-tip'>暂无螺纹钢数据</p>")}
        </section>

        <section class="panel">
            <h2>钢筋网片 月度价格趋势图</h2>
            {material_fig_html.get('钢筋网片', "<p class='empty-tip'>暂无钢筋网片数据</p>")}
        </section>

        <section class="list-grid">
            {''.join(top_list_blocks)}
        </section>

        <section class="panel table-block">
            <div class="table-head"><h3>规格总览列表</h3></div>
            <div class="table-wrap">{overview_table_html}</div>
        </section>

        <section class="panel table-block">
            <div class="table-head"><h3>最新规格价格列表</h3></div>
            <div class="table-wrap">{latest_table_html}</div>
        </section>
    </div>
</body>
</html>
"""

    html_out.write_text(html, encoding="utf-8")


def build_material_figure(summary: pd.DataFrame, material: str) -> go.Figure | None:
    part = summary[summary["材料类型"] == material]
    if part.empty:
        return None

    fig = go.Figure()
    for spec in sorted(part["规格组"].unique()):
        d = part[part["规格组"] == spec].sort_values("年-月")
        fig.add_trace(
            go.Scatter(
                x=d["年-月"],
                y=d["平均单价"],
                mode="lines+markers",
                line=dict(width=2.2),
                marker=dict(size=7),
                name=str(spec),
                legendgroup=material,
                hovertemplate=(
                    f"材料: {material}<br>规格: {spec}<br>月份: %{{x}}<br>平均单价: %{{y:.2f}}<br>"
                    "样本数: %{customdata}<extra></extra>"
                ),
                customdata=d["样本数"],
            )
        )

    fig.update_layout(
        height=420,
        title=f"{material} 按规格月均价趋势（kg口径，已过滤离散值）",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=35, r=20, t=75, b=45),
        hovermode="x unified",
        template="plotly_white",
    )
    fig.update_yaxes(title_text="平均单价")
    fig.update_xaxes(title_text="年-月")
    return fig


def main():
    args = parse_args()
    if args.input_file:
        input_file = Path(args.input_file).expanduser().resolve()
    else:
        latest = find_latest_summary_file()
        if latest is None:
            raise FileNotFoundError(f"未在 {SUMMARY_INPUT_ROOT} 找到 {INPUT_GLOB}")
        input_file = latest.resolve()

    if not input_file.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    excel_out, html_out = build_output_paths(input_file, args.excel_suffix, args.html_suffix)

    print(f"读取文件: {input_file}")
    df = pd.read_excel(input_file, sheet_name=args.sheet)
    df.columns = [str(c).strip() for c in df.columns]

    cols = list(df.columns)
    item_col = find_column(cols, ["物品名称", "物料名称", "材料名称"])
    material_class_col = find_column(cols, ["材料归类", "物资归类", "物料归类"])
    spec_col = find_column(cols, ["规格", "规格型号", "物品规格"])
    unit_col = find_column(cols, ["标准单位", "单位标准化", "单位"])
    price_col = find_column(cols, ["单价", "含税单价", "不含税单价"])
    year_col = find_column(cols, ["年", "年份"])
    month_col = find_column(cols, ["月", "月份"])
    year_month_col = find_column(cols, ["年-月", "年月", "月份", "标准日期"])

    required = {
        "物品名称": item_col,
        "规格": spec_col,
        "单位": unit_col,
        "单价": price_col,
    }
    if year_month_col is None and (year_col is None or month_col is None):
        required["年/月或年月"] = None

    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(f"缺少必要列: {missing}。当前列: {cols}")

    select_cols = [item_col, spec_col, unit_col, price_col]
    if material_class_col:
        select_cols.append(material_class_col)
    if year_month_col:
        select_cols.append(year_month_col)
    else:
        select_cols.extend([year_col, month_col])

    work = df[select_cols].copy()
    rename_map = {
        item_col: "物品名称",
        spec_col: "规格",
        unit_col: "单位",
        price_col: "单价",
    }
    if material_class_col:
        rename_map[material_class_col] = "材料归类"
    if year_month_col:
        rename_map[year_month_col] = "年月"
    else:
        rename_map[year_col] = "年"
        rename_map[month_col] = "月"
    work = work.rename(columns=rename_map)

    work["材料类型"] = work.apply(
        lambda r: detect_material_from_class(r.get("材料归类", None), r["物品名称"]),
        axis=1,
    )
    work = work[work["材料类型"].notna()].copy()

    work = work[work["单位"].apply(is_kg_unit)].copy()
    work["单价"] = pd.to_numeric(work["单价"], errors="coerce")
    work = work[work["单价"].notna() & (work["单价"] > 0)].copy()

    if "年月" in work.columns:
        work["年-月"] = work["年月"].apply(parse_year_month_text)
        month_source_desc = "年月"
    else:
        work["年-月"] = work.apply(lambda r: month_text(r["年"], r["月"]), axis=1)
        month_source_desc = "年 + 月"
    work = work[work["年-月"].notna()].copy()

    work["规格组"] = None
    idx_mesh = work["材料类型"].eq("钢筋网片")
    idx_other = ~idx_mesh
    work.loc[idx_mesh, "规格组"] = work.loc[idx_mesh, "规格"].apply(canonical_mesh_spec)
    work.loc[idx_other, "规格组"] = work.loc[idx_other, "规格"].apply(canonical_diameter)
    work = work[work["规格组"].notna()].copy()

    # 分组离散值过滤
    grp_cols = ["材料类型", "规格组", "年-月"]
    work["保留"] = (
        work.groupby(grp_cols)["单价"]
        .transform(
            lambda x: outlier_filter_series(
                x,
                mag_factor=max(float(args.outlier_mag_factor), 1.0),
                iqr_k=max(float(args.outlier_iqr_k), 0.5),
                iqr_min_n=max(int(args.outlier_iqr_min_n), 3),
            )
        )
        .fillna(False)
        .astype(bool)
    )

    clean_df = work[work["保留"]].copy()
    outlier_df = work[~work["保留"]].copy()

    summary = (
        clean_df.groupby(["材料类型", "规格组", "年-月"], as_index=False)
        .agg(
            样本数=("单价", "count"),
            平均单价=("单价", "mean"),
            中位单价=("单价", "median"),
            最低单价=("单价", "min"),
            最高单价=("单价", "max"),
        )
        .sort_values(["材料类型", "规格组", "年-月"])
    )

    overview = (
        clean_df.groupby(["材料类型", "规格组"], as_index=False)
        .agg(
            月份数=("年-月", "nunique"),
            总样本数=("单价", "count"),
            总均价=("单价", "mean"),
        )
        .sort_values(["材料类型", "规格组"])
    )

    # 导出 Excel
    with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="清洗后明细", index=False)
        outlier_df.to_excel(writer, sheet_name="离散值明细", index=False)
        summary.to_excel(writer, sheet_name="月度规格均价", index=False)
        overview.to_excel(writer, sheet_name="规格总览", index=False)

        cfg = pd.DataFrame(
            [
                ["月份来源", month_source_desc],
                ["单价列", "单价"],
                ["单位过滤", "仅 kg / 公斤 / 千克"],
                ["材料分组", "钢筋、螺纹钢、钢筋网片"],
                ["钢筋规格", "按 Φ直径自动识别"],
                ["网片规格", "优先 Φd@间距-片尺寸；缺失片尺寸时保留已识别规格"],
                [
                    "离散值过滤",
                    f"中位数数量级过滤(÷{args.outlier_mag_factor:g}~x{args.outlier_mag_factor:g})"
                    f" + 样本>={max(int(args.outlier_iqr_min_n), 3)}时{args.outlier_iqr_k:g}倍IQR过滤",
                ],
            ],
            columns=["参数", "取值"],
        )
        cfg.to_excel(writer, sheet_name="口径说明", index=False)

    # 抓取在线历史行情
    print("拉取在线钢材期货历史数据…")
    history_data, history_err = fetch_online_steel_history()
    if history_err:
        print(f"  历史行情部分失败: {history_err}")
    history_fig = build_online_history_fig(history_data)
    online_history_fig_html = (
        history_fig.to_html(
            full_html=False,
            include_plotlyjs=False,
            config={"displaylogo": False, "responsive": True},
        )
        if history_fig is not None
        else ""
    )

    # 导出 HTML 图（3个独立图）
    material_order = ["钢筋", "螺纹钢", "钢筋网片"]
    material_fig_html: dict[str, str] = {}
    for mat in material_order:
        fig = build_material_figure(summary, mat)
        if fig is not None:
            material_fig_html[mat] = fig.to_html(
                full_html=False,
                include_plotlyjs="cdn",
                config={"displaylogo": False, "responsive": True},
            )

    build_html_dashboard(
        html_out=html_out,
        material_fig_html=material_fig_html,
        clean_df=clean_df,
        outlier_df=outlier_df,
        summary=summary,
        overview=overview,
        online_history_fig_html=online_history_fig_html,
    )

    print("\n处理完成")
    print(f"输出 Excel: {excel_out}")
    print(f"输出 HTML : {html_out}")
    print(f"原始命中行: {len(work)}")
    print(f"清洗后行数: {len(clean_df)}")
    print(f"离散值行数: {len(outlier_df)}")
    print(f"规格数: {clean_df[['材料类型', '规格组']].drop_duplicates().shape[0]}")


if __name__ == "__main__":
    main()
