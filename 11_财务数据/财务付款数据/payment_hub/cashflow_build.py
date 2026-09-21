# -*- coding: utf-8 -*-
"""
步骤 cashflow：B_std → 项目现金流 HTML（原 5_）
"""

from __future__ import annotations

import json
from html import escape
from pathlib import Path
from typing import Dict, List

import pandas as pd

from .config import Paths
from .utils import iter_excel_files, resolve_business_groups

REQUIRED_COLUMNS = ["项目编码", "项目名称", "日期", "实付金额", "增值税额"]
SHEET_CANDIDATES = ["明细数据", "明细数据_万泰铢"]


def choose_sheet_name(excel_file: Path) -> str | None:
    xls = pd.ExcelFile(excel_file)
    for sheet_name in SHEET_CANDIDATES:
        if sheet_name in xls.sheet_names:
            return sheet_name
    for sheet_name in xls.sheet_names:
        preview = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=1)
        if all(col in preview.columns for col in REQUIRED_COLUMNS):
            return str(sheet_name)
    return None


def read_single_file(excel_file: Path) -> pd.DataFrame:
    sheet_name = choose_sheet_name(excel_file)
    if not sheet_name:
        print(f"跳过(无目标sheet): {excel_file.name}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)
    # 现金流金额必须用「明细数据」原币，不能用万泰铢表
    if sheet_name == "明细数据_万泰铢" and "明细数据" in pd.ExcelFile(excel_file).sheet_names:
        sheet_name = "明细数据"
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print(f"跳过(缺列 {missing}): {excel_file.name}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)
    out = df[REQUIRED_COLUMNS].copy()
    out["来源文件"] = excel_file.name
    return out


def build_cashflow_df(input_root: Path) -> pd.DataFrame:
    files = iter_excel_files(input_root)
    if not files:
        raise FileNotFoundError(f"未找到 Excel: {input_root}")
    print(f"扫描到文件数量: {len(files)}")
    frames = [read_single_file(fp) for fp in files]
    merged = pd.concat(frames, ignore_index=True)
    if merged.empty:
        raise ValueError("没有可用于汇总的数据")

    merged["日期"] = pd.to_datetime(merged["日期"], errors="coerce")
    merged = merged.dropna(subset=["日期"])
    merged["项目编码"] = merged["项目编码"].astype(str).str.strip()
    merged["项目名称"] = merged["项目名称"].astype(str).str.strip()
    merged["实付金额"] = pd.to_numeric(merged["实付金额"], errors="coerce").fillna(0.0)
    merged["增值税额"] = pd.to_numeric(merged["增值税额"], errors="coerce").fillna(0.0)
    merged["实际现金流"] = merged["实付金额"] + merged["增值税额"]

    grouped = (
        merged.groupby(["项目编码", "项目名称", "日期"], as_index=False)
        .agg(实际现金流=("实际现金流", "sum"))
        .sort_values(by=["项目编码", "项目名称", "日期"])
    )
    grouped["日期"] = grouped["日期"].dt.strftime("%Y-%m-%d")
    return grouped


def load_project_status(project_info: Path) -> Dict[str, str]:
    if not project_info.exists():
        print(f"警告：项目信息表不存在，跳过状态: {project_info}")
        return {}
    df = pd.read_excel(project_info, dtype={"项目编码": str})
    if "项目编码" not in df.columns:
        return {}
    df["项目编码"] = df["项目编码"].astype(str).str.strip()
    status_col = "项目状态" if "项目状态" in df.columns else None
    if not status_col:
        return {}
    df[status_col] = df[status_col].astype(str).str.strip().fillna("未知")
    return dict(zip(df["项目编码"], df[status_col]))


def build_project_payload(
    cashflow_df: pd.DataFrame, status_map: Dict[str, str]
) -> Dict[str, Dict]:
    payload: Dict[str, Dict] = {}
    projects = (
        cashflow_df[["项目编码", "项目名称"]]
        .drop_duplicates()
        .sort_values(["项目编码", "项目名称"], ascending=False)
    )
    for idx, row in projects.reset_index(drop=True).iterrows():
        code = str(row["项目编码"]).strip()
        name = str(row["项目名称"]).strip()
        key = f"p{idx:04d}"
        pdf = (
            cashflow_df[
                (cashflow_df["项目编码"] == code) & (cashflow_df["项目名称"] == name)
            ]
            .copy()
            .sort_values("日期")
        )
        pdf["累计实际现金流"] = pdf["实际现金流"].cumsum()
        pdf["累计实际现金流_万"] = (pdf["累计实际现金流"] / 10000.0).round(0)
        total = float(pdf["实际现金流"].sum())
        payload[key] = {
            "projectCode": code,
            "projectName": name,
            "projectStatus": status_map.get(code, "未知"),
            "dates": pdf["日期"].tolist(),
            "values": [int(v) for v in pdf["累计实际现金流_万"].tolist()],
            "total": round(total, 2),
            "totalWan": int(round(total / 10000.0)),
        }
    return payload


def render_html(payload: Dict[str, Dict], output_html: Path) -> None:
    nav_links_html: List[str] = []
    sections_html: List[str] = []

    for key, item in payload.items():
        label = f"{item['projectCode']} | {item['projectName']}"
        status_escaped = escape(item["projectStatus"])
        nav_links_html.append(
            f'<a class="nav-link" href="#section-{key}" '
            f'data-status="{status_escaped}">{escape(label)}</a>'
        )
        section_title = (
            f"{escape(item['projectName'])} "
            f"<span class='code'>({escape(item['projectCode'])})</span>"
        )
        sections_html.append(
            f"""
            <section class="project-row" id="section-{key}" data-status="{status_escaped}">
              <h2>{section_title}</h2>
              <p class="meta">项目状态：{status_escaped} &nbsp;｜&nbsp; 累计实际现金流（实付金额 + 增值税额）：{item['totalWan']:,} 万</p>
              <div class="chart" id="chart-{key}"></div>
            </section>
            """.strip()
        )

    payload_json = json.dumps(payload, ensure_ascii=False)
    total_projects = len(payload)
    seen_statuses: List[str] = []
    for item in payload.values():
        s = item["projectStatus"]
        if s not in seen_statuses:
            seen_statuses.append(s)
    status_buttons_html = '<button class="status-btn active" data-filter="__all__">全部</button>'
    for s in seen_statuses:
        status_buttons_html += (
            f'<button class="status-btn" data-filter="{escape(s)}">{escape(s)}</button>'
        )

    html = f"""
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>财务付款项目现金流量分析</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root {{
      --bg: #f8f4ea; --card: #fffdfa; --ink: #26231f;
      --muted: #6d675f; --accent: #bc4f2f; --line: #e5dac6;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Noto Serif SC", "Source Han Serif SC", "PingFang SC", serif;
      color: var(--ink);
      background:
        radial-gradient(80rem 60rem at 10% -10%, #f5e3c9 0, transparent 60%),
        radial-gradient(50rem 50rem at 100% 0%, #f7edd9 0, transparent 50%),
        var(--bg);
    }}
    .wrap {{ max-width: 1500px; margin: 0 auto; padding: 28px 16px 56px; }}
    .page-layout {{
      display: grid; grid-template-columns: minmax(0, 1fr) 300px;
      gap: 16px; align-items: start;
    }}
    .main-list {{ display: flex; flex-direction: column; gap: 12px; }}
    h1 {{ margin: 0 0 8px; letter-spacing: 0.3px; font-size: clamp(24px, 2.8vw, 36px); }}
    .sub {{ margin: 0 0 10px; color: var(--muted); }}
    .status-bar {{ display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 16px; }}
    .status-btn {{
      padding: 6px 14px; border: 1px solid #d4bfa0; border-radius: 20px;
      background: #fff9ef; color: #4f3e2e; font-size: 13px; cursor: pointer;
      font-family: inherit; transition: background 0.15s, border-color 0.15s;
    }}
    .status-btn:hover {{ background: #f4e3c8; border-color: #c09060; }}
    .status-btn.active {{ background: #bc4f2f; border-color: #bc4f2f; color: #fff; }}
    .sidebar {{
      position: sticky; top: 12px; border: 1px solid var(--line);
      background: var(--card); border-radius: 14px; padding: 12px;
      box-shadow: 0 6px 22px rgba(97, 76, 42, 0.09);
      max-height: calc(100vh - 24px); overflow: auto;
    }}
    .sidebar h3 {{ margin: 0 0 10px; font-size: 15px; color: #503722; }}
    .nav-list {{ display: flex; flex-direction: column; gap: 6px; }}
    .nav-link {{
      text-decoration: none; color: #4f3e2e; background: #fff9ef;
      border: 1px solid #eadcc6; border-radius: 8px; font-size: 13px;
      line-height: 1.35; padding: 7px 8px;
    }}
    .nav-link:hover {{ background: #f4e3c8; border-color: #d7b786; }}
    .project-row {{
      border: 1px solid var(--line); background: var(--card); border-radius: 14px;
      padding: 14px; box-shadow: 0 6px 22px rgba(97, 76, 42, 0.09);
      scroll-margin-top: 84px;
    }}
    .project-row h2 {{ margin: 0 0 4px; font-size: 18px; }}
    .code {{ color: var(--muted); font-size: 14px; font-weight: normal; }}
    .meta {{ margin: 0 0 10px; color: #5e4631; font-size: 13px; }}
    .chart {{ width: 100%; height: 340px; }}
    @media (max-width: 980px) {{
      .page-layout {{ grid-template-columns: 1fr; }}
      .sidebar {{ position: static; max-height: none; }}
    }}
    @media (max-width: 600px) {{ .chart {{ height: 300px; }} }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>财务付款项目实际现金流量分析</h1>
    <p class="sub">项目数：{total_projects}。口径：先按日期汇总（实付金额 + 增值税额），再按日期累计形成现金流曲线；展示单位为万（无小数）。</p>
    <div class="status-bar">{status_buttons_html}</div>
    <div class="page-layout">
      <div class="main-list" id="main-list">{''.join(sections_html)}</div>
      <aside class="sidebar">
        <h3>项目选择（右侧固定）</h3>
        <div class="nav-list" id="nav-list">{''.join(nav_links_html)}</div>
      </aside>
    </div>
  </div>
  <script>
    const projectData = {payload_json};
    function monthStart(dateStr) {{
      const d = new Date(`${{dateStr}}T00:00:00`);
      return new Date(d.getFullYear(), d.getMonth(), 1);
    }}
    function monthEnd(dateStr) {{
      const d = new Date(`${{dateStr}}T00:00:00`);
      return new Date(d.getFullYear(), d.getMonth() + 1, 0);
    }}
    function renderChart(key) {{
      const item = projectData[key];
      const xMin = monthStart(item.dates[0]);
      const xMax = monthEnd(item.dates[item.dates.length - 1]);
      const trace = {{
        x: item.dates, y: item.values, mode: 'lines+markers', type: 'scatter',
        line: {{ color: '#bc4f2f', width: 2 }},
        marker: {{ size: 5, color: '#7f2910' }},
        hovertemplate: '日期: %{{x|%Y年%-m月%-d日}}<br>累计实际现金流: %{{y:,.0f}} 万<extra></extra>'
      }};
      const layout = {{
        margin: {{ l: 72, r: 16, t: 12, b: 72 }},
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: '#fffdf8',
        xaxis: {{
          title: '日期（年月）', type: 'date', tickmode: 'linear',
          tick0: xMin, dtick: 'M1', tickformat: '%Y年%-m月', tickangle: -35,
          tickfont: {{ size: 9 }}, automargin: true, titlefont: {{ size: 12 }},
          title_standoff: 14, showgrid: true, gridcolor: '#efe5d5', range: [xMin, xMax]
        }},
        yaxis: {{
          title: '累计实际现金流（万）', tickformat: ',.0f', separatethousands: true,
          tickfont: {{ size: 10 }}, automargin: true, titlefont: {{ size: 12 }},
          title_standoff: 14, gridcolor: '#efe5d5'
        }},
      }};
      Plotly.newPlot(`chart-${{key}}`, [trace], layout, {{ displaylogo: false, responsive: true }});
    }}
    Object.keys(projectData).forEach(renderChart);
    function applyFilter(filter) {{
      document.querySelectorAll('.status-btn').forEach(btn => {{
        btn.classList.toggle('active', btn.dataset.filter === filter);
      }});
      document.querySelectorAll('#main-list .project-row').forEach(sec => {{
        sec.style.display = (filter === '__all__' || sec.dataset.status === filter) ? '' : 'none';
      }});
      document.querySelectorAll('#nav-list .nav-link').forEach(a => {{
        a.style.display = (filter === '__all__' || a.dataset.status === filter) ? '' : 'none';
      }});
    }}
    document.querySelectorAll('.status-btn').forEach(btn => {{
      btn.addEventListener('click', () => applyFilter(btn.dataset.filter));
    }});
  </script>
</body>
</html>
    """.strip()

    output_html.parent.mkdir(parents=True, exist_ok=True)
    output_html.write_text(html, encoding="utf-8")


def run_cashflow_for_group(
    group_std_dir: Path,
    out_html: Path,
    project_info: Path,
    *,
    group_label: str = "",
) -> str:
    label = group_label or Path(group_std_dir).name
    cashflow_df = build_cashflow_df(Path(group_std_dir))
    status_map = load_project_status(Path(project_info))
    payload = build_project_payload(cashflow_df, status_map)
    out_html = Path(out_html)
    render_html(payload, out_html)
    print(f"  ✅ [{label}] 现金流 HTML：{len(payload)} 个项目 → {out_html}")
    return str(out_html)


def run_cashflow(paths: Paths) -> list[str]:
    """
    按 A_raw 一级文件夹对应 B_std/{组名} 分别生成现金流 HTML：
      0_html/财务付款/{组名}_财务付款项目现金流量分析.html
    """
    groups = resolve_business_groups(paths.raw_dir, paths.std_dir)
    if not groups:
        print("未发现业务分组，跳过 cashflow")
        return []

    print(f"现金流按文件夹分组: {', '.join(g for g, _ in groups)}")
    written: list[str] = []
    project_info = Path(paths.project_info_xlsx)
    for group_name, group_std in groups:
        print(f"\n{'=' * 64}")
        print(f"  分组 cashflow · {group_name}  ← {group_std}")
        print(f"{'=' * 64}")
        out = run_cashflow_for_group(
            group_std,
            paths.group_cashflow_html(group_name),
            project_info,
            group_label=group_name,
        )
        written.append(out)

    print(f"\n✅ cashflow 全部完成：{len(written)} 个 HTML")
    return written
