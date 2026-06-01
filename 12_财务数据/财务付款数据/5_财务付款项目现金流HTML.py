from __future__ import annotations

import json
from html import escape
from pathlib import Path
from typing import Dict, List

import pandas as pd


class Config:
    INPUT_ROOT = Path(
        "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/B_财务付款_std"
    )
    OUTPUT_HTML = Path(
        "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/C_财务付款_sum/财务付款项目现金流量分析.html"
    )
    SHEET_CANDIDATES = ["明细数据", "明细数据_万泰铢"]


REQUIRED_COLUMNS = ["项目编码", "项目名称", "日期", "实付金额", "增值税额"]


def iter_excel_files(input_root: Path) -> List[Path]:
    files: List[Path] = []
    for pattern in ("*.xlsx", "*.xls", "*.xlsm"):
        files.extend(path for path in input_root.rglob(pattern) if path.is_file() and not path.name.startswith("~$"))
    return sorted(set(files))


def choose_sheet_name(excel_file: Path) -> str | None:
    xls = pd.ExcelFile(excel_file)

    for sheet_name in Config.SHEET_CANDIDATES:
        if sheet_name in xls.sheet_names:
            return sheet_name

    for sheet_name in xls.sheet_names:
        preview = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=1)
        if all(col in preview.columns for col in REQUIRED_COLUMNS):
            return sheet_name

    return None


def read_single_file(excel_file: Path) -> pd.DataFrame:
    sheet_name = choose_sheet_name(excel_file)
    if not sheet_name:
        print(f"跳过文件(缺少目标sheet): {excel_file}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        print(f"跳过文件(缺少列 {missing}): {excel_file}")
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

    out = df[REQUIRED_COLUMNS].copy()
    out["来源文件"] = excel_file.name
    return out


def build_cashflow_df(input_root: Path) -> pd.DataFrame:
    excel_files = iter_excel_files(input_root)
    if not excel_files:
        raise FileNotFoundError(f"未找到Excel文件: {input_root}")

    print(f"扫描到文件数量: {len(excel_files)}")
    all_frames = [read_single_file(file_path) for file_path in excel_files]
    merged = pd.concat(all_frames, ignore_index=True)
    if merged.empty:
        raise ValueError("没有可用于汇总的数据。")

    merged["日期"] = pd.to_datetime(merged["日期"], errors="coerce")
    merged = merged.dropna(subset=["日期"])

    merged["项目编码"] = merged["项目编码"].astype(str).str.strip()
    merged["项目名称"] = merged["项目名称"].astype(str).str.strip()

    merged["实付金额"] = pd.to_numeric(merged["实付金额"], errors="coerce").fillna(0.0)
    merged["增值税额"] = pd.to_numeric(merged["增值税额"], errors="coerce").fillna(0.0)
    merged["实际现金流"] = merged["实付金额"] + merged["增值税额"]

    grouped = (
        merged.groupby(["项目编码", "项目名称", "日期"], as_index=False)["实际现金流"]
        .sum()
        .sort_values(["项目编码", "项目名称", "日期"])
    )

    grouped["日期"] = grouped["日期"].dt.strftime("%Y-%m-%d")
    return grouped


def build_project_payload(cashflow_df: pd.DataFrame) -> Dict[str, Dict]:
    payload: Dict[str, Dict] = {}

    projects = (
        cashflow_df[["项目编码", "项目名称"]]
        .drop_duplicates()
        .sort_values(["项目编码", "项目名称"])
    )

    for idx, row in projects.reset_index(drop=True).iterrows():
        project_code = str(row["项目编码"]).strip()
        project_name = str(row["项目名称"]).strip()
        key = f"p{idx:04d}"

        project_df = cashflow_df[
            (cashflow_df["项目编码"] == project_code)
            & (cashflow_df["项目名称"] == project_name)
        ].copy().sort_values("日期")

        project_df["累计实际现金流"] = project_df["实际现金流"].cumsum()
        project_df["累计实际现金流_万"] = (project_df["累计实际现金流"] / 10000.0).round(0)

        total = float(project_df["实际现金流"].sum())
        total_wan = round(total / 10000.0)
        payload[key] = {
            "projectCode": project_code,
            "projectName": project_name,
            "dates": project_df["日期"].tolist(),
          "values": [int(v) for v in project_df["累计实际现金流_万"].tolist()],
          "total": round(total, 2),
          "totalWan": int(total_wan),
        }

    return payload


def render_html(payload: Dict[str, Dict], output_html: Path) -> None:
    nav_links_html = []
    sections_html = []

    for key, item in payload.items():
        label = f"{item['projectCode']} | {item['projectName']}"
        nav_links_html.append(
            f"<a class=\"nav-link\" href=\"#section-{key}\">{escape(label)}</a>"
        )

        section_title = (
            f"{escape(item['projectName'])} "
            f"<span class='code'>({escape(item['projectCode'])})</span>"
        )
        sections_html.append(
            f"""
            <section class=\"project-row\" id=\"section-{key}\">
              <h2>{section_title}</h2>
              <p class=\"meta\">累计实际现金流（实付金额 + 增值税额）：{item['totalWan']:,} 万</p>
              <div class=\"chart\" id=\"chart-{key}\"></div>
            </section>
            """.strip()
        )

    payload_json = json.dumps(payload, ensure_ascii=False)
    total_projects = len(payload)

    html = f"""
<!doctype html>
<html lang=\"zh-CN\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>财务付款项目现金流量分析</title>
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <style>
    :root {{
      --bg: #f8f4ea;
      --card: #fffdfa;
      --ink: #26231f;
      --muted: #6d675f;
      --accent: #bc4f2f;
      --line: #e5dac6;
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
      display: grid;
      grid-template-columns: minmax(0, 1fr) 300px;
      gap: 16px;
      align-items: start;
    }}
    .main-list {{ display: flex; flex-direction: column; gap: 12px; }}
    h1 {{
      margin: 0 0 8px;
      letter-spacing: 0.3px;
      font-size: clamp(24px, 2.8vw, 36px);
    }}
    .sub {{ margin: 0 0 16px; color: var(--muted); }}
    .sidebar {{
      position: sticky;
      top: 12px;
      border: 1px solid var(--line);
      background: var(--card);
      border-radius: 14px;
      padding: 12px;
      box-shadow: 0 6px 22px rgba(97, 76, 42, 0.09);
      max-height: calc(100vh - 24px);
      overflow: auto;
    }}
    .sidebar h3 {{
      margin: 0 0 10px;
      font-size: 15px;
      color: #503722;
    }}
    .nav-list {{
      display: flex;
      flex-direction: column;
      gap: 6px;
    }}
    .nav-link {{
      text-decoration: none;
      color: #4f3e2e;
      background: #fff9ef;
      border: 1px solid #eadcc6;
      border-radius: 8px;
      font-size: 13px;
      line-height: 1.35;
      padding: 7px 8px;
    }}
    .nav-link:hover {{
      background: #f4e3c8;
      border-color: #d7b786;
    }}
    .project-row {{
      border: 1px solid var(--line);
      background: var(--card);
      border-radius: 14px;
      padding: 14px;
      box-shadow: 0 6px 22px rgba(97, 76, 42, 0.09);
      scroll-margin-top: 84px;
    }}
    .project-row h2 {{ margin: 0 0 4px; font-size: 18px; }}
    .code {{ color: var(--muted); font-size: 14px; font-weight: normal; }}
    .meta {{ margin: 0 0 10px; color: #5e4631; font-size: 13px; }}
    .chart {{ width: 100%; height: 300px; }}
    @media (max-width: 980px) {{
      .page-layout {{ grid-template-columns: 1fr; }}
      .sidebar {{ position: static; max-height: none; }}
    }}
    @media (max-width: 600px) {{
      .chart {{ height: 260px; }}
    }}
  </style>
</head>
<body>
  <div class=\"wrap\">
    <h1>财务付款项目实际现金流量分析</h1>
    <p class=\"sub\">项目数：{total_projects}。口径：先按日期汇总（实付金额 + 增值税额），再按日期累计形成现金流曲线；展示单位为万（无小数）。</p>

    <div class=\"page-layout\">
      <div class=\"main-list\">
        {''.join(sections_html)}
      </div>

      <aside class=\"sidebar\">
        <h3>项目选择（右侧固定）</h3>
        <div class=\"nav-list\">{''.join(nav_links_html)}</div>
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
        x: item.dates,
        y: item.values,
        mode: 'lines+markers',
        type: 'scatter',
        line: {{ color: '#bc4f2f', width: 2 }},
        marker: {{ size: 5, color: '#7f2910' }},
        hovertemplate: '日期: %{{x|%Y年%-m月%-d日}}<br>累计实际现金流: %{{y:,.0f}} 万<extra></extra>'
      }};
      const layout = {{
        margin: {{ l: 52, r: 16, t: 10, b: 40 }},
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor: '#fffdf8',
        xaxis: {{
          title: '日期（年月）',
          type: 'date',
          tickmode: 'linear',
          tick0: xMin,
          dtick: 'M1',
          tickformat: '%Y年%-m月',
          tickangle: -25,
          showgrid: true,
          gridcolor: '#efe5d5',
          range: [xMin, xMax]
        }},
        yaxis: {{
          title: '累计实际现金流（万）',
          tickformat: ',.0f',
          separatethousands: true,
          gridcolor: '#efe5d5'
        }},
      }};
      const config = {{ displaylogo: false, responsive: true }};
      Plotly.newPlot(`chart-${{key}}`, [trace], layout, config);
    }}

    Object.keys(projectData).forEach(renderChart);
  </script>
</body>
</html>
    """.strip()

    output_html.parent.mkdir(parents=True, exist_ok=True)
    output_html.write_text(html, encoding="utf-8")


def main() -> None:
    cashflow_df = build_cashflow_df(Config.INPUT_ROOT)
    payload = build_project_payload(cashflow_df)
    render_html(payload, Config.OUTPUT_HTML)

    print(f"完成: {Config.OUTPUT_HTML}")
    print(f"项目数量: {len(payload)}")


if __name__ == "__main__":
    main()
