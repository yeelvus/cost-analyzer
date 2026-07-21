# -*- coding: utf-8 -*-
"""
总汇总看板生成器

读取 11.8_广联达预算人材机文件统计-v2-正常运行.py 输出的“【总汇总】所有标段_*.xlsx”，
生成一个双层结构的交互式 HTML 看板：
1. 上方展示全项目总览，纵向列出所有项目，横向按固定列显示数据条
2. 点击项目名称后，跳转到下方项目详情区
3. 详情区展示项目金额卡片、结构占比和主材数量/金额卡片

使用示例:
python 总汇总看板生成.py
python 总汇总看板生成.py --run-source
python 总汇总看板生成.py --summary-file /path/to/【总汇总】所有标段_20260313_120000.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from budget.config import (
    DASHBOARD_HTML_NAME,
    GLODON_STD_DIR,
    MATERIAL_NAMES,
    MATERIAL_SHEET_NAME,
    TOTAL_SUMMARY_PREFIX,
)

DEFAULT_SOURCE_SCRIPT = Path(__file__).resolve().parent / "glodon" / "pipeline.py"
DEFAULT_OUTPUT_DIR = GLODON_STD_DIR
SUMMARY_PREFIX = TOTAL_SUMMARY_PREFIX
SUMMARY_SHEET = "总汇总"
MATERIAL_SHEET = MATERIAL_SHEET_NAME
DEFAULT_DASHBOARD_NAME = DASHBOARD_HTML_NAME
TEXT_COLUMNS = {"序号", "项目编码", "项目名称", "币种"}
SUMMARY_TOTAL_COLUMN = "标段人材机汇总"
DETAIL_AMOUNT_COLUMNS = ["人_明细", "管_明细", "设机_明细", "材类明细", "其他类别_明细"]
OVERVIEW_COLUMNS = [SUMMARY_TOTAL_COLUMN, *DETAIL_AMOUNT_COLUMNS]
SUMMARY_CARD_COLUMNS = [SUMMARY_TOTAL_COLUMN, *DETAIL_AMOUNT_COLUMNS]


def to_float(value):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return 0.0
    return float(num)


def run_source_script(source_script: Path | None = None):
    """先跑广联达 pipeline，再生成看板。"""
    if source_script and source_script.exists():
        cmd = [sys.executable, str(source_script)]
        subprocess.run(cmd, check=True)
        return
    from budget.glodon.pipeline import run_glodon_pipeline

    run_glodon_pipeline()


def extract_timestamp_from_name(file_name: str):
    match = re.search(r"(\d{8}_\d{6})", file_name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y%m%d_%H%M%S")
    except ValueError:
        return None


def find_latest_summary_file(output_dir: Path):
    if not output_dir.exists():
        raise FileNotFoundError(f"输出目录不存在: {output_dir}")

    candidates = []
    for path in output_dir.iterdir():
        if not path.is_file():
            continue
        if not path.name.startswith(SUMMARY_PREFIX):
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            continue
        candidates.append(path)

    if not candidates:
        raise FileNotFoundError(f"未找到总汇总文件: {output_dir}/{SUMMARY_PREFIX}*.xlsx")

    def sort_key(p: Path):
        ts = extract_timestamp_from_name(p.name)
        mtime = datetime.fromtimestamp(p.stat().st_mtime)
        return (ts or datetime.min, mtime)

    return max(candidates, key=sort_key)


def clean_and_split_summary(summary_file: Path):
    df = pd.DataFrame(pd.read_excel(summary_file, sheet_name=SUMMARY_SHEET, dtype=object))
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    if "项目名称" not in df.columns:
        raise ValueError(f"{SUMMARY_SHEET} 中缺少‘项目名称’列")

    for col in OVERVIEW_COLUMNS:
        if col not in df.columns:
            df[col] = 0.0

    numeric_target_cols = [col for col in OVERVIEW_COLUMNS if col in df.columns]
    for col in numeric_target_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    for col in ["序号", "项目编码", "项目名称", "币种"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    is_total = pd.Series(False, index=df.index)
    if "项目名称" in df.columns:
        is_total |= df["项目名称"].eq("合计")
    if "序号" in df.columns:
        is_total |= df["序号"].eq("合计")

    projects_df = df[~is_total].copy()
    total_df = df[is_total].copy()

    if projects_df.empty:
        raise ValueError("总汇总中无有效项目数据")

    numeric_cols = [c for c in OVERVIEW_COLUMNS if c in projects_df.columns]

    if not numeric_cols:
        raise ValueError("总汇总中无可用数值列")

    if not total_df.empty:
        total_row = total_df.iloc[0]
        totals = {c: round(to_float(total_row.get(c, 0.0)), 2) for c in numeric_cols}
    else:
        totals = {c: round(float(projects_df[c].sum()), 2) for c in numeric_cols}

    total_col = SUMMARY_TOTAL_COLUMN if SUMMARY_TOTAL_COLUMN in numeric_cols else numeric_cols[0]
    detail_cols = [c for c in DETAIL_AMOUNT_COLUMNS if c in numeric_cols]
    material_cols = MATERIAL_NAMES
    card_cols = [c for c in SUMMARY_CARD_COLUMNS if c in numeric_cols]
    overview_cols = [c for c in OVERVIEW_COLUMNS if c in numeric_cols]
    column_max = {c: round(float(projects_df[c].max()), 2) for c in overview_cols}

    return (
        projects_df,
        numeric_cols,
        totals,
        total_col,
        detail_cols,
        material_cols,
        card_cols,
        overview_cols,
        column_max,
    )


def load_material_map(summary_file: Path):
  try:
    mdf = pd.DataFrame(pd.read_excel(summary_file, sheet_name=MATERIAL_SHEET, dtype=object))
  except Exception:
    return {}, {}

  mdf = mdf.dropna(how="all").copy()
  if mdf.empty:
    return {}, {}

  for col in ["项目编码", "项目名称", "币种"]:
    if col in mdf.columns:
      mdf[col] = mdf[col].fillna("").astype(str).str.strip()

  if "项目名称" in mdf.columns:
    mdf = mdf[mdf["项目名称"] != "合计"].copy()

  for col in mdf.columns:
    if col.endswith("_数量") or col.endswith("_金额") or col.endswith("_平均单价"):
      mdf[col] = pd.to_numeric(mdf[col], errors="coerce").fillna(0.0)

  by_code = {}
  by_name = {}

  for _, row in mdf.iterrows():
    code = str(row.get("项目编码", "")).strip()
    name = str(row.get("项目名称", "")).strip()

    materials = []
    for mat in MATERIAL_NAMES:
      materials.append({
        "name": mat,
        "quantity": round(to_float(row.get(f"{mat}_数量", 0.0)), 2),
        "unit": str(row.get(f"{mat}_单位", "")).strip(),
        "amount": round(to_float(row.get(f"{mat}_金额", 0.0)), 2),
        "avg_price": round(to_float(row.get(f"{mat}_平均单价", 0.0)), 2),
      })

    if code:
      by_code[code] = materials
    if name:
      by_name[name] = materials

  return by_code, by_name


def build_payload(summary_file: Path):
    (
        projects_df,
        numeric_cols,
        totals,
        total_col,
        detail_cols,
        material_cols,
        card_cols,
        overview_cols,
        column_max,
    ) = clean_and_split_summary(summary_file)

    material_by_code, material_by_name = load_material_map(summary_file)

    projects = []
    project_rows = projects_df.reset_index(drop=True).to_dict(orient="records")
    for idx, row in enumerate(project_rows):
        code = str(row.get("项目编码", "")).strip()
        name = str(row.get("项目名称", "")).strip() or f"未命名项目{idx + 1}"

        materials = material_by_code.get(code) or material_by_name.get(name) or []

        projects.append({
            "id": idx,
            "序号": str(row.get("序号", "")).strip(),
            "项目编码": code,
            "项目名称": name,
            "币种": str(row.get("币种", "")).strip(),
            "metrics": {c: round(to_float(row.get(c, 0.0)), 2) for c in numeric_cols},
            "materials": materials,
        })

    payload = {
        "meta": {
            "source_file": summary_file.name,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "project_count": len(projects),
            "overall_total": round(to_float(totals.get(total_col, 0.0)), 2),
        },
        "columns": {
            "numeric_cols": numeric_cols,
            "total_col": total_col,
            "detail_cols": detail_cols,
            "material_cols": material_cols,
            "card_cols": card_cols,
            "overview_cols": overview_cols,
            "column_max": column_max,
        },
        "totals": totals,
        "projects": projects,
    }
    return payload


def build_html(payload):
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")

    template = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>项目总汇总看板</title>
  <style>
    :root {
      --bg-a: #f5f9ff;
      --bg-b: #eff7f2;
      --ink: #17324a;
      --muted: #60758a;
      --card: rgba(255, 255, 255, 0.84);
      --card-solid: rgba(255, 255, 255, 0.96);
      --shadow: 0 18px 48px rgba(28, 62, 93, 0.14);
      --radius-xl: 28px;
      --radius-lg: 22px;
      --radius-pill: 999px;
    }

    * { box-sizing: border-box; }

    html { scroll-behavior: smooth; }

    body {
      margin: 0;
      color: var(--ink);
      font-family: "Source Han Sans SC", "PingFang SC", "Microsoft YaHei", "Hiragino Sans GB", sans-serif;
      background:
        radial-gradient(1100px 560px at -10% -15%, #d8f5f0 0%, transparent 52%),
        radial-gradient(760px 400px at 115% 10%, #ffe6be 0%, transparent 54%),
        linear-gradient(160deg, var(--bg-a), var(--bg-b) 55%, #f7fbff 100%);
      min-height: 100vh;
    }

    .container {
      max-width: 1560px;
      margin: 0 auto;
      padding: 26px 20px 42px;
    }

    .layout {
      display: grid;
      grid-template-columns: 260px minmax(0, 1fr);
      gap: 16px;
      align-items: start;
    }

    .content-col {
      min-width: 0;
    }

    .toc {
      position: sticky;
      top: 14px;
      max-height: calc(100vh - 28px);
      overflow: auto;
      padding: 12px;
    }

    .toc-title {
      margin: 0 0 8px;
      font-size: 14px;
      color: #2f4338;
      font-weight: 700;
    }

    .toc-links {
      display: grid;
      gap: 6px;
    }

    .toc-link {
      display: block;
      padding: 6px 8px;
      border-radius: 8px;
      color: #2a473c;
      text-decoration: none;
      font-size: 12px;
      line-height: 1.35;
      background: rgba(227, 237, 230, 0.4);
    }

    .toc-link:hover {
      background: rgba(196, 106, 47, 0.18);
    }

    .panel {
      background: var(--card);
      border: 1px solid rgba(255, 255, 255, 0.92);
      border-radius: var(--radius-xl);
      box-shadow: var(--shadow);
      backdrop-filter: blur(10px);
      margin-bottom: 18px;
    }

    .hero {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
      padding: 24px;
      background:
        linear-gradient(135deg, rgba(255,255,255,0.94), rgba(245,255,250,0.88)),
        linear-gradient(135deg, rgba(15,139,141,0.08), rgba(37,99,235,0.06));
    }

    .eyebrow {
      margin: 0 0 10px;
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.2em;
      text-transform: uppercase;
      color: #0f766e;
    }

    .hero h1 {
      margin: 0;
      font-size: clamp(28px, 4vw, 42px);
      line-height: 1.05;
      letter-spacing: 0.02em;
    }

    .hero p {
      margin: 12px 0 0;
      color: var(--muted);
      font-size: 14px;
    }

    .hero-stats {
      display: grid;
      grid-template-columns: repeat(2, minmax(170px, 1fr));
      gap: 12px;
      min-width: min(100%, 380px);
    }

    .hero-stat {
      border-radius: 22px;
      padding: 16px 18px;
      background: rgba(255,255,255,0.92);
      border: 1px solid rgba(223, 235, 245, 0.95);
      box-shadow: 0 10px 28px rgba(23, 50, 74, 0.08);
    }

    .hero-stat span {
      display: block;
      font-size: 12px;
      color: var(--muted);
      margin-bottom: 8px;
    }

    .hero-stat strong {
      display: block;
      font-size: clamp(20px, 3vw, 28px);
      line-height: 1.1;
    }

    .section-head {
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 18px;
      padding: 22px 24px 0;
    }

    .section-head h2 {
      margin: 6px 0 0;
      font-size: 22px;
    }

    .section-head p {
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 14px;
    }

    .section-kicker {
      font-size: 12px;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      color: #0f766e;
      font-weight: 700;
    }

    .overview-shell {
      padding: 18px 24px 24px;
      overflow: auto;
    }

    .overview-table {
      min-width: 1460px;
      border-radius: 24px;
      overflow: hidden;
      border: 1px solid #dce8f2;
      background: rgba(255,255,255,0.78);
    }

    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
    }

    thead th {
      position: sticky;
      top: 0;
      z-index: 3;
      background: rgba(247, 252, 255, 0.98);
      color: var(--ink);
      border-bottom: 1px solid #d9e6f0;
      padding: 14px 12px;
      font-size: 13px;
      text-align: center;
      white-space: nowrap;
    }

    thead tr:first-child th {
      background: linear-gradient(180deg, rgba(232,245,248,0.98), rgba(246,251,255,0.98));
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: #0f5d63;
    }

    thead th.project-col,
    tbody td.project-col {
      position: sticky;
      left: 0;
      z-index: 4;
      min-width: 250px;
      text-align: left;
      background: rgba(252, 254, 255, 0.98);
      box-shadow: 8px 0 18px rgba(203, 220, 233, 0.24);
    }

    tbody td {
      padding: 12px;
      border-bottom: 1px solid #e4eef6;
      background: rgba(255,255,255,0.74);
    }

    tbody tr:hover td {
      background: rgba(246, 252, 255, 0.98);
    }

    .project-button {
      appearance: none;
      border: 0;
      background: transparent;
      padding: 0;
      margin: 0;
      cursor: pointer;
      text-align: left;
      width: 100%;
      color: inherit;
    }

    .project-name {
      font-size: 15px;
      font-weight: 700;
      line-height: 1.4;
    }

    .project-meta {
      margin-top: 6px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      font-size: 12px;
      color: var(--muted);
    }

    .bar-cell {
      position: relative;
      min-width: 110px;
      height: 44px;
      display: flex;
      align-items: center;
      justify-content: flex-end;
      padding: 0 12px;
      border-radius: 14px;
      background: #f6fbff;
      overflow: hidden;
      border: 1px solid #e6eff7;
    }

    .bar-fill {
      position: absolute;
      left: 0;
      top: 0;
      bottom: 0;
      border-radius: 14px;
      opacity: 0.96;
    }

    .tone-total .bar-fill {
      background: linear-gradient(90deg, rgba(15,139,141,0.94), rgba(20,184,166,0.8));
    }

    .tone-detail .bar-fill {
      background: linear-gradient(90deg, rgba(37,99,235,0.88), rgba(96,165,250,0.76));
    }

    .tone-material .bar-fill {
      background: linear-gradient(90deg, rgba(245,158,11,0.88), rgba(251,191,36,0.76));
    }

    .bar-value {
      position: relative;
      z-index: 1;
      font-variant-numeric: tabular-nums;
      font-size: 13px;
      font-weight: 700;
      color: #0f2b42;
    }

    .detail-card-list {
      display: grid;
      gap: 14px;
    }

    .project-detail-card {
      border-radius: 20px;
      background: var(--card-solid);
      border: 1px solid #e2edf6;
      box-shadow: 0 10px 26px rgba(22, 54, 83, 0.08);
      padding: 18px;
      scroll-margin-top: 16px;
    }

    .detail-card-head {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 12px;
      margin-bottom: 12px;
    }

    .detail-card-head h3 {
      margin: 0;
      font-size: 22px;
    }

    .detail-card-head p {
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 13px;
    }

    .detail-hero {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
      padding: 24px;
      scroll-margin-top: 16px;
    }

    .detail-hero h2 {
      margin: 8px 0 0;
      font-size: clamp(24px, 3.6vw, 34px);
      line-height: 1.15;
    }

    .detail-hero p {
      margin: 10px 0 0;
      color: var(--muted);
      font-size: 14px;
    }

    .chip-wrap {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      justify-content: flex-end;
    }

    .chip {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 10px 14px;
      border-radius: var(--radius-pill);
      background: #ebfbf6;
      border: 1px solid #c5eee2;
      color: #0f766e;
      font-size: 13px;
      font-weight: 700;
      white-space: nowrap;
    }

    .chip.alt {
      background: #eff6ff;
      border-color: #d8e6ff;
      color: #1d4ed8;
    }

    .cards-shell,
    .visual-shell,
    .materials-shell {
      padding: 0 24px 24px;
    }

    .cards-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
      gap: 12px;
    }

    .summary-card {
      border-radius: var(--radius-lg);
      background: var(--card-solid);
      border: 1px solid #e2edf6;
      padding: 16px;
      min-height: 118px;
      box-shadow: 0 10px 26px rgba(22, 54, 83, 0.08);
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      position: relative;
      overflow: hidden;
    }

    .summary-card::before {
      content: "";
      position: absolute;
      left: 0;
      top: 0;
      width: 100%;
      height: 4px;
      background: linear-gradient(90deg, #0f8b8d, #2563eb, #f59e0b);
    }

    .summary-card .label {
      margin: 0;
      font-size: 13px;
      color: var(--muted);
    }

    .summary-card .value {
      margin: 10px 0 8px;
      font-size: clamp(22px, 3vw, 32px);
      line-height: 1;
      font-weight: 800;
      letter-spacing: 0.02em;
    }

    .summary-card .sub {
      font-size: 12px;
      color: var(--muted);
    }

    .visual-grid {
      display: grid;
      grid-template-columns: minmax(340px, 430px) minmax(0, 1fr);
      gap: 14px;
    }

    .visual-card {
      border-radius: var(--radius-lg);
      background: var(--card-solid);
      border: 1px solid #e2edf6;
      padding: 18px;
      box-shadow: 0 10px 26px rgba(22, 54, 83, 0.08);
    }

    .panel-title {
      margin: 0 0 14px;
      font-size: 16px;
      font-weight: 800;
    }

    .donut-panel {
      display: grid;
      grid-template-columns: 180px 1fr;
      gap: 18px;
      align-items: start;
    }

    .donut-shell {
      display: grid;
      place-items: center;
    }

    .donut {
      width: 180px;
      height: 180px;
      border-radius: 50%;
      position: relative;
      background: conic-gradient(#d8e1ea 0% 100%);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.7), 0 16px 36px rgba(18, 50, 74, 0.12);
    }

    .donut::after {
      content: "";
      position: absolute;
      inset: 24px;
      border-radius: 50%;
      background: rgba(255,255,255,.97);
      box-shadow: inset 0 0 0 1px #edf2f7;
    }

    .donut-center {
      position: absolute;
      inset: 0;
      z-index: 1;
      display: grid;
      place-items: center;
      text-align: center;
      pointer-events: none;
      padding: 0 32px;
    }

    .donut-center .big {
      font-size: 20px;
      font-weight: 800;
      line-height: 1.08;
    }

    .donut-center .small {
      margin-top: 8px;
      font-size: 12px;
      color: var(--muted);
    }

    .legend-list {
      display: grid;
      gap: 8px;
      max-height: 200px;
      overflow: auto;
      padding-right: 6px;
    }

    .legend-item {
      display: grid;
      grid-template-columns: 14px 1fr auto auto;
      gap: 10px;
      align-items: center;
      padding: 8px 0;
      border-bottom: 1px dashed #e2ebf3;
      font-size: 13px;
    }

    .dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
    }

    .legend-amount {
      font-weight: 700;
      font-variant-numeric: tabular-nums;
    }

    .legend-pct {
      color: var(--muted);
      font-variant-numeric: tabular-nums;
    }

    .material-bar-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(88px, 1fr));
      gap: 12px;
      min-height: 300px;
      align-items: end;
    }

    .material-bar-item {
      display: grid;
      gap: 8px;
      justify-items: center;
    }

    .material-bar-value {
      font-size: 12px;
      color: var(--ink);
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      min-height: 32px;
      text-align: center;
    }

    .material-bar-track {
      width: 100%;
      max-width: 64px;
      height: 180px;
      display: flex;
      align-items: flex-end;
      border-radius: 18px 18px 8px 8px;
      background: linear-gradient(180deg, #f7fafc 0%, #eef3f8 100%);
      padding: 6px;
      border: 1px solid #e4edf6;
    }

    .material-bar-fill {
      width: 100%;
      border-radius: 12px;
      background: linear-gradient(180deg, #fbbf24 0%, #f59e0b 55%, #d97706 100%);
      box-shadow: 0 10px 20px rgba(245, 158, 11, 0.24);
      min-height: 6px;
    }

    .material-bar-label {
      font-size: 13px;
      font-weight: 700;
      text-align: center;
    }

    .materials-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
    }

    .material-card {
      border-radius: 18px;
      background: var(--card-solid);
      border: 1px solid #e6edf6;
      padding: 16px;
      box-shadow: 0 8px 20px rgba(18, 50, 74, 0.08);
      display: grid;
      gap: 10px;
    }

    .material-card h3 {
      margin: 0;
      font-size: 16px;
    }

    .material-row {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      color: var(--muted);
      font-size: 13px;
    }

    .material-row strong {
      color: var(--ink);
      font-variant-numeric: tabular-nums;
    }

    .empty-tip {
      color: var(--muted);
      font-size: 14px;
      padding: 10px 0;
    }

    @media (max-width: 1180px) {
      .visual-grid {
        grid-template-columns: 1fr;
      }

      .donut-panel {
        grid-template-columns: 1fr;
      }
    }

    @media (max-width: 980px) {
      .container {
        padding-left: 14px;
        padding-right: 14px;
      }

      .layout {
        grid-template-columns: 1fr;
      }

      .toc {
        position: static;
        max-height: none;
      }

      .toc-links {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .hero,
      .detail-hero {
        flex-direction: column;
      }

      .hero-stats {
        width: 100%;
        min-width: 0;
      }

      .section-head,
      .overview-shell,
      .cards-shell,
      .visual-shell,
      .materials-shell {
        padding-left: 16px;
        padding-right: 16px;
      }

      .chip-wrap {
        justify-content: flex-start;
      }

      .overview-table {
        min-width: 1320px;
      }
    }
  </style>
</head>
<body>
  <main class="container">
    <div class="layout">
      <aside class="panel toc">
        <h3 class="toc-title">项目目录 TOC</h3>
        <div class="toc-links" id="tocLinks"></div>
      </aside>
      <div class="content-col">
        <section class="panel hero">
          <div>
            <p class="eyebrow">预算汇总看板</p>
            <h1>主材总汇总总览与项目详情</h1>
            <p id="metaInfo"></p>
          </div>
          <div class="hero-stats">
            <article class="hero-stat">
              <span>项目数量</span>
              <strong id="projectCountChip">0</strong>
            </article>
            <article class="hero-stat">
              <span>全项目总额</span>
              <strong id="overallTotalChip">0</strong>
            </article>
          </div>
        </section>

        <section class="panel">
          <div class="section-head">
            <div>
              <div class="section-kicker">1. 全项目看板</div>
              <h2>项目总览数据条（万元）</h2>
              <p>纵向列出所有项目，横向固定展示汇总金额和五类金额（单位：万元）。点击项目名称后跳转到下方详情。</p>
            </div>
          </div>
          <div class="overview-shell">
            <div class="overview-table" id="overviewTable"></div>
          </div>
        </section>

        <section class="panel cards-shell" id="projectDetailSection">
          <div class="section-head" style="padding-top: 0;">
            <div>
              <div class="section-kicker">2. 项目详情</div>
              <h2>项目卡片清单（一行一个项目）</h2>
              <p>每个项目一个独立详情卡片，内容包含金额卡片、结构占比、主材金额条形图与主材数量/金额卡片。</p>
            </div>
          </div>
          <div class="detail-card-list" id="detailCards"></div>
        </section>
      </div>
    </div>
  </main>

  <script>
    const DASHBOARD_DATA = __DASHBOARD_DATA__;

    const PALETTE = [
      '#0f8b8d', '#2563eb', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6',
      '#14b8a6', '#fb7185', '#22c55e', '#f97316', '#64748b', '#0891b2'
    ];

    const totalCol = DASHBOARD_DATA.columns.total_col;
    const detailCols = DASHBOARD_DATA.columns.detail_cols;
    const materialCols = DASHBOARD_DATA.columns.material_cols;
    const cardCols = DASHBOARD_DATA.columns.card_cols;
    const overviewCols = DASHBOARD_DATA.columns.overview_cols;
    const columnMax = DASHBOARD_DATA.columns.column_max || {};
    const projects = DASHBOARD_DATA.projects || [];

    function esc(text) {
      return String(text || '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
    }

    function num(value) {
      return Number(value || 0);
    }

    function fmt(value) {
      return new Intl.NumberFormat('zh-CN', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2,
      }).format(num(value));
    }

    function fmtWan(value) {
      const v = num(value);
      const wan = v / 10000;
      return new Intl.NumberFormat('zh-CN', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2,
      }).format(wan) + '万';
    }

    function pct(part, total) {
      const p = num(part);
      const t = num(total);
      if (!t || t <= 0) return '0.0%';
      return ((p / t) * 100).toFixed(1) + '%';
    }

    function projectDisplayName(project) {
      const code = project['项目编码'] ? `【${project['项目编码']}】` : '';
      return `${code}${project['项目名称'] || '未命名项目'}`;
    }

    function barWidth(value, maxValue) {
      const v = num(value);
      const m = num(maxValue);
      if (!v || !m || m <= 0) return 0;
      return Math.min(100, Math.max(6, (v / m) * 100));
    }

    function columnTone(column) {
      if (column === totalCol) return 'total';
      if (detailCols.includes(column)) return 'detail';
      return 'material';
    }

    function detailId(project) {
      return `detail-project-${project.id}`;
    }

    function initHeader() {
      const meta = DASHBOARD_DATA.meta;
      const text = `数据源: ${meta.source_file} | 项目数: ${meta.project_count} | 生成时间: ${meta.generated_at}`;
      document.getElementById('metaInfo').textContent = text;
      document.getElementById('projectCountChip').textContent = String(meta.project_count || 0);
      document.getElementById('overallTotalChip').textContent = fmtWan(meta.overall_total || 0);
    }

    function renderTOC() {
      const html = projects.map((project) => {
        return `<a class="toc-link" href="#${detailId(project)}">${esc(projectDisplayName(project))}</a>`;
      }).join('');
      document.getElementById('tocLinks').innerHTML = html;
    }

    function renderOverviewTable() {
      const rowsHtml = projects.map((project) => {
        const cells = overviewCols.map((column) => {
          const value = num(project.metrics[column]);
          const width = barWidth(value, columnMax[column]);
          const tone = columnTone(column);
          return `
            <td>
              <div class="bar-cell tone-${tone}">
                <span class="bar-fill" style="width:${width}%"></span>
                <span class="bar-value">${esc(fmtWan(value))}</span>
              </div>
            </td>
          `;
        }).join('');

        const sequence = project['序号'] ? `序号 ${esc(project['序号'])}` : '未标序号';
        const currency = project['币种'] || '未标币种';

        return `
          <tr>
            <td class="project-col">
              <a class="project-button" href="#${detailId(project)}">
                <div class="project-name">${esc(projectDisplayName(project))}</div>
                <div class="project-meta">
                  <span>${sequence}</span>
                  <span>${esc(currency)}</span>
                </div>
              </a>
            </td>
            ${cells}
          </tr>
        `;
      }).join('');

      const headerCells = overviewCols.map((column) => `<th>${esc(column)}</th>`).join('');

      document.getElementById('overviewTable').innerHTML = `
        <table>
          <thead>
            <tr>
              <th class="project-col" rowspan="2">项目名称</th>
              <th colspan="1">汇总金额（万）</th>
              <th colspan="${detailCols.length}">五类金额（万）</th>
            </tr>
            <tr>
              ${headerCells}
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      `;
    }

    function renderAmountCardsForProject(project) {
      const totalValue = num(project.metrics[totalCol]);
      const allTotal = num(DASHBOARD_DATA.meta.overall_total);
      const detailSum = detailCols.reduce((sum, column) => sum + num(project.metrics[column]), 0);

      const defs = cardCols.map((column) => {
        const value = num(project.metrics[column]);
        const sub = column === totalCol
          ? `占全项目总额 ${pct(value, allTotal)}`
          : `占项目总额 ${pct(value, totalValue)}`;
        return {
          label: column,
          value,
          sub,
        };
      });

      defs.push({
        label: '五类金额合计',
        value: detailSum,
        sub: `五类合计 / 项目总额 ${pct(detailSum, totalValue)}`,
      });

      return defs.map((item) => `
        <article class="summary-card">
          <p class="label">${esc(item.label)}</p>
          <div class="value">${esc(fmtWan(item.value))}</div>
          <div class="sub">${esc(item.sub)}</div>
        </article>
      `).join('');
    }

    function renderDonutAndLegendForProject(project) {
      const totalValue = num(project.metrics[totalCol]);
      const parts = detailCols
        .map((column, index) => ({
          label: column,
          value: num(project.metrics[column]),
          color: PALETTE[index % PALETTE.length],
        }))
        .filter((item) => item.value > 0);

      if (!parts.length || totalValue <= 0) {
        return `
          <div class="donut-panel">
            <div class="donut-shell">
              <div class="donut" style="background:conic-gradient(#d8e1ea 0% 100%)">
                <div class="donut-center">
                  <div class="big">${esc(fmtWan(totalValue))}</div>
                  <div class="small">项目总额</div>
                </div>
              </div>
            </div>
            <div class="legend-list"><div class="empty-tip">当前项目暂无五类结构金额数据</div></div>
          </div>
        `;
      }

      let cursor = 0;
      const segments = [];
      parts.forEach((part) => {
        const ratio = (part.value / totalValue) * 100;
        const start = cursor;
        const end = Math.min(100, cursor + ratio);
        segments.push(`${part.color} ${start}% ${end}%`);
        cursor = end;
      });

      if (cursor < 100) {
        segments.push(`#d8e1ea ${cursor}% 100%`);
      }

      const legendHtml = parts
        .sort((a, b) => b.value - a.value)
        .map((part) => `
          <div class="legend-item">
            <span class="dot" style="background:${part.color}"></span>
            <span>${esc(part.label)}</span>
            <span class="legend-amount">${esc(fmtWan(part.value))}</span>
            <span class="legend-pct">${esc(pct(part.value, totalValue))}</span>
          </div>
        `)
        .join('');

      return `
        <div class="donut-panel">
          <div class="donut-shell">
            <div class="donut" style="background:conic-gradient(${segments.join(', ')})">
              <div class="donut-center">
                <div class="big">${esc(fmtWan(totalValue))}</div>
                <div class="small">项目总额</div>
              </div>
            </div>
          </div>
          <div class="legend-list">${legendHtml}</div>
        </div>
      `;
    }

    function materialAmountFromProject(project, materialName) {
      const metricValue = num(project.metrics[materialName]);
      if (metricValue > 0) return metricValue;
      const material = (project.materials || []).find((item) => item.name === materialName);
      return num(material ? material.amount : 0);
    }

    function renderMaterialBarsForProject(project) {
      const barItems = materialCols.map((label) => ({
        label,
        value: materialAmountFromProject(project, label),
      }));
      const maxValue = barItems.length ? Math.max(...barItems.map((item) => item.value)) : 0;
      const hasData = barItems.some((item) => item.value > 0);

      if (!hasData) {
        return '<div class="empty-tip">当前项目暂无主材金额数据</div>';
      }

      return barItems.map((item) => {
        const height = maxValue > 0 ? Math.max(6, (item.value / maxValue) * 100) : 6;
        return `
          <div class="material-bar-item">
            <div class="material-bar-value">${esc(fmtWan(item.value))}</div>
            <div class="material-bar-track">
              <div class="material-bar-fill" style="height:${height}%"></div>
            </div>
            <div class="material-bar-label">${esc(item.label)}</div>
          </div>
        `;
      }).join('');
    }

    function renderMaterialCardsForProject(project) {
      const materials = (project.materials || []).map((item) => ({ ...item }));
      const hasData = materials.some((item) => num(item.amount) > 0 || num(item.quantity) > 0);

      if (!materials.length || !hasData) {
        return '<div class="empty-tip">当前项目暂无材料数量/金额数据，请确认总汇总文件中包含“材料汇总”工作表并已汇总到对应项目。</div>';
      }

      return materials.map((item) => {
        const unit = item.unit || '';
        return `
          <article class="material-card">
            <h3>${esc(item.name)}</h3>
            <div class="material-row"><span>数量</span><strong>${esc(fmt(item.quantity))} ${esc(unit)}</strong></div>
            <div class="material-row"><span>金额</span><strong>${esc(fmtWan(item.amount))}</strong></div>
            <div class="material-row"><span>平均单价</span><strong>${esc(fmt(item.avg_price))}</strong></div>
          </article>
        `;
      }).join('');
    }

    function renderDetailCards() {
      const allTotal = num(DASHBOARD_DATA.meta.overall_total);
      const html = projects.map((project) => {
        const totalValue = num(project.metrics[totalCol]);
        const code = project['项目编码'] ? `项目编码: ${project['项目编码']}` : '项目编码: 未填写';
        const currency = project['币种'] ? `币种: ${project['币种']}` : '币种: 未填写';
        const sequence = project['序号'] ? `序号: ${project['序号']}` : '序号: 未填写';
        return `
          <article class="project-detail-card" id="${detailId(project)}">
            <div class="detail-card-head">
              <div>
                <h3>${esc(projectDisplayName(project))}</h3>
                <p>${esc(`${code} | ${currency} | ${sequence}`)}</p>
              </div>
              <div class="chip-wrap">
                <span class="chip">项目总额 ${esc(fmtWan(totalValue))}</span>
                <span class="chip alt">占全项目 ${esc(pct(totalValue, allTotal))}</span>
              </div>
            </div>
            <div class="cards-grid">${renderAmountCardsForProject(project)}</div>
            <div class="visual-grid" style="margin-top: 12px;">
              <article class="visual-card">
                <h3 class="panel-title">五类结构金额占比</h3>
                ${renderDonutAndLegendForProject(project)}
              </article>
              <article class="visual-card">
                <h3 class="panel-title">主材金额条形图</h3>
                <div class="material-bar-grid">${renderMaterialBarsForProject(project)}</div>
              </article>
            </div>
            <div style="margin-top: 12px;">
              <h3 class="panel-title">主材数量与金额卡片</h3>
              <div class="materials-grid">${renderMaterialCardsForProject(project)}</div>
            </div>
          </article>
        `;
      }).join('');
      document.getElementById('detailCards').innerHTML = html;
    }

    function bootstrap() {
      if (!projects.length) {
        document.body.innerHTML = '<main class="container"><section class="panel"><div class="section-head"><div><h2>无可展示项目数据</h2></div></div></section></main>';
        return;
      }

      initHeader();
      renderTOC();
      renderOverviewTable();
      renderDetailCards();
    }

    bootstrap();
  </script>
</body>
</html>
"""

    return template.replace("__DASHBOARD_DATA__", data_json)


def parse_args():
    parser = argparse.ArgumentParser(description="生成总汇总项目看板 HTML")
    parser.add_argument(
        "--run-source",
        action="store_true",
        help="先执行 11.8 主脚本，再读取最新总汇总文件",
    )
    parser.add_argument(
        "--source-script",
        default=str(DEFAULT_SOURCE_SCRIPT),
        help="11.8 主脚本路径",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="总汇总 Excel 所在目录",
    )
    parser.add_argument(
        "--summary-file",
        default="",
        help="指定总汇总 Excel 路径（不填则自动找最新）",
    )
    parser.add_argument(
        "--output-html",
        default="",
        help="输出HTML路径（不填默认写到 output-dir/总汇总项目看板.html）",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    source_script = Path(args.source_script).expanduser()
    output_dir = Path(args.output_dir).expanduser()

    if args.run_source:
        print(f"[INFO] 先执行源脚本: {source_script}")
        run_source_script(source_script)

    if args.summary_file:
        summary_file = Path(args.summary_file).expanduser()
        if not summary_file.exists():
            raise FileNotFoundError(f"指定的总汇总文件不存在: {summary_file}")
    else:
        summary_file = find_latest_summary_file(output_dir)

    payload = build_payload(summary_file)
    html = build_html(payload)

    if args.output_html:
        output_html = Path(args.output_html).expanduser()
    else:
        output_html = output_dir / DEFAULT_DASHBOARD_NAME

    output_html.parent.mkdir(parents=True, exist_ok=True)
    output_html.write_text(html, encoding="utf-8")

    print("=" * 64)
    print(f"[OK] 看板生成完成: {output_html}")
    print(f"[INFO] 数据源: {summary_file}")
    print("=" * 64)


if __name__ == "__main__":
    main()