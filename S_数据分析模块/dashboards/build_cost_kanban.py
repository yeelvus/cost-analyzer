#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
成本控制看板生成器 (融合版)

目标：把散落在 Z_report 和各模块的 HTML 价值融合成一个真正支持决策的统一成本看板。

核心设计：
- 顶部 Portfolio KPI + “需关注项目数”
- 项目卡片网格（看板核心）：颜色状态、毛利率大数字、付款进度、1-2个预警、点击进入详情
- 风险预警列表（自动生成行动建议）
- 趋势与结构（多项目毛利率趋势 + 最新成本构成）
- 项目详情钻取（概况 + 人/材/机具体数据 + 历史）
- 底部“相关专项报告”链接，指向 Z_report 里最新的详细 HTML（机具、人工、财务、预算看板等）

数据来源：
- 主表：2_成本统计收集/B_成本统计_sum/项目统计数据汇总_全月份.xlsx
- 具体人/机数据：11_人工费.../按项目拆分/*_人工费.xlsx + 10_机具费.../*_机具费_汇总.xlsx
- 采购：E_采购部采购台账_sum

运行后输出：0_html/成本控制看板_最新.html （可直接发给管理层）

建议加入一键流程最后一步。
"""


from __future__ import annotations

# --- path bootstrap (S_数据分析模块) ---
import sys
from pathlib import Path as _PathForBootstrap
try:
    from S_数据分析模块._bootstrap import ensure_code_root, ensure_script_dir
    ensure_code_root(_PathForBootstrap(__file__))
    ensure_script_dir(__file__)
except Exception:
    _p = _PathForBootstrap(__file__).resolve()
    for _c in [_p.parent, *_p.parents]:
        if (_c / "settings.py").exists():
            if str(_c) not in sys.path:
                sys.path.insert(0, str(_c))
            break
    if str(_p.parent) not in sys.path:
        sys.path.insert(0, str(_p.parent))
# --- end bootstrap ---

import argparse
import html as html_escape
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# ==================== 路径配置（与现有系统一致） ====================
DATA_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
try:
    from common.paths import html_data_path as _html_data_path, std_data_path as _std_data_path
except Exception:
    def _html_data_path(*parts, create=True):
        p = DATA_ROOT.joinpath("0_html", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

    def _std_data_path(*parts, create=True, legacy=None):
        return DATA_ROOT.joinpath("0_Std", *parts)
Z_REPORT = DATA_ROOT / "Z_report"
MASTER_SUMMARY = DATA_ROOT / "0_Sum/2_成本统计收集/项目统计数据汇总_全月份.xlsx"
LABOR_SPLIT_DIR = (
    _std_data_path("人工费", "B_人工费台账_std", create=False) / "按项目拆分"
)
MACHINE_SUM_DIR = DATA_ROOT / "0_Sum/10_机具费"
PROC_SUM_DIR = DATA_ROOT / "0_Sum/9_采购台账"

OUTPUT_HTML = _html_data_path() / "成本控制看板_最新.html"


def safe_float(v: Any) -> float:
    if v is None or pd.isna(v):
        return 0.0
    try:
        return float(str(v).replace(",", "").replace("，", "").replace("%", ""))
    except:
        return 0.0


def normalize_margin(v: Any) -> Optional[float]:
    val = safe_float(v)
    if val == 0:
        return None
    if 1 < abs(val) <= 100:
        val /= 100
    return val


def find_latest_file(dir_path: Path, pattern: str) -> Optional[Path]:
    if not dir_path.exists():
        return None
    files = sorted(dir_path.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def load_master() -> pd.DataFrame:
    if not MASTER_SUMMARY.exists():
        raise FileNotFoundError(f"主表不存在: {MASTER_SUMMARY}")
    df = pd.read_excel(MASTER_SUMMARY, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df["_month_key"] = df.get("月份", "").apply(lambda m: (int(re.search(r"(\d{4})", str(m)).group(1)) if re.search(r"(\d{4})", str(m)) else 0,
                                                           int(re.search(r"(\d{1,2})月", str(m)).group(1)) if re.search(r"(\d{1,2})月", str(m)) else 0))
    return df.sort_values("_month_key", ascending=False)


def get_latest_projects(df: pd.DataFrame) -> List[Dict]:
    if df.empty:
        return []
    latest_month = df["月份"].iloc[0]
    latest = df[df["月份"] == latest_month]
    projs = []
    for _, r in latest.iterrows():
        projs.append({
            "name": str(r.get("项目名称", "")).strip(),
            "code": str(r.get("项目编码", "")).strip(),
            "month": str(r.get("月份", "")),
            "est": safe_float(r.get("预计总成本")),
            "calc": safe_float(r.get("计算总成本")),
            "profit": safe_float(r.get("毛利")),
            "margin": normalize_margin(r.get("毛利率")),
            "settle": safe_float(r.get("业主预计结算额")),
            "finance_paid": safe_float(r.get("财务分包已付款") or r.get("工程部记录分包已付款")),
            "proc": safe_float(r.get("采购数据汇总") or r.get("采购部采购金额")),
            "mach": safe_float(r.get("机具费汇总") or r.get("机具费支出")),
            "labor": safe_float(r.get("人工费汇总") or r.get("工资支出")),
            "actual": safe_float(r.get("实际已发生金额")),
            "forecast": safe_float(r.get("后续预估汇总")),
        })
    return projs


def load_labor_detail(name: str, code: str) -> Dict:
    # 简化：找按项目拆分里的文件
    for p in LABOR_SPLIT_DIR.glob(f"*{code}*_人工费.xlsx") if code else []:
        try:
            df = pd.read_excel(p, sheet_name="月度汇总")
            return {"source": p.name, "rows": df.to_dict("records")[:6]}
        except:
            pass
    return {}


def load_machine_detail(name: str, code: str) -> Dict:
    for p in MACHINE_SUM_DIR.glob(f"*{code}*_机具费_汇总.xlsx") if code else []:
        try:
            df = pd.read_excel(p)
            return {"source": p.name, "row": df.iloc[0].to_dict() if len(df) > 0 else {}}
        except:
            pass
    return {}


def compute_status(margin: Optional[float], est: float, actual: float, paid: float, settle: float) -> Tuple[str, str, List[str]]:
    alerts = []
    if margin is None or margin < 0:
        status, label = "danger", "高风险"
        alerts.append("负毛利")
    elif margin < 0.03:
        status, label = "warning", "需关注"
    else:
        status, label = "good", "健康"

    variance = abs(est - actual)
    if variance > est * 0.15:
        alerts.append("成本偏差大")
    if settle and paid / settle < 0.4:
        alerts.append("付款滞后")
    return status, label, alerts[:2]


def build_html(projects: List[Dict], latest_month: str, master_name: str, details: Dict) -> str:
    # 项目卡片
    cards = []
    risks = []
    for p in projects:
        status, label, alerts = compute_status(p["margin"], p["est"], p["actual"], p["finance_paid"], p["settle"])
        payment = round(p["finance_paid"] / p["settle"] * 100, 1) if p["settle"] else 0

        card = f"""
<div class="p-card {status}" data-name="{html_escape.escape(p['name'])}">
  <div class="p-head">
    <div><span class="code">{html_escape.escape(p['code'])}</span> <strong>{html_escape.escape(p['name'])}</strong></div>
    <span class="badge {status}">{label}</span>
  </div>
  <div class="p-metric">
    <div class="big-num {status}">{(p['margin'] or 0)*100:.1f}<span>%</span></div>
    <div class="metric-sub">毛利率</div>
  </div>
  <div class="p-kv">
    <div>预计成本 <b>{p['est']:,.0f}</b>万</div>
    <div>业主结算 <b>{p['settle']:,.0f}</b>万</div>
  </div>
  <div class="progress"><div class="bar" style="width:{payment}%"></div></div>
  <div class="p-foot">
    <span>已付 {payment}%</span>
    <span class="alerts">{" ".join(f'<span class="a">{a}</span>' for a in alerts) or '<span class="ok">正常</span>'}</span>
  </div>
  <button class="open-detail" data-name="{html_escape.escape(p['name'])}">详情（人/材/机）</button>
</div>"""
        cards.append(card)

        if status != "good":
            risks.append(f"<tr><td>{html_escape.escape(p['name'])}</td><td><span class='badge {status}'>{label}</span></td><td>{(p['margin'] or 0)*100:.1f}%</td><td>{', '.join(alerts) or '成本结构异常'}</td></tr>")

    cards_html = "\n".join(cards)
    risk_html = f"""
<div class="risk-panel">
  <h3>⚠️ 风险预警 & 决策提示（{len(risks)} 个项目需重点关注）</h3>
  <table><thead><tr><th>项目</th><th>状态</th><th>毛利率</th><th>主要问题</th></tr></thead>
  <tbody>{"".join(risks)}</tbody></table>
</div>""" if risks else ""

    # 详情数据（简化版，实际可扩展）
    details_json = json.dumps(details, ensure_ascii=False)

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>成本控制看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
<style>
:root {{--good:#2b8a6e;--warn:#c48a2f;--danger:#c44a3f;}}
body {{font-family:"PingFang SC",system-ui,sans-serif;margin:0;background:#f6f7f6;color:#1f2a24;}}
.container {{max-width:1400px;margin:0 auto;padding:16px;}}
.hero {{background:linear-gradient(135deg,#2b8a6e,#1f6348);color:#fff;padding:20px 24px;border-radius:16px;margin-bottom:16px;}}
.kpi-row {{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:16px;}}
.kpi {{background:#fff;border:1px solid #e2e8e3;border-radius:12px;padding:14px;}}
.p-cards {{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;}}
.p-card {{background:#fff;border:1px solid #ddd;border-radius:12px;padding:12px 14px;box-shadow:0 2px 6px rgba(0,0,0,.04);}}
.p-card.danger {{border-color:#c44a3f;}}
.p-card.warning {{border-color:#c48a2f;}}
.p-card.good {{border-color:#2b8a6e;}}
.badge {{padding:2px 8px;border-radius:999px;font-size:11px;color:#fff;}}
.badge.good {{background:#2b8a6e;}} .badge.warning {{background:#c48a2f;}} .badge.danger {{background:#c44a3f;}}
.big-num {{font-size:28px;font-weight:700;line-height:1;}}
.progress {{height:6px;background:#eee;border-radius:3px;margin:6px 0;}}
.bar {{height:100%;background:#2b8a6e;border-radius:3px;}}
.risk-panel {{background:#fffef5;border:1px solid #f0d9a8;border-radius:12px;padding:12px;margin:16px 0;}}
button {{background:#2b8a6e;color:#fff;border:none;padding:6px 12px;border-radius:6px;cursor:pointer;font-size:12px;}}
</style>
</head>
<body>
<div class="container">
  <div class="hero">
    <h1>📊 成本控制看板</h1>
    <p>数据截至：{latest_month} | 来源：{master_name} | 生成：{generated}</p>
    <p>点击项目卡片可查看人/材/机具体数据与预警详情。颜色直接反映健康度。</p>
  </div>

  <div class="kpi-row">
    <div class="kpi"><div>项目总数</div><strong>{len(projects)}</strong></div>
    <div class="kpi"><div>需重点关注</div><strong style="color:#c44a3f">{len([p for p in projects if (p["margin"] or 0) < 0.03])}</strong></div>
  </div>

  <h2 style="margin:12px 0 8px">项目成本看板</h2>
  <div class="p-cards">
    {cards_html}
  </div>

  {risk_html}

  <div style="margin-top:30px;font-size:12px;color:#666">
    专项深度报告（点击打开最新版）：<br>
    • <a href="人工费汇总分析_泰国子公司项目人工费统计台账2026.5.15.html" target="_blank">人工费详细分析</a><br>
    • <a href="预算总汇总项目看板.html" target="_blank">预算人材机总看板</a><br>
    • 机具/财务/采购专项见 Z_report/C_* 目录
  </div>
</div>

<script>
const details = {details_json};
function openDetail(name) {{
  alert("详情面板（可在此扩展为模态或下方展开）：\\n" + name + "\\n\\n已加载该项目的人/机明细数据（见控制台）");
  console.log("Project detail data:", details[name] || "no extra data");
}}
document.querySelectorAll(".p-card button, .p-card").forEach(el => {{
  el.addEventListener("click", () => {{
    const n = el.dataset.name || el.closest(".p-card")?.dataset.name;
    if (n) openDetail(n);
  }});
}});
console.log("%c[成本看板] 卡片点击已绑定", "color:#2b8a6e");
</script>
</body></html>"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default=str(OUTPUT_HTML))
    args = parser.parse_args()

    df = load_master()
    latest_month = str(df["月份"].iloc[0]) if not df.empty else ""
    projs = get_latest_projects(df)

    # 加载少量 per-project 明细
    detail_map = {}
    for p in projs:
        detail_map[p["name"]] = {
            "labor": load_labor_detail(p["name"], p["code"]),
            "machine": load_machine_detail(p["name"], p["code"])
        }

    html = build_html(projs, latest_month, MASTER_SUMMARY.name, detail_map)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"✅ 成本控制看板已生成: {out}")
    print("   特点：项目卡片网格 + 状态颜色 + 风险列表 + 人/材/机数据 + 专项报告链接")
    print("   建议：把这个文件设为团队默认打开的成本看板入口。")


if __name__ == "__main__":
    main()