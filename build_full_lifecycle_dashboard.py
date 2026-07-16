#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全过程多项目成本分析看板生成器（P0/P1）

数据源优先级：
1. 1_成本报公司/*成本统计表.xlsx  → 责任成本 / 已发生 / 预估 / 预计毛利 / 完成比例
2. 18_output/项目汇总索引.xlsx    → 台账实际（采购/机具/人工/合同）+ 预算人材机
3. 0_项目基本信息情况              → 状态、经理、工期
4. master 全月份（可选）          → 多月趋势

输出：
  Z_report/全过程成本分析看板_最新.html

用法：
  python build_full_lifecycle_dashboard.py
  python build_full_lifecycle_dashboard.py --month 2026年06月
"""

from __future__ import annotations

import argparse
import html as html_lib
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

try:
    from settings import DATA_ROOT, REPORT_ROOT
except Exception:
    DATA_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
    REPORT_ROOT = DATA_ROOT / "Z_report"

COMPANY_DIR = DATA_ROOT / "1_成本报公司"
INDEX_XLSX = DATA_ROOT / "18_output" / "项目汇总索引.xlsx"
PROJECT_INFO = DATA_ROOT / "0_项目基本信息情况" / "项目基本信息表.xlsx"
MASTER_XLSX = DATA_ROOT / "2_成本统计收集" / "B_成本统计_sum" / "项目统计数据汇总_全月份.xlsx"
OUTPUT_HTML = REPORT_ROOT / "全过程成本分析看板_最新.html"


# ───────────────────────── helpers ─────────────────────────

def sf(v: Any) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "").replace("%", "")
    if not s or s.lower() in {"nan", "none", "-"}:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def sn(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def norm_margin(v: Any) -> Optional[float]:
    x = sf(v)
    if x == 0:
        return None
    if abs(x) > 1.5:  # 百分数
        x /= 100.0
    return x


def esc(s: Any) -> str:
    return html_lib.escape("" if s is None else str(s))


def find_latest_company_report(month_hint: Optional[str] = None) -> Path:
    files = [
        p for p in COMPANY_DIR.glob("*泰国子公司各项目成本统计表*.xlsx")
        if not p.name.startswith("~$") and "对比" not in p.name and "内部" not in p.name
    ]
    if not files:
        raise FileNotFoundError(f"未找到报公司成本统计表: {COMPANY_DIR}")
    if month_hint:
        for p in files:
            if month_hint.replace("年", "").replace("月", "") in p.name.replace("年", "").replace("月", ""):
                return p
            # 2026年06月 / 2026年6月
            m = re.search(r"(\d{4})年0?(\d{1,2})月", month_hint)
            if m:
                y, mo = m.group(1), m.group(2)
                if y in p.name and (f"{int(mo):02d}月" in p.name or f"{int(mo)}月" in p.name):
                    return p
    return max(files, key=lambda p: p.stat().st_mtime)


def month_from_filename(path: Path) -> str:
    m = re.search(r"(\d{4})年0?(\d{1,2})月", path.name)
    if m:
        return f"{m.group(1)}年{int(m.group(2)):02d}月"
    return path.stem


# ───────────────────────── loaders ─────────────────────────

def load_company_report(path: Path) -> Tuple[List[Dict], List[Dict], str]:
    """解析报公司表：在建 + 已完工。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    month = month_from_filename(path)

    # ---- 在建 ----
    active: List[Dict] = []
    active_sheet = None
    for name in wb.sheetnames:
        if "在建" in name:
            active_sheet = wb[name]
            break
    if active_sheet is not None:
        # 找表头行（含「项目名称」）
        header_row = 2
        for r in range(1, 6):
            vals = [sn(active_sheet.cell(r, c).value) for c in range(1, 12)]
            if any("项目名称" in v for v in vals):
                header_row = r
                break
        headers = []
        for c in range(1, 25):
            h = sn(active_sheet.cell(header_row, c).value).replace("\n", "")
            headers.append(h or f"c{c}")
        # 处理重复「税金补差」
        seen = {}
        for i, h in enumerate(headers):
            if h in seen:
                headers[i] = f"{h}_{seen[h]}"
                seen[h] += 1
            else:
                seen[h] = 1

        def col(*keys: str) -> Optional[int]:
            for i, h in enumerate(headers):
                for k in keys:
                    if k in h:
                        return i
            return None

        idx = {
            "code": col("项目编码"),
            "name": col("项目名称"),
            "pm": col("项目经理"),
            "contract": col("业主合同额"),
            "resp": col("责任成本"),
            "bid_profit": col("投标毛利"),
            "bid_margin": col("含税金差（投标）", "含税金差"),
            "settle": col("业主预计结算额"),
            "est_cost": col("梳理成本"),
            "actual": col("已发生"),
            "forecast": col("预估发生"),
            "exp_profit": col("预估毛利"),
            "exp_margin": col("预计毛利率"),
            "progress": col("完成比例"),
            "area": col("建筑面积"),
            "note": col("备注"),
            "progress_note": col("进度概述"),
        }

        for r in range(header_row + 1, active_sheet.max_row + 1):
            name = sn(active_sheet.cell(r, (idx["name"] or 2) + 1).value)
            if not name or name in {"合计", "小计"} or name.startswith("一、") or name.startswith("二、"):
                continue
            code = sn(active_sheet.cell(r, (idx["code"] or 1) + 1).value) if idx["code"] is not None else ""
            if not code and not name:
                continue

            def g(key: str) -> Any:
                i = idx.get(key)
                if i is None:
                    return None
                return active_sheet.cell(r, i + 1).value

            contract = sf(g("contract"))
            resp = sf(g("resp"))
            settle = sf(g("settle"))
            est_cost = sf(g("est_cost"))
            actual = sf(g("actual"))
            forecast = sf(g("forecast"))
            exp_profit = sf(g("exp_profit"))
            bid_profit = sf(g("bid_profit"))
            exp_margin = norm_margin(g("exp_margin"))
            bid_margin = norm_margin(g("bid_margin"))
            progress = sf(g("progress"))
            if progress > 1.5:
                progress /= 100.0

            if exp_margin is None and settle and est_cost:
                exp_margin = (settle - est_cost) / settle if settle else None
            if bid_margin is None and contract and bid_profit:
                bid_margin = bid_profit / contract if contract else None

            consume = (actual / resp) if resp > 0 else None
            gap = (consume - progress) if (consume is not None and progress) else None
            erosion = None
            if bid_margin is not None and exp_margin is not None:
                erosion = bid_margin - exp_margin

            active.append({
                "bucket": "在建/过程",
                "code": code,
                "name": name,
                "pm": sn(g("pm")),
                "contract": contract,
                "resp_cost": resp,
                "bid_profit": bid_profit,
                "bid_margin": bid_margin,
                "settle": settle,
                "est_cost": est_cost,
                "actual": actual,
                "forecast": forecast,
                "exp_profit": exp_profit,
                "exp_margin": exp_margin,
                "progress": progress,
                "consume": consume,
                "cost_progress_gap": gap,
                "margin_erosion": erosion,
                "area": sf(g("area")),
                "note": sn(g("note")),
                "progress_note": sn(g("progress_note")),
            })

    # ---- 已完工 ----
    done: List[Dict] = []
    done_sheet = None
    for name in wb.sheetnames:
        if "已完工" in name or "完工" in name:
            done_sheet = wb[name]
            break
    if done_sheet is not None:
        header_row = 2
        for r in range(1, 6):
            vals = [sn(done_sheet.cell(r, c).value) for c in range(1, 8)]
            if any(v == "项目名称" or "项目名称" in v for v in vals):
                header_row = r
                break
        for r in range(header_row + 1, done_sheet.max_row + 1):
            seq = done_sheet.cell(r, 1).value
            name = sn(done_sheet.cell(r, 2).value)
            if not name:
                continue
            # 分组标题
            if isinstance(seq, str) and not str(seq).replace(".", "").isdigit():
                continue
            if name.startswith("与业主") or name.startswith("一") or name.startswith("二"):
                continue
            revenue = sf(done_sheet.cell(r, 3).value)
            cost_tax = sf(done_sheet.cell(r, 4).value)
            cost = sf(done_sheet.cell(r, 5).value)
            profit = sf(done_sheet.cell(r, 7).value)
            margin = norm_margin(done_sheet.cell(r, 8).value)
            if margin is None and revenue:
                margin = (revenue - (cost_tax or cost)) / revenue
            done.append({
                "bucket": "已完工",
                "code": "",
                "name": name,
                "pm": "",
                "contract": revenue,
                "resp_cost": cost_tax or cost,
                "bid_profit": 0.0,
                "bid_margin": None,
                "settle": revenue,
                "est_cost": cost_tax or cost,
                "actual": cost or cost_tax,
                "forecast": 0.0,
                "exp_profit": profit,
                "exp_margin": margin,
                "progress": 1.0,
                "consume": 1.0,
                "cost_progress_gap": 0.0,
                "margin_erosion": None,
                "area": 0.0,
                "note": "已完工",
                "progress_note": "",
            })

    wb.close()
    return active, done, month


def load_index() -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not INDEX_XLSX.exists():
        return pd.DataFrame(), pd.DataFrame()
    wan = pd.read_excel(INDEX_XLSX, sheet_name="万元汇总")
    try:
        bud = pd.read_excel(INDEX_XLSX, sheet_name="预算金额")
    except Exception:
        bud = pd.DataFrame()
    for df in (wan, bud):
        if not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            if "项目编码" in df.columns:
                df["项目编码"] = df["项目编码"].astype(str).str.strip()
    return wan, bud


def load_project_info() -> pd.DataFrame:
    if not PROJECT_INFO.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(PROJECT_INFO, sheet_name="在建项目")
    except Exception:
        df = pd.read_excel(PROJECT_INFO)
    df.columns = [str(c).strip() for c in df.columns]
    if "项目编码" in df.columns:
        df["项目编码"] = df["项目编码"].astype(str).str.strip()
    return df


def load_master_trend() -> Dict[str, List[Dict]]:
    """按项目名称返回月份序列（若 master 有有效毛利/成本）。"""
    if not MASTER_XLSX.exists():
        return {}
    df = pd.read_excel(MASTER_XLSX)
    df.columns = [str(c).strip() for c in df.columns]
    out: Dict[str, List[Dict]] = {}
    for _, r in df.iterrows():
        name = sn(r.get("项目名称"))
        if not name:
            continue
        item = {
            "month": sn(r.get("月份")),
            "est": sf(r.get("预计总成本")),
            "profit": sf(r.get("毛利")),
            "margin": norm_margin(r.get("毛利率")),
            "settle": sf(r.get("业主预计结算额")),
            "mat": sf(r.get("材料采购实际已发生")),
            "mach": sf(r.get("机具费支出") or r.get("机具费汇总")),
            "labor": sf(r.get("工资支出") or r.get("人工费汇总")),
        }
        out.setdefault(name, []).append(item)
    return out


def enrich(projects: List[Dict], wan: pd.DataFrame, bud: pd.DataFrame, info: pd.DataFrame) -> List[Dict]:
    wan_map = {}
    if not wan.empty and "项目编码" in wan.columns:
        for _, r in wan.iterrows():
            wan_map[sn(r.get("项目编码"))] = r
            wan_map[sn(r.get("项目名称"))] = r
    bud_map = {}
    if not bud.empty and "项目编码" in bud.columns:
        for _, r in bud.iterrows():
            bud_map[sn(r.get("项目编码"))] = r
            bud_map[sn(r.get("项目名称"))] = r
    info_map = {}
    if not info.empty and "项目编码" in info.columns:
        for _, r in info.iterrows():
            info_map[sn(r.get("项目编码"))] = r
            info_map[sn(r.get("项目名称"))] = r

    for p in projects:
        key_c, key_n = p["code"], p["name"]
        # 勿用 `a or b`：pandas Series 的真值判断会抛错
        w = wan_map.get(key_c) if key_c and key_c in wan_map else wan_map.get(key_n)
        b = bud_map.get(key_c) if key_c and key_c in bud_map else bud_map.get(key_n)
        i = info_map.get(key_c) if key_c and key_c in info_map else info_map.get(key_n)

        if w is not None:
            p["status"] = sn(w.get("项目状态")) or p.get("status", "")
            p["ledger_contract"] = sf(w.get("总包合同金额"))
            p["ledger_sub"] = sf(w.get("分包合同金额"))
            p["ledger_mat"] = sf(w.get("材料采购金额"))
            p["ledger_mach"] = sf(w.get("机具使用费"))
            p["ledger_labor"] = sf(w.get("人工费"))
            p["ledger_ready_mix"] = sf(w.get("商混站金额"))
            p["ledger_payable"] = sf(w.get("材料欠款金额"))
            # 若报公司无编码，从索引补
            if not p["code"] and sn(w.get("项目编码")):
                p["code"] = sn(w.get("项目编码"))
        else:
            p.setdefault("status", "")
            p["ledger_contract"] = 0.0
            p["ledger_sub"] = 0.0
            p["ledger_mat"] = 0.0
            p["ledger_mach"] = 0.0
            p["ledger_labor"] = 0.0
            p["ledger_ready_mix"] = 0.0
            p["ledger_payable"] = 0.0

        if b is not None:
            p["budget_currency"] = sn(b.get("币种")) or "—"
            p["budget_total"] = sf(b.get("标段人材机汇总"))
            p["budget_labor"] = sf(b.get("人_明细"))
            p["budget_mat"] = sf(b.get("材类明细"))
            p["budget_mach"] = sf(b.get("设机_明细"))
            p["budget_mgmt"] = sf(b.get("管_明细"))
            # 预算单位：原币；若 THB 则 /10000 成万泰铢展示
            if p["budget_currency"] == "THB":
                for k in ("budget_total", "budget_labor", "budget_mat", "budget_mach", "budget_mgmt"):
                    p[k] = p[k] / 10000.0
                p["budget_unit"] = "万泰铢"
            else:
                p["budget_unit"] = p["budget_currency"] or "原币"
        else:
            p["budget_currency"] = ""
            p["budget_total"] = 0.0
            p["budget_labor"] = 0.0
            p["budget_mat"] = 0.0
            p["budget_mach"] = 0.0
            p["budget_mgmt"] = 0.0
            p["budget_unit"] = ""

        if i is not None:
            p["status"] = p.get("status") or sn(i.get("项目状态"))
            if not p["pm"]:
                p["pm"] = sn(i.get("项目经理"))
            p["full_name"] = sn(i.get("项目全称"))
        else:
            p.setdefault("full_name", "")

        # 预警
        p["alerts"] = compute_alerts(p)
        p["risk"] = risk_level(p)
    return projects


def compute_alerts(p: Dict) -> List[str]:
    alerts = []
    m = p.get("exp_margin")
    if m is not None and m < 0:
        alerts.append("负毛利")
    elif m is not None and m < 0.05:
        alerts.append("低毛利")
    er = p.get("margin_erosion")
    if er is not None and er > 0.05 and (p.get("progress") or 0) >= 0.5:
        alerts.append(f"毛利侵蚀{er*100:.1f}pct")
    gap = p.get("cost_progress_gap")
    if gap is not None and gap > 0.15:
        alerts.append("成本超进度")
    if p.get("bucket") != "已完工":
        if not p.get("resp_cost"):
            alerts.append("缺责任成本")
        if not p.get("budget_total"):
            alerts.append("缺预算")
        ledgers = p.get("ledger_mat", 0) + p.get("ledger_mach", 0) + p.get("ledger_labor", 0)
        if ledgers <= 0 and (p.get("actual") or 0) > 0:
            alerts.append("台账未对齐")
    return alerts


def risk_level(p: Dict) -> str:
    alerts = p.get("alerts") or []
    if "负毛利" in alerts or any(a.startswith("毛利侵蚀") for a in alerts):
        return "danger"
    if any(a in alerts for a in ("低毛利", "成本超进度")):
        return "warning"
    if any(a.startswith("缺") or a == "台账未对齐" for a in alerts):
        return "muted"
    return "good"


# ───────────────────────── HTML ─────────────────────────

def build_html(projects: List[Dict], month: str, source_name: str, trend: Dict) -> str:
    active = [p for p in projects if p["bucket"] != "已完工"]
    done = [p for p in projects if p["bucket"] == "已完工"]

    def sumf(rows, key):
        return sum(sf(r.get(key)) for r in rows)

    # Portfolio KPIs — 在建为主
    k_contract = sumf(active, "contract")
    k_settle = sumf(active, "settle")
    k_resp = sumf(active, "resp_cost")
    k_actual = sumf(active, "actual")
    k_forecast = sumf(active, "forecast")
    k_est = sumf(active, "est_cost")
    k_exp_profit = sumf(active, "exp_profit")
    k_bid_profit = sumf(active, "bid_profit")
    n_risk = sum(1 for p in active if p["risk"] in ("danger", "warning"))
    portfolio_margin = (k_exp_profit / k_settle) if k_settle else None
    consume_all = (k_actual / k_resp) if k_resp else None

    # 表格行
    def row_html(p: Dict) -> str:
        m = p.get("exp_margin")
        bm = p.get("bid_margin")
        prog = p.get("progress") or 0
        cons = p.get("consume")
        risk = p["risk"]
        alerts = " ".join(f'<span class="tag">{esc(a)}</span>' for a in (p.get("alerts") or [])) or '<span class="tag ok">正常</span>'
        return f"""
<tr class="proj-row risk-{risk}" data-bucket="{esc(p['bucket'])}" data-pm="{esc(p.get('pm') or '')}"
    data-status="{esc(p.get('status') or '')}" data-risk="{risk}" data-name="{esc(p['name'])}"
    data-json='{esc(json.dumps(p, ensure_ascii=False))}'>
  <td><span class="dot {risk}"></span>{esc(p.get('code') or '—')}</td>
  <td class="name-cell"><strong>{esc(p['name'])}</strong><div class="sub">{esc(p.get('pm') or '')} · {esc(p.get('status') or p['bucket'])}</div></td>
  <td class="num">{p['contract']:,.0f}</td>
  <td class="num">{p['resp_cost']:,.0f}</td>
  <td class="num">{p['actual']:,.0f}</td>
  <td class="num">{p['forecast']:,.0f}</td>
  <td class="num">{p['est_cost']:,.0f}</td>
  <td class="num">{p['settle']:,.0f}</td>
  <td class="num margin-{risk}">{(f"{m*100:.1f}%" if m is not None else "—")}</td>
  <td class="num">{(f"{bm*100:.1f}%" if bm is not None else "—")}</td>
  <td>
    <div class="mini-bars">
      <div class="mb"><span>进度</span><div class="track"><i style="width:{min(prog,1)*100:.0f}%"></i></div><b>{prog*100:.0f}%</b></div>
      <div class="mb"><span>消耗</span><div class="track cons"><i style="width:{min(cons or 0,1.2)*100/1.2:.0f}%"></i></div><b>{(cons*100 if cons is not None else 0):.0f}%</b></div>
    </div>
  </td>
  <td>{alerts}</td>
</tr>"""

    rows_active = "\n".join(row_html(p) for p in sorted(active, key=lambda x: (x["risk"] != "danger", x["risk"] != "warning", -(x.get("exp_margin") or 0))))
    rows_done = "\n".join(row_html(p) for p in sorted(done, key=lambda x: -(x.get("exp_margin") or 0)))

    # chart data
    chart_names = [p["name"] for p in active]
    chart_margins = [round((p["exp_margin"] or 0) * 100, 2) for p in active]
    chart_bid = [round((p["bid_margin"] or 0) * 100, 2) if p.get("bid_margin") is not None else None for p in active]
    chart_consume = [round((p["consume"] or 0) * 100, 1) for p in active]
    chart_progress = [round((p["progress"] or 0) * 100, 1) for p in active]

    projects_json = json.dumps(projects, ensure_ascii=False)
    trend_json = json.dumps(trend, ensure_ascii=False)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>全过程多项目成本分析看板 · {esc(month)}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg:#f4f6f5; --card:#fff; --text:#1a2420; --muted:#6b7c74;
  --good:#1f8a65; --warn:#c4892a; --danger:#c2473d; --muted-risk:#8a9390;
  --line:#e3eae6; --hero1:#1b5e4a; --hero2:#2b8a6e;
}}
* {{box-sizing:border-box;}}
body {{margin:0;font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;background:var(--bg);color:var(--text);}}
.wrap {{max-width:1480px;margin:0 auto;padding:18px;}}
.hero {{
  background:linear-gradient(135deg,var(--hero1),var(--hero2));
  color:#fff;border-radius:18px;padding:22px 26px;margin-bottom:16px;
  box-shadow:0 8px 24px rgba(27,94,74,.18);
}}
.hero h1 {{margin:0 0 6px;font-size:24px;font-weight:700;}}
.hero p {{margin:4px 0;opacity:.92;font-size:13px;}}
.chain {{
  display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;
}}
.chain span {{
  background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.22);
  border-radius:999px;padding:4px 12px;font-size:12px;
}}
.kpis {{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:16px;}}
.kpi {{
  background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px 14px;
}}
.kpi .l {{font-size:12px;color:var(--muted);}}
.kpi .v {{font-size:20px;font-weight:700;margin-top:4px;}}
.kpi .u {{font-size:11px;color:var(--muted);}}
.panel {{
  background:var(--card);border:1px solid var(--line);border-radius:14px;
  padding:14px 16px;margin-bottom:14px;
}}
.panel h2 {{margin:0 0 10px;font-size:16px;}}
.filters {{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px;}}
.filters select, .filters input {{
  border:1px solid var(--line);border-radius:8px;padding:6px 10px;font-size:13px;background:#fff;
}}
.filters .hint {{font-size:12px;color:var(--muted);margin-left:auto;}}
.charts {{display:grid;grid-template-columns:1.2fr 1fr;gap:12px;}}
@media (max-width:960px) {{.charts {{grid-template-columns:1fr;}}}}
.chart-box {{background:#fafbfa;border:1px solid var(--line);border-radius:12px;padding:10px;height:300px;}}
table {{width:100%;border-collapse:collapse;font-size:12.5px;}}
th, td {{padding:8px 6px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle;}}
th {{position:sticky;top:0;background:#f0f5f2;z-index:1;font-weight:600;color:#334;}}
.num {{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}}
.name-cell .sub {{font-size:11px;color:var(--muted);margin-top:2px;}}
.proj-row {{cursor:pointer;}}
.proj-row:hover {{background:#f3faf6;}}
.dot {{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;}}
.dot.good {{background:var(--good);}} .dot.warning {{background:var(--warn);}}
.dot.danger {{background:var(--danger);}} .dot.muted {{background:var(--muted-risk);}}
.margin-danger {{color:var(--danger);font-weight:700;}}
.margin-warning {{color:var(--warn);font-weight:600;}}
.margin-good {{color:var(--good);font-weight:600;}}
.tag {{
  display:inline-block;background:#eef2f0;color:#445;border-radius:6px;
  padding:1px 6px;margin:1px;font-size:11px;
}}
.tag.ok {{background:#e7f6ef;color:var(--good);}}
.mini-bars .mb {{display:flex;align-items:center;gap:4px;margin:2px 0;}}
.mini-bars span {{width:28px;color:var(--muted);font-size:11px;}}
.mini-bars b {{width:36px;text-align:right;font-size:11px;}}
.track {{flex:1;height:6px;background:#e8eeea;border-radius:3px;overflow:hidden;min-width:50px;}}
.track i {{display:block;height:100%;background:var(--good);}}
.track.cons i {{background:#3a7bd5;}}
.table-scroll {{max-height:560px;overflow:auto;border:1px solid var(--line);border-radius:10px;}}
.legend {{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin:8px 0;}}
.legend i {{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:4px;}}
/* drawer */
.mask {{
  position:fixed;inset:0;background:rgba(15,25,30,.48);display:none;z-index:40;
  align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(2px);
}}
.mask.show {{display:flex;}}
.drawer {{
  position:relative;width:min(720px,96vw);max-height:min(88vh,900px);background:#fff;
  border-radius:16px;box-shadow:0 20px 50px rgba(0,0,0,.22);
  overflow:auto;padding:20px 22px 28px;z-index:50;
  transform:translateY(12px) scale(.98);opacity:0;transition:transform .2s ease, opacity .2s ease;
}}
.mask.show .drawer {{transform:translateY(0) scale(1);opacity:1;}}
.drawer h3 {{margin:0 40px 4px 0;}}
.drawer .close {{
  position:sticky;top:0;float:right;z-index:2;
  border:none;background:#eef2f0;border-radius:8px;padding:6px 12px;cursor:pointer;font-size:13px;
}}
.drawer .close:hover {{background:#e2e8e5;}}
.lifecycle {{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:6px;margin:12px 0;
}}
.lifecycle .s {{
  background:#f6faf8;border:1px solid var(--line);border-radius:10px;padding:8px;text-align:center;
}}
.lifecycle .s .t {{font-size:11px;color:var(--muted);}}
.lifecycle .s .n {{font-size:14px;font-weight:700;margin-top:4px;}}
.kv {{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px;}}
.kv .item {{background:#fafbfa;border:1px solid var(--line);border-radius:10px;padding:8px 10px;}}
.kv .item .t {{font-size:11px;color:var(--muted);}}
.kv .item .n {{font-size:15px;font-weight:600;margin-top:2px;}}
.note {{font-size:12px;color:var(--muted);margin-top:16px;line-height:1.5;}}
.sec-title {{margin:18px 0 8px;font-size:15px;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <h1>全过程多项目成本分析看板</h1>
    <p>数据月：<b>{esc(month)}</b> ｜ 主源：{esc(source_name)} ｜ 生成：{esc(generated)}</p>
    <p>口径：金额默认 <b>万泰铢</b>。预算为广联达人材机（THB 已÷10000；CNY 保留原币并标注）。</p>
    <div class="chain">
      <span>① 预算人材机</span>
      <span>② 责任成本 / 投标毛利</span>
      <span>③ 已发生（过程）</span>
      <span>④ 预估剩余</span>
      <span>⑤ 预计结算 / 毛利</span>
    </div>
  </div>

  <div class="kpis">
    <div class="kpi"><div class="l">在建项目数</div><div class="v">{len(active)}</div><div class="u">已完工 {len(done)} 个另列</div></div>
    <div class="kpi"><div class="l">需关注（红+黄）</div><div class="v" style="color:var(--danger)">{n_risk}</div><div class="u">负毛利/低毛利/超进度</div></div>
    <div class="kpi"><div class="l">业主合同额合计</div><div class="v">{k_contract:,.0f}</div><div class="u">万泰铢 · 在建</div></div>
    <div class="kpi"><div class="l">预计结算额合计</div><div class="v">{k_settle:,.0f}</div><div class="u">万泰铢</div></div>
    <div class="kpi"><div class="l">责任成本合计</div><div class="v">{k_resp:,.0f}</div><div class="u">目标成本</div></div>
    <div class="kpi"><div class="l">已发生 / 预估剩余</div><div class="v">{k_actual:,.0f}</div><div class="u">预估 {k_forecast:,.0f}</div></div>
    <div class="kpi"><div class="l">预计总成本（梳理）</div><div class="v">{k_est:,.0f}</div><div class="u">消耗 {((consume_all or 0)*100):.0f}% 责任成本</div></div>
    <div class="kpi"><div class="l">组合预计毛利 / 率</div><div class="v">{k_exp_profit:,.0f}</div><div class="u">{(f"{portfolio_margin*100:.1f}%" if portfolio_margin is not None else "—")} · 投标毛利 {k_bid_profit:,.0f}</div></div>
  </div>

  <div class="panel">
    <h2>组合视图</h2>
    <div class="charts">
      <div class="chart-box"><canvas id="chartMargin"></canvas></div>
      <div class="chart-box"><canvas id="chartProgress"></canvas></div>
    </div>
    <div class="legend">
      <span><i style="background:var(--good)"></i>健康</span>
      <span><i style="background:var(--warn)"></i>需关注</span>
      <span><i style="background:var(--danger)"></i>高风险</span>
      <span><i style="background:var(--muted-risk)"></i>数据缺口</span>
    </div>
  </div>

  <div class="panel">
    <h2>在建 / 过程项目（点击行查看全过程钻取）</h2>
    <div class="filters">
      <label>风险
        <select id="fRisk">
          <option value="">全部</option>
          <option value="danger">高风险</option>
          <option value="warning">需关注</option>
          <option value="muted">数据缺口</option>
          <option value="good">健康</option>
        </select>
      </label>
      <label>项目经理
        <select id="fPm"><option value="">全部</option></select>
      </label>
      <label>搜索 <input id="fSearch" placeholder="编码/名称"/></label>
      <span class="hint">单位：万泰铢 · 消耗=已发生/责任成本 · 进度=完成比例</span>
    </div>
    <div class="table-scroll">
      <table id="tblActive">
        <thead>
          <tr>
            <th>编码</th><th>项目</th><th class="num">合同额</th><th class="num">责任成本</th>
            <th class="num">已发生</th><th class="num">预估</th><th class="num">预计总成本</th>
            <th class="num">预计结算</th><th class="num">预计毛利率</th><th class="num">投标毛利率</th>
            <th>进度 vs 消耗</th><th>预警</th>
          </tr>
        </thead>
        <tbody>
          {rows_active}
        </tbody>
      </table>
    </div>
  </div>

  <div class="panel">
    <h2 class="sec-title" style="margin-top:0">已完工项目（报送口径）</h2>
    <div class="table-scroll" style="max-height:320px">
      <table>
        <thead>
          <tr>
            <th>编码</th><th>项目</th><th class="num">收入</th><th class="num">成本</th>
            <th class="num">已发生</th><th class="num">预估</th><th class="num">成本</th>
            <th class="num">收入</th><th class="num">毛利率</th><th class="num">—</th>
            <th>状态</th><th>备注</th>
          </tr>
        </thead>
        <tbody>{rows_done}</tbody>
      </table>
    </div>
  </div>

  <div class="note">
    <b>使用说明</b><br/>
    1. 主决策口径来自「成本报公司」月表（责任成本、已发生、预估、预计毛利），与总部语言一致。<br/>
    2. 钻取面板中的「台账」来自 18_output 索引（采购/机具/人工等），用于对照客观发生；与项目自报差异大时需对账。<br/>
    3. 预算人材机：THB 已换算万泰铢；CNY 项目显示原币金额，完整可比需配置汇率（P2）。<br/>
    4. master 全月份毛利率字段当前大量为空，趋势仅在有数时显示。建议优先修复 scraper/join。<br/>
    5. 设计说明见：Desktop/code/S_数据分析模块/全过程成本分析看板_设计说明.md
  </div>
</div>

<div class="mask" id="mask" role="dialog" aria-modal="true">
  <div class="drawer" id="drawer" onclick="event.stopPropagation()">
    <button class="close" id="btnClose" type="button">关闭</button>
    <h3 id="dTitle">项目</h3>
    <div id="dSub" class="note" style="margin-top:0"></div>
    <div class="lifecycle" id="dLife"></div>
    <div class="kv" id="dKv"></div>
    <h4 style="margin:16px 0 6px">台账实际（万元汇总）</h4>
    <div class="kv" id="dLedger"></div>
    <h4 style="margin:16px 0 6px">预算人材机</h4>
    <div class="kv" id="dBudget"></div>
    <h4 style="margin:16px 0 6px">预警与说明</h4>
    <div id="dAlerts"></div>
  </div>
</div>

<script>
const PROJECTS = {projects_json};
const TREND = {trend_json};
const chartNames = {json.dumps(chart_names, ensure_ascii=False)};
const chartMargins = {json.dumps(chart_margins)};
const chartBid = {json.dumps(chart_bid)};
const chartConsume = {json.dumps(chart_consume)};
const chartProgress = {json.dumps(chart_progress)};

// PM filter options
const pms = [...new Set(PROJECTS.map(p => p.pm).filter(Boolean))].sort();
const fPm = document.getElementById('fPm');
pms.forEach(pm => {{
  const o = document.createElement('option');
  o.value = pm; o.textContent = pm; fPm.appendChild(o);
}});

function applyFilter() {{
  const risk = document.getElementById('fRisk').value;
  const pm = document.getElementById('fPm').value;
  const q = document.getElementById('fSearch').value.trim().toLowerCase();
  document.querySelectorAll('#tblActive tbody tr').forEach(tr => {{
    let ok = true;
    if (risk && tr.dataset.risk !== risk) ok = false;
    if (pm && tr.dataset.pm !== pm) ok = false;
    if (q) {{
      const t = (tr.dataset.name + ' ' + tr.children[0].innerText).toLowerCase();
      if (!t.includes(q)) ok = false;
    }}
    tr.style.display = ok ? '' : 'none';
  }});
}}
['fRisk','fPm','fSearch'].forEach(id => document.getElementById(id).addEventListener('input', applyFilter));

// Charts
const ctx1 = document.getElementById('chartMargin');
new Chart(ctx1, {{
  type: 'bar',
  data: {{
    labels: chartNames,
    datasets: [
      {{ label: '预计毛利率%', data: chartMargins, backgroundColor: 'rgba(31,138,101,.75)' }},
      {{ label: '投标毛利率%', data: chartBid, backgroundColor: 'rgba(58,123,213,.45)' }},
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ title: {{ display: true, text: '在建项目：投标 vs 预计毛利率' }}, legend: {{ position: 'bottom' }} }},
    scales: {{ x: {{ ticks: {{ maxRotation: 60, minRotation: 30, font: {{ size: 10 }} }} }}, y: {{ title: {{ display: true, text: '%' }} }} }}
  }}
}});

new Chart(document.getElementById('chartProgress'), {{
  type: 'bar',
  data: {{
    labels: chartNames,
    datasets: [
      {{ label: '形象进度%', data: chartProgress, backgroundColor: 'rgba(31,138,101,.7)' }},
      {{ label: '成本消耗%', data: chartConsume, backgroundColor: 'rgba(196,137,42,.65)' }},
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ title: {{ display: true, text: '进度 vs 成本消耗（消耗=已发生/责任成本）' }}, legend: {{ position: 'bottom' }} }},
    scales: {{ x: {{ ticks: {{ maxRotation: 60, minRotation: 30, font: {{ size: 10 }} }} }}, y: {{ title: {{ display: true, text: '%' }} }} }}
  }}
}});

// Drawer
const mask = document.getElementById('mask');
const drawer = document.getElementById('drawer');
function closeDrawer() {{
  mask.classList.remove('show');
  document.body.style.overflow = '';
}}
function openDrawer() {{
  mask.classList.add('show');
  document.body.style.overflow = 'hidden';
}}
document.getElementById('btnClose').onclick = (e) => {{ e.stopPropagation(); closeDrawer(); }};
mask.onclick = closeDrawer;
document.addEventListener('keydown', (e) => {{ if (e.key === 'Escape') closeDrawer(); }});

function fmt(n, d=0) {{
  if (n === null || n === undefined || n === '') return '—';
  const x = Number(n);
  if (Number.isNaN(x)) return '—';
  return x.toLocaleString('zh-CN', {{ maximumFractionDigits: d, minimumFractionDigits: d }});
}}
function pct(n) {{
  if (n === null || n === undefined) return '—';
  return (Number(n)*100).toFixed(1) + '%';
}}

function openProject(p) {{
  document.getElementById('dTitle').textContent = (p.code ? p.code + ' · ' : '') + p.name;
  document.getElementById('dSub').textContent = [p.pm, p.status || p.bucket, p.full_name].filter(Boolean).join(' ｜ ');
  document.getElementById('dLife').innerHTML = `
    <div class="s"><div class="t">①预算人材机</div><div class="n">${{fmt(p.budget_total,1)}}</div><div class="t">${{p.budget_unit||''}}</div></div>
    <div class="s"><div class="t">②责任成本</div><div class="n">${{fmt(p.resp_cost,1)}}</div><div class="t">万泰铢</div></div>
    <div class="s"><div class="t">③已发生</div><div class="n">${{fmt(p.actual,1)}}</div><div class="t">消耗 ${{pct(p.consume)}}</div></div>
    <div class="s"><div class="t">④预估剩余</div><div class="n">${{fmt(p.forecast,1)}}</div><div class="t">预计总成本 ${{fmt(p.est_cost,1)}}</div></div>
    <div class="s"><div class="t">⑤预计毛利</div><div class="n">${{fmt(p.exp_profit,1)}}</div><div class="t">${{pct(p.exp_margin)}}</div></div>
  `;
  document.getElementById('dKv').innerHTML = `
    <div class="item"><div class="t">业主合同额</div><div class="n">${{fmt(p.contract,1)}}</div></div>
    <div class="item"><div class="t">预计结算额</div><div class="n">${{fmt(p.settle,1)}}</div></div>
    <div class="item"><div class="t">投标毛利 / 率</div><div class="n">${{fmt(p.bid_profit,1)}} / ${{pct(p.bid_margin)}}</div></div>
    <div class="item"><div class="t">毛利侵蚀</div><div class="n">${{p.margin_erosion==null?'—':(p.margin_erosion*100).toFixed(1)+' pct'}}</div></div>
    <div class="item"><div class="t">形象进度</div><div class="n">${{pct(p.progress)}}</div></div>
    <div class="item"><div class="t">成本-进度差</div><div class="n">${{p.cost_progress_gap==null?'—':(p.cost_progress_gap*100).toFixed(1)+' pct'}}</div></div>
  `;
  document.getElementById('dLedger').innerHTML = `
    <div class="item"><div class="t">材料采购</div><div class="n">${{fmt(p.ledger_mat,1)}}</div></div>
    <div class="item"><div class="t">机具费</div><div class="n">${{fmt(p.ledger_mach,1)}}</div></div>
    <div class="item"><div class="t">人工费</div><div class="n">${{fmt(p.ledger_labor,1)}}</div></div>
    <div class="item"><div class="t">商混</div><div class="n">${{fmt(p.ledger_ready_mix,1)}}</div></div>
    <div class="item"><div class="t">分包合同</div><div class="n">${{fmt(p.ledger_sub,1)}}</div></div>
    <div class="item"><div class="t">材料欠款</div><div class="n">${{fmt(p.ledger_payable,1)}}</div></div>
  `;
  document.getElementById('dBudget').innerHTML = `
    <div class="item"><div class="t">预算合计 (${{p.budget_unit||'—'}})</div><div class="n">${{fmt(p.budget_total,1)}}</div></div>
    <div class="item"><div class="t">人</div><div class="n">${{fmt(p.budget_labor,1)}}</div></div>
    <div class="item"><div class="t">材</div><div class="n">${{fmt(p.budget_mat,1)}}</div></div>
    <div class="item"><div class="t">机</div><div class="n">${{fmt(p.budget_mach,1)}}</div></div>
  `;
  const alerts = (p.alerts && p.alerts.length) ? p.alerts.map(a => `<span class="tag">${{a}}</span>`).join(' ') : '<span class="tag ok">无预警</span>';
  document.getElementById('dAlerts').innerHTML = alerts +
    (p.note ? `<div class="note">备注：${{p.note}}</div>` : '') +
    (p.progress_note ? `<div class="note">进度：${{p.progress_note}}</div>` : '');
  openDrawer();
}}

document.querySelectorAll('.proj-row').forEach(tr => {{
  tr.addEventListener('click', () => {{
    try {{
      const p = JSON.parse(tr.getAttribute('data-json'));
      openProject(p);
    }} catch (e) {{ console.error(e); }}
  }});
}});
</script>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(description="全过程多项目成本分析看板")
    parser.add_argument("--month", default=None, help="如 2026年06月")
    parser.add_argument("--output", default=str(OUTPUT_HTML))
    args = parser.parse_args()

    report_path = find_latest_company_report(args.month)
    active, done, month = load_company_report(report_path)
    wan, bud = load_index()
    info = load_project_info()
    projects = enrich(active + done, wan, bud, info)
    trend = load_master_trend()

    REPORT_ROOT.mkdir(parents=True, exist_ok=True)
    html = build_html(projects, month, report_path.name, trend)
    out = Path(args.output)
    out.write_text(html, encoding="utf-8")

    n_active = sum(1 for p in projects if p["bucket"] != "已完工")
    n_risk = sum(1 for p in projects if p["bucket"] != "已完工" and p["risk"] in ("danger", "warning"))
    print(f"[OK] 看板已生成: {out}")
    print(f"     数据月: {month} | 源文件: {report_path.name}")
    print(f"     在建/过程: {n_active} | 已完工: {len(projects)-n_active} | 需关注: {n_risk}")
    print(f"     设计说明: Desktop/code/S_数据分析模块/全过程成本分析看板_设计说明.md")


if __name__ == "__main__":
    main()
