#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""账簿现金流量图

读取账簿「汇总」sheet，生成可交互 HTML：
  - 泰铢 / 人民币 两张现金流量图（当日净额 + 累计结余）
  - 勾选项目名称（可多选）自动重算
  - 点击某一天查看当天明细（可复制）
  - 选择时点查看结余
  - 换汇（人民币⇄泰铢）按币种分别计入；期末调汇默认不计入现金

用法：
  python 账簿现金流量图.py
  python 账簿现金流量图.py --excel /path/to.xlsx --out /path/to.html
"""
from __future__ import annotations

import argparse
import json
import math
import webbrowser
from datetime import datetime
from pathlib import Path

import pandas as pd

DEFAULT_EXCEL = Path(
    "/Users/wyattsun/Desktop/财务数据/账簿汇总_分类整理2026-09-05_副本2.xlsx"
)
DEFAULT_OUT = Path("/Users/wyattsun/Desktop/财务数据/账簿现金流量分析.html")
SHEET = "汇总"

NEED_COLS = [
    "日期",
    "凭证字号",
    "项目编码",
    "项目名称",
    "一级科目",
    "二级科目",
    "三级科目",
    "原始摘要",
    "汇率",
    "借方THB",
    "借方CNY",
    "实付金额THB",
    "实付金额CNY（原始）",
    "实付金额CNY",
    "币种",
    "收支",
    "方向",
    "支出性质",
    "科目编码",
]

NUM_COLS = [
    "汇率",
    "借方THB",
    "借方CNY",
    "实付金额THB",
    "实付金额CNY（原始）",
    "实付金额CNY",
]

FX_NATURES = {"人民币➡️泰铢", "泰铢➡️人民币", "人民币→泰铢", "泰铢→人民币"}
REVAL_NATURES = {"z_期末调汇", "期末调汇"}
INTERNAL_NATURES = {"银行➡️现金", "现金➡️银行", "取现"}


def _s(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    t = str(v).strip()
    return "" if t.lower() in {"nan", "none", "nat"} else t


def _f(v) -> float:
    if v is None or v is False:
        return 0.0
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        try:
            if v != v:
                return 0.0
        except Exception:
            pass
        return float(v)
    t = str(v).strip().replace(",", "").replace("，", "").replace(" ", "")
    if not t or t.lower() in {"nan", "none", "nat", "-"}:
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def load_ledger(excel: Path, sheet: str = SHEET) -> pd.DataFrame:
    if not excel.exists():
        raise FileNotFoundError(f"找不到账簿: {excel}")
    df = pd.read_excel(excel, sheet_name=sheet)
    missing = [c for c in NEED_COLS if c not in df.columns]
    if missing:
        raise KeyError(f"「{sheet}」缺少列: {missing}\n现有: {list(df.columns)}")
    out = df[NEED_COLS].copy()
    out["日期"] = pd.to_datetime(out["日期"], errors="coerce")
    out = out.dropna(subset=["日期"])
    for c in NUM_COLS:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    for c in ["凭证字号", "项目编码", "项目名称", "一级科目", "二级科目", "三级科目", "原始摘要", "币种", "收支", "方向", "支出性质"]:
        out[c] = out[c].map(_s)
    out["科目编码"] = out["科目编码"].map(lambda v: str(int(v)) if isinstance(v, (int, float)) and v == v and float(v).is_integer() else _s(v))
    out["收支"] = out["收支"].str.replace(" ", "", regex=False)
    out["币种"] = out["币种"].str.upper()
    out.loc[out["项目名称"] == "", "项目名称"] = "(未填项目)"
    return out


def signed_cash(row: pd.Series) -> tuple[float, float]:
    """按币种把本行记入真实现金：收入=+借方，支出=−实付。其它币种记 0。"""
    io = row["收支"]
    cur = row["币种"]
    thb = cny = 0.0
    if io == "收入":
        if cur == "THB":
            thb = _f(row["借方THB"])
        elif cur == "CNY":
            cny = _f(row["借方CNY"])
    elif io == "支出":
        if cur == "THB":
            thb = -_f(row["实付金额THB"])
        elif cur == "CNY":
            cny = -_f(row["实付金额CNY"] or row["实付金额CNY（原始）"])
    return round(thb, 2), round(cny, 2)


def build_payload(df: pd.DataFrame, source: Path) -> dict:
    records = []
    projects = {}
    for _, row in df.iterrows():
        thb, cny = signed_cash(row)
        nature = row["支出性质"]
        rec = {
            "d": row["日期"].strftime("%Y-%m-%d"),
            "v": row["凭证字号"],
            "code": row["项目编码"],
            "name": row["项目名称"],
            "l1": row["一级科目"],
            "l2": row["二级科目"],
            "l3": row["三级科目"],
            "s": row["原始摘要"],
            "fxr": round(_f(row["汇率"]), 6),
            "dThb": round(_f(row["借方THB"]), 2),
            "dCny": round(_f(row["借方CNY"]), 2),
            "pThb": round(_f(row["实付金额THB"]), 2),
            "pCny0": round(_f(row["实付金额CNY（原始）"]), 2),
            "pCny": round(_f(row["实付金额CNY"]), 2),
            "cur": row["币种"],
            "io": row["收支"],
            "dir": row["方向"],
            "nat": nature,
            "acc": row["科目编码"],
            "thb": thb,
            "cny": cny,
            "isFx": nature in FX_NATURES or row["一级科目"] in FX_NATURES,
            "isReval": nature in REVAL_NATURES,
            "isInternal": nature in INTERNAL_NATURES,
        }
        records.append(rec)
        key = rec["name"]
        info = projects.setdefault(key, {"name": key, "code": rec["code"], "n": 0})
        info["n"] += 1
        if rec["code"] and not info.get("code"):
            info["code"] = rec["code"]

    proj_list = sorted(projects.values(), key=lambda x: (x["name"] == "(未填项目)", x["name"]))
    dates = sorted({r["d"] for r in records})
    return {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source": str(source),
        "rowCount": len(records),
        "dateMin": dates[0] if dates else "",
        "dateMax": dates[-1] if dates else "",
        "projects": proj_list,
        "rows": records,
    }


def render_html(payload: dict) -> str:
    data_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace(
        "<", "\\u003c"
    )
    return HTML_TEMPLATE.replace("__DATA_JSON__", data_json)


HTML_TEMPLATE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>账簿现金流量分析</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root {
      --bg: #f4f1ea; --card: #fffcf6; --ink: #1f1c18; --muted: #6b645b;
      --line: #e6ddd0; --thb: #0f766e; --cny: #b45309; --accent: #1d4e89;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "PingFang SC", "Noto Sans SC", "Hiragino Sans GB", sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    .wrap { max-width: 1480px; margin: 0 auto; padding: 20px 16px 48px; }
    h1 { margin: 0 0 6px; font-size: 26px; }
    .sub { color: var(--muted); margin: 0 0 16px; font-size: 13px; }
    .layout { display: grid; grid-template-columns: 280px minmax(0,1fr); gap: 14px; align-items: start; }
    .card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 12px; }
    .side { position: sticky; top: 12px; max-height: calc(100vh - 24px); overflow: auto; }
    .side h3 { margin: 0 0 8px; font-size: 14px; }
    .search { width: 100%; padding: 6px 8px; border: 1px solid var(--line); border-radius: 8px; margin-bottom: 8px; }
    .btns { display: flex; gap: 6px; margin-bottom: 8px; flex-wrap: wrap; }
    .btns button, .copy-btn, .toolbar button {
      border: 1px solid var(--line); background: #fff; border-radius: 8px;
      padding: 4px 8px; cursor: pointer; font-size: 12px;
    }
    .btns button:hover, .copy-btn:hover { background: #f3efe6; }
    .plist label { display: flex; gap: 6px; align-items: flex-start; font-size: 12px; padding: 2px 0; }
    .plist label .meta { color: var(--muted); }
    .toolbar { display: flex; flex-wrap: wrap; gap: 10px 16px; align-items: end; margin-bottom: 10px; }
    .toolbar label { font-size: 12px; color: var(--muted); display: flex; flex-direction: column; gap: 4px; }
    .toolbar input[type=date], .toolbar select { padding: 4px 6px; border: 1px solid var(--line); border-radius: 8px; }
    .kpis { display: grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap: 8px; margin-bottom: 10px; }
    .kpi { background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; }
    .kpi .lab { font-size: 12px; color: var(--muted); }
    .kpi .val { font-size: 20px; font-weight: 650; margin-top: 2px; }
    .kpi.thb .val { color: var(--thb); }
    .kpi.cny .val { color: var(--cny); }
    .chart { height: 340px; }
    .hint { font-size: 12px; color: var(--muted); margin: 4px 0 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { border-bottom: 1px solid var(--line); padding: 5px 6px; text-align: left; vertical-align: top; white-space: nowrap; }
    td.wrap { white-space: normal; min-width: 180px; }
    td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
    tbody tr:hover { background: #f7f1e6; }
    .neg { color: #b42318; }
    .pos { color: #067647; }
    .tag { display: inline-block; padding: 1px 6px; border-radius: 999px; font-size: 11px; background: #eee8dc; }
    .tag.fx { background: #fde68a; }
    .tag.reval { background: #e5e7eb; }
    .table-wrap { max-height: 420px; overflow: auto; border: 1px solid var(--line); border-radius: 10px; }
    @media (max-width: 980px) {
      .layout { grid-template-columns: 1fr; }
      .side { position: static; max-height: 280px; }
      .kpis { grid-template-columns: 1fr 1fr; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>账簿现金流量分析</h1>
    <p class="sub" id="meta"></p>
    <div class="layout">
      <aside class="card side">
        <h3>项目名称（可多选）</h3>
        <input class="search" id="q" placeholder="筛选项目…" />
        <div class="btns">
          <button type="button" id="all">全选</button>
          <button type="button" id="none">全不选</button>
          <button type="button" id="inv">反选</button>
        </div>
        <div class="plist" id="plist"></div>
      </aside>
      <main>
        <div class="card">
          <div class="toolbar">
            <label>开始日期 <input type="date" id="from" /></label>
            <label>结束日期 <input type="date" id="to" /></label>
            <label>结余时点 <input type="date" id="asof" /></label>
            <label>账户
              <select id="acc">
                <option value="all">现金+银行（1001+1002）</option>
                <option value="1002">仅银行 1002</option>
                <option value="1001">仅现金 1001</option>
              </select>
            </label>
            <label style="flex-direction:row;align-items:center;gap:6px;margin-top:18px">
              <input type="checkbox" id="fx" checked /> 含换汇
            </label>
            <label style="flex-direction:row;align-items:center;gap:6px;margin-top:18px">
              <input type="checkbox" id="reval" /> 含期末调汇
            </label>
            <label style="flex-direction:row;align-items:center;gap:6px;margin-top:18px">
              <input type="checkbox" id="internal" checked /> 含内部转账/取现
            </label>
          </div>
          <div class="kpis">
            <div class="kpi thb"><div class="lab">泰铢结余（时点）</div><div class="val" id="kThb">—</div></div>
            <div class="kpi cny"><div class="lab">人民币结余（时点）</div><div class="val" id="kCny">—</div></div>
            <div class="kpi"><div class="lab">区间泰铢净额</div><div class="val" id="kThbNet">—</div></div>
            <div class="kpi"><div class="lab">区间人民币净额</div><div class="val" id="kCnyNet">—</div></div>
          </div>
          <div class="hint">换汇：人民币支出同时记泰铢收入（或反向），两张图分别反映各币种现金。期末调汇是账面调整，默认不计入。点击图中某一天查看明细。</div>
          <div class="chart" id="chartThb"></div>
          <div class="chart" id="chartCny"></div>
        </div>
        <div class="card" style="margin-top:12px">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap">
            <h3 style="margin:0" id="detailTitle">当日明细</h3>
            <div>
              <button class="copy-btn" type="button" id="copy">复制明细</button>
              <button class="copy-btn" type="button" id="copyBal">复制结余</button>
            </div>
          </div>
          <p class="hint" id="detailHint">点击现金流量图上的某一天，或先选结余时点。</p>
          <div class="table-wrap">
            <table id="tbl">
              <thead>
                <tr>
                  <th>日期</th><th>凭证字号</th><th>项目编码</th><th>项目名称</th>
                  <th>一级</th><th>二级</th><th>三级</th><th>原始摘要</th>
                  <th>币种</th><th>收支</th><th>方向</th><th>支出性质</th><th>科目</th>
                  <th class="num">汇率</th>
                  <th class="num">借方THB</th><th class="num">借方CNY</th>
                  <th class="num">实付THB</th><th class="num">实付CNY原始</th><th class="num">实付CNY</th>
                  <th class="num">计入THB</th><th class="num">计入CNY</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  </div>
  <script>
    const DATA = __DATA_JSON__;
    const plist = document.getElementById('plist');
    const selected = new Set(DATA.projects.map(p => p.name));
    let activeDate = '';

    function fmt(n) {
      if (n === 0 || n == null || Number.isNaN(n)) return '0';
      const abs = Math.abs(n);
      const s = abs.toLocaleString('zh-CN', {maximumFractionDigits: 2});
      return n < 0 ? '-' + s : s;
    }
    function clsNum(n) { return n < 0 ? 'num neg' : (n > 0 ? 'num pos' : 'num'); }

    function renderList(filter) {
      const q = (filter || '').trim().toLowerCase();
      plist.innerHTML = DATA.projects.map(p => {
        if (q && !(p.name + p.code).toLowerCase().includes(q)) return '';
        const checked = selected.has(p.name) ? 'checked' : '';
        return `<label><input type="checkbox" data-name="${p.name.replace(/"/g,'&quot;')}" ${checked} /><span>${p.name}<div class="meta">${p.code || ''} · ${p.n} 笔</div></span></label>`;
      }).join('');
    }

    function selectedSet() {
      return selected;
    }

    function includeRow(r) {
      if (!selected.has(r.name)) return false;
      const acc = document.getElementById('acc').value;
      if (acc !== 'all' && String(r.acc) !== acc) return false;
      if (!document.getElementById('fx').checked && r.isFx) return false;
      if (!document.getElementById('reval').checked && r.isReval) return false;
      if (!document.getElementById('internal').checked && r.isInternal) return false;
      const from = document.getElementById('from').value;
      const to = document.getElementById('to').value;
      if (from && r.d < from) return false;
      if (to && r.d > to) return false;
      return true;
    }

    function includeRowAsOf(r, asof) {
      if (!selected.has(r.name)) return false;
      const acc = document.getElementById('acc').value;
      if (acc !== 'all' && String(r.acc) !== acc) return false;
      if (!document.getElementById('fx').checked && r.isFx) return false;
      if (!document.getElementById('reval').checked && r.isReval) return false;
      if (!document.getElementById('internal').checked && r.isInternal) return false;
      if (asof && r.d > asof) return false;
      return true;
    }

    function daily(field, from, to) {
      const map = {};
      for (const r of DATA.rows) {
        if (!includeRow(r)) continue;
        map[r.d] = (map[r.d] || 0) + (r[field] || 0);
      }
      const dates = Object.keys(map).sort();
      let cum = 0;
      const xs = [], nets = [], cums = [];
      for (const d of dates) {
        if (from && d < from) { cum += map[d]; continue; }
        if (to && d > to) continue;
        cum += map[d];
        xs.push(d); nets.push(Math.round(map[d]*100)/100); cums.push(Math.round(cum*100)/100);
      }
      return {xs, nets, cums};
    }

    function draw(id, title, color, series) {
      const traces = [
        {type:'bar', x:series.xs, y:series.nets, name:'当日净额', marker:{color: color+'55'}, yaxis:'y2', hovertemplate:'%{x}<br>当日净额 %{y:,.2f}<extra></extra>'},
        {type:'scatter', mode:'lines+markers', x:series.xs, y:series.cums, name:'累计结余', line:{color:color, width:2}, marker:{size:5}, hovertemplate:'%{x}<br>累计结余 %{y:,.2f}<extra></extra>'}
      ];
      const asof = document.getElementById('asof').value;
      const shapes = [];
      if (asof) shapes.push({type:'line', x0:asof, x1:asof, y0:0, y1:1, yref:'paper', line:{color:'#94a3b8', dash:'dot'}});
      Plotly.react(id, traces, {
        title: {text: title, font:{size:14}},
        margin: {t:36, r:56, l:56, b:40},
        hovermode: 'x unified',
        legend: {orientation:'h', y:1.12},
        xaxis: {type:'date'},
        yaxis: {title:'累计结余', zeroline:true},
        yaxis2: {title:'当日净额', overlaying:'y', side:'right', showgrid:false, zeroline:false},
        bargap: 0.15,
        plot_bgcolor: '#fffcf6',
        paper_bgcolor: '#fffcf6',
        shapes
      }, {displayModeBar: true, responsive: true});
    }

    function refresh() {
      const from = document.getElementById('from').value;
      const to = document.getElementById('to').value;
      const asof = document.getElementById('asof').value || to;
      const thb = daily('thb', from, to);
      const cny = daily('cny', from, to);
      draw('chartThb', '泰铢现金流量（THB）', '#0f766e', thb);
      draw('chartCny', '人民币现金流量（CNY）', '#b45309', cny);

      let balThb = 0, balCny = 0, netThb = 0, netCny = 0;
      for (const r of DATA.rows) {
        if (includeRowAsOf(r, asof)) { balThb += r.thb; balCny += r.cny; }
        if (includeRow(r)) { netThb += r.thb; netCny += r.cny; }
      }
      document.getElementById('kThb').textContent = fmt(balThb);
      document.getElementById('kCny').textContent = fmt(balCny);
      document.getElementById('kThbNet').textContent = fmt(netThb);
      document.getElementById('kCnyNet').textContent = fmt(netCny);
      if (activeDate) renderDetail(activeDate);
    }

    function renderDetail(date) {
      activeDate = date;
      document.getElementById('detailTitle').textContent = date + ' 当日明细';
      const rows = DATA.rows.filter(r => r.d === date && includeRow(r));
      document.getElementById('detailHint').textContent =
        `共 ${rows.length} 笔 · 泰铢净额 ${fmt(rows.reduce((s,r)=>s+r.thb,0))} · 人民币净额 ${fmt(rows.reduce((s,r)=>s+r.cny,0))} · 可拖选单元格复制，或点「复制明细」`;
      const tb = document.querySelector('#tbl tbody');
      tb.innerHTML = rows.map(r => {
        const tags = [];
        if (r.isFx) tags.push('<span class="tag fx">换汇</span>');
        if (r.isReval) tags.push('<span class="tag reval">调汇</span>');
        if (r.isInternal) tags.push('<span class="tag">内部</span>');
        return `<tr>
          <td>${r.d}</td><td>${r.v||''}</td><td>${r.code||''}</td><td>${r.name||''}</td>
          <td>${r.l1||''}</td><td>${r.l2||''}</td><td>${r.l3||''}</td>
          <td class="wrap">${r.s||''} ${tags.join(' ')}</td>
          <td>${r.cur||''}</td><td>${r.io||''}</td><td>${r.dir||''}</td>
          <td>${r.nat||''}</td><td>${r.acc||''}</td>
          <td class="num">${r.fxr||''}</td>
          <td class="${clsNum(r.dThb)}">${fmt(r.dThb)}</td>
          <td class="${clsNum(r.dCny)}">${fmt(r.dCny)}</td>
          <td class="${clsNum(r.pThb)}">${fmt(r.pThb)}</td>
          <td class="${clsNum(r.pCny0)}">${fmt(r.pCny0)}</td>
          <td class="${clsNum(r.pCny)}">${fmt(r.pCny)}</td>
          <td class="${clsNum(r.thb)}">${fmt(r.thb)}</td>
          <td class="${clsNum(r.cny)}">${fmt(r.cny)}</td>
        </tr>`;
      }).join('') || '<tr><td colspan="21">当天无符合筛选的记录</td></tr>';
    }

    function tableTsv() {
      const rows = [...document.querySelectorAll('#tbl tr')].map(tr =>
        [...tr.children].map(td => td.innerText.replace(/\t/g,' ').replace(/\n/g,' ')).join('\t')
      );
      return rows.join('\n');
    }

    function bindChart(id) {
      const el = document.getElementById(id);
      el.on('plotly_click', ev => {
        const x = ev.points && ev.points[0] && ev.points[0].x;
        if (!x) return;
        const d = String(x).slice(0,10);
        document.getElementById('asof').value = d;
        renderDetail(d);
        refresh();
      });
    }

    document.getElementById('meta').textContent =
      `来源：${DATA.source} · ${DATA.rowCount} 笔 · ${DATA.dateMin} ～ ${DATA.dateMax} · 生成 ${DATA.generatedAt}`;
    document.getElementById('from').value = DATA.dateMin;
    document.getElementById('to').value = DATA.dateMax;
    document.getElementById('asof').value = DATA.dateMax;
    renderList('');
    refresh();
    bindChart('chartThb');
    bindChart('chartCny');

    plist.addEventListener('change', e => {
      const t = e.target;
      if (t && t.dataset.name) {
        if (t.checked) selected.add(t.dataset.name); else selected.delete(t.dataset.name);
        refresh();
      }
    });
    document.getElementById('q').addEventListener('input', e => renderList(e.target.value));
    document.getElementById('all').onclick = () => { DATA.projects.forEach(p => selected.add(p.name)); renderList(document.getElementById('q').value); refresh(); };
    document.getElementById('none').onclick = () => { selected.clear(); renderList(document.getElementById('q').value); refresh(); };
    document.getElementById('inv').onclick = () => {
      DATA.projects.forEach(p => selected.has(p.name) ? selected.delete(p.name) : selected.add(p.name));
      renderList(document.getElementById('q').value); refresh();
    };
    ['from','to','asof','acc','fx','reval','internal'].forEach(id => {
      document.getElementById(id).addEventListener('change', refresh);
    });
    document.getElementById('asof').addEventListener('change', () => {
      if (document.getElementById('asof').value) renderDetail(document.getElementById('asof').value);
    });
    document.getElementById('copy').onclick = async () => {
      await navigator.clipboard.writeText(tableTsv());
      alert('明细已复制，可粘贴到 Excel');
    };
    document.getElementById('copyBal').onclick = async () => {
      const asof = document.getElementById('asof').value;
      const txt = `结余时点\t${asof}\n泰铢结余\t${document.getElementById('kThb').textContent}\n人民币结余\t${document.getElementById('kCny').textContent}\n区间泰铢净额\t${document.getElementById('kThbNet').textContent}\n区间人民币净额\t${document.getElementById('kCnyNet').textContent}`;
      await navigator.clipboard.writeText(txt);
      alert('结余已复制');
    };
  </script>
</body>
</html>
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="账簿汇总 → 泰铢/人民币现金流量 HTML")
    parser.add_argument("--excel", default=str(DEFAULT_EXCEL), help="账簿 Excel 路径")
    parser.add_argument("--sheet", default=SHEET, help="工作表名，默认 汇总")
    parser.add_argument("--out", default=str(DEFAULT_OUT), help="输出 HTML 路径")
    parser.add_argument("--no-open", action="store_true", help="生成后不自动打开浏览器")
    args = parser.parse_args()

    excel = Path(args.excel).expanduser()
    out = Path(args.out).expanduser()
    print(f"读取: {excel}  [{args.sheet}]")
    df = load_ledger(excel, args.sheet)
    print(f"有效行: {len(df):,}  项目: {df['项目名称'].nunique()}  日期: {df['日期'].min().date()} ~ {df['日期'].max().date()}")
    payload = build_payload(df, excel)
    html = render_html(payload)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    size_mb = out.stat().st_size / (1024 * 1024)
    print(f"已生成: {out}  ({size_mb:.1f} MB)")
    if not args.no_open:
        webbrowser.open(out.as_uri())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
