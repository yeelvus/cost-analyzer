from __future__ import annotations

import json
import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from flask import Flask, jsonify, request

EXCEL_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/E_采购部采购台账_sum/采购台账_simplify_std_merged.xlsx"
)
JSON_DB_PATH = Path(__file__).with_name("procurement_query_db.json")

CANONICAL_FIELDS = [
    "project_code",
    "project_name",
    "material_category",
    "item_name",
    "spec",
    "unit",
    "quantity",
    "unit_price",
    "amount_without_tax",
    "tax_7",
    "amount_with_tax",
    "currency",
    "exchange_rate",
    "year_month",
    "standard_date",
    "supplier",
    "remark",
    "contract_date",
]

COLUMN_ALIASES = {
    "project_code": ["项目编码"],
    "project_name": ["项目名称"],
    "material_category": ["材料归类", "物料分类"],
    "item_name": ["物品名称", "材料名称", "品名"],
    "spec": ["规格", "规格型号"],
    "unit": ["单位"],
    "quantity": ["数量"],
    "unit_price": ["单价"],
    "amount_without_tax": ["不含税金额"],
    "tax_7": ["7%增值税", "税额"],
    "amount_with_tax": ["含税金额"],
    "currency": ["货币种类", "币种"],
    "exchange_rate": ["汇率"],
    "year_month": ["年月", "月份"],
    "standard_date": ["标准日期", "日期"],
    "supplier": ["供货商名称", "供应商", "供应商名称"],
    "remark": ["备注"],
    "contract_date": ["签订日期"],
}

HTML_PAGE = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>采购台账查询系统</title>
  <style>
    :root {
      --bg: #f7f4ef;
      --panel: #ffffff;
      --ink: #1e2a36;
      --sub: #617182;
      --line: #d7d3cb;
      --accent: #0f766e;
      --accent-2: #0b5f59;
      --warn: #9a3412;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "PingFang SC", "Noto Sans CJK SC", "Microsoft YaHei", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(1200px 480px at 110% -10%, #dbe9e6 0%, transparent 70%),
        radial-gradient(900px 380px at -10% -20%, #f4e7d8 0%, transparent 70%),
        var(--bg);
    }
    .wrap {
      width: min(1280px, 96vw);
      margin: 24px auto;
      display: grid;
      gap: 14px;
    }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 14px;
      box-shadow: 0 8px 30px rgba(30, 42, 54, 0.06);
      padding: 16px;
    }
    h1 { margin: 0 0 8px; font-size: 24px; }
    .sub { color: var(--sub); font-size: 13px; }
    .grid {
      display: grid;
      grid-template-columns: repeat(6, minmax(120px, 1fr));
      gap: 10px;
      margin-top: 12px;
    }
    .field { display: flex; flex-direction: column; gap: 6px; }
    .field label { font-size: 12px; color: var(--sub); }
    input, select, button {
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 8px 10px;
      font-size: 14px;
      background: #fff;
      color: var(--ink);
    }
    button {
      cursor: pointer;
      background: var(--accent);
      color: #fff;
      border-color: var(--accent);
      transition: transform .12s ease, background .12s ease;
    }
    button:hover { background: var(--accent-2); transform: translateY(-1px); }
    .toolbar {
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
      margin-top: 8px;
    }
    .meta {
      display: flex;
      gap: 18px;
      flex-wrap: wrap;
      font-size: 13px;
      color: var(--sub);
      margin-top: 10px;
    }
    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 12px;
      max-height: 70vh;
      background: #fff;
    }
    table {
      border-collapse: collapse;
      width: max(100%, 980px);
      font-size: 13px;
    }
    thead th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #f2f7f6;
      color: #26424b;
      border-bottom: 1px solid var(--line);
      text-align: left;
    }
    th, td { padding: 8px 10px; border-bottom: 1px solid #ece9e3; }
    tbody tr:hover { background: #f9fcfb; }
    .empty { color: var(--warn); font-size: 14px; padding: 8px 0 2px; }
    @media (max-width: 980px) {
      .grid { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
      h1 { font-size: 20px; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>采购台账查询系统</h1>
      <div class="sub">支持关键词检索、年月筛选、供应商筛选，并可切换“最近价格”或“供应商情况”视图。</div>

      <div class="grid">
        <div class="field">
          <label>关键词</label>
          <input id="q" placeholder="项目/材料/规格/供应商" />
        </div>

        <div class="field">
          <label>年月</label>
          <select id="year_month"><option value="">全部</option></select>
        </div>

        <div class="field">
          <label>供应商</label>
          <select id="supplier"><option value="">全部</option></select>
        </div>

        <div class="field">
          <label>视图模式</label>
          <select id="mode">
            <option value="latest">最近价格</option>
            <option value="supplier">供应商情况</option>
            <option value="all">明细记录</option>
          </select>
        </div>

        <div class="field">
          <label>返回条数</label>
          <input id="limit" type="number" min="20" max="5000" value="200" />
        </div>

        <div class="field" style="justify-content:flex-end;">
          <label>&nbsp;</label>
          <button id="searchBtn">查询</button>
        </div>
      </div>

      <div class="meta" id="meta"></div>
      <div class="empty" id="empty"></div>
    </div>

    <div class="card table-wrap">
      <table id="resultTable">
        <thead><tr id="theadRow"></tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </div>

  <script>
    async function loadMeta() {
      const r = await fetch('/api/meta');
      const data = await r.json();

      const ymSel = document.getElementById('year_month');
      const supSel = document.getElementById('supplier');

      for (const ym of data.year_months) {
        const op = document.createElement('option');
        op.value = ym;
        op.textContent = ym;
        ymSel.appendChild(op);
      }
      for (const s of data.suppliers) {
        const op = document.createElement('option');
        op.value = s;
        op.textContent = s;
        supSel.appendChild(op);
      }
    }

    function renderTable(columns, rows) {
      const theadRow = document.getElementById('theadRow');
      const tbody = document.getElementById('tbody');
      theadRow.innerHTML = '';
      tbody.innerHTML = '';

      for (const c of columns) {
        const th = document.createElement('th');
        th.textContent = c;
        theadRow.appendChild(th);
      }

      for (const row of rows) {
        const tr = document.createElement('tr');
        for (const c of columns) {
          const td = document.createElement('td');
          td.textContent = row[c] ?? '';
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
    }

    async function runQuery() {
      const q = document.getElementById('q').value.trim();
      const year_month = document.getElementById('year_month').value;
      const supplier = document.getElementById('supplier').value;
      const mode = document.getElementById('mode').value;
      const limit = document.getElementById('limit').value || '200';

      const params = new URLSearchParams({ q, year_month, supplier, mode, limit });
      const r = await fetch('/api/query?' + params.toString());
      const data = await r.json();

      document.getElementById('meta').innerHTML =
        `<span>模式: ${data.mode_label}</span>` +
        `<span>总命中: ${data.total_hits}</span>` +
        `<span>返回: ${data.returned}</span>` +
        `<span>数据版本时间: ${data.generated_at}</span>`;

      const empty = document.getElementById('empty');
      empty.textContent = data.returned === 0 ? '未匹配到结果，请调整关键词或筛选条件。' : '';

      renderTable(data.columns, data.rows);
    }

    document.getElementById('searchBtn').addEventListener('click', runQuery);
    document.getElementById('q').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') runQuery();
    });

    loadMeta().then(runQuery);
  </script>
</body>
</html>
"""


def _pick_column(columns: list[str], alias_list: list[str]) -> str | None:
    for alias in alias_list:
        if alias in columns:
            return alias
    return None


def _normalize_year_month(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""

    # 支持 2025年03月 / 2025-03 / 2025/3 / 2025.3 等格式。
    m = re.search(r"((?:19|20)\d{2})\D{0,3}(0?[1-9]|1[0-2])", text)
    if not m:
        return ""
    year = int(m.group(1))
    month = int(m.group(2))
    return f"{year:04d}-{month:02d}"


def _month_sort_key(ym: str) -> int:
    m = re.fullmatch(r"((?:19|20)\d{2})-(0[1-9]|1[0-2])", ym or "")
    if not m:
        return -1
    return int(m.group(1)) * 100 + int(m.group(2))


def _to_float(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    text = str(v).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return str(v).strip()


def _to_json_safe(v: Any) -> Any:
  if v is None:
    return ""
  if isinstance(v, float) and math.isnan(v):
    return ""
  if isinstance(v, (pd.Timestamp, datetime, date)):
    return v.isoformat()
  return v


def rebuild_json_db(excel_path: Path = EXCEL_PATH, json_path: Path = JSON_DB_PATH) -> dict[str, Any]:
    if not excel_path.exists():
        raise FileNotFoundError(f"未找到Excel文件: {excel_path}")

    df = pd.read_excel(excel_path, sheet_name="合并数据")
    source_columns = [str(c).strip() for c in df.columns]

    canonical_to_source: dict[str, str] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        hit = _pick_column(source_columns, aliases)
        if hit:
            canonical_to_source[canonical] = hit

    records: list[dict[str, Any]] = []
    for _, row in df.iterrows():
      rec: dict[str, Any] = {}
      for field in CANONICAL_FIELDS:
        source_col = canonical_to_source.get(field)
        raw_value: Any = row[source_col] if source_col else ""
        rec[field] = _to_json_safe(raw_value)

      rec["year_month"] = _normalize_year_month(rec.get("year_month"))
      rec["unit_price"] = _to_float(rec.get("unit_price"))
      rec["amount_with_tax"] = _to_float(rec.get("amount_with_tax"))
      rec["quantity"] = _to_float(rec.get("quantity"))

      rec["project_name"] = _safe_str(rec.get("project_name"))
      rec["item_name"] = _safe_str(rec.get("item_name"))
      rec["spec"] = _safe_str(rec.get("spec"))
      rec["supplier"] = _safe_str(rec.get("supplier"))
      rec["material_category"] = _safe_str(rec.get("material_category"))
      rec["unit"] = _safe_str(rec.get("unit"))
      rec["project_code"] = _safe_str(rec.get("project_code"))

      records.append(rec)

    suppliers = sorted({r["supplier"] for r in records if r.get("supplier")})
    year_months = sorted({r["year_month"] for r in records if r.get("year_month")}, key=_month_sort_key, reverse=True)

    payload = {
        "meta": {
            "source_excel": str(excel_path),
            "record_count": len(records),
            "year_months": year_months,
            "suppliers": suppliers,
            "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
        "records": records,
    }

    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def _record_match(rec: dict[str, Any], keyword: str) -> bool:
    if not keyword:
        return True
    keyword_l = keyword.lower()
    haystack = " | ".join(
        [
            rec.get("project_name", ""),
            rec.get("material_category", ""),
            rec.get("item_name", ""),
            rec.get("spec", ""),
            rec.get("supplier", ""),
            rec.get("project_code", ""),
        ]
    ).lower()
    return keyword_l in haystack


def _fmt(v: Any) -> Any:
    if v is None:
        return ""
    if isinstance(v, float):
        return round(v, 4)
    return v


def _latest_price_view(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str, str], dict[str, Any]] = {}

    for rec in records:
        key = (rec.get("item_name", ""), rec.get("spec", ""), rec.get("unit", ""))
        prev = grouped.get(key)
        if prev is None:
            grouped[key] = rec
            continue

        prev_k = _month_sort_key(prev.get("year_month", ""))
        curr_k = _month_sort_key(rec.get("year_month", ""))
        if curr_k > prev_k:
            grouped[key] = rec

    out = []
    for rec in grouped.values():
        out.append(
            {
                "项目名称": rec.get("project_name", ""),
                "物品名称": rec.get("item_name", ""),
                "规格": rec.get("spec", ""),
                "单位": rec.get("unit", ""),
                "最近单价": _fmt(rec.get("unit_price")),
                "最近年月": rec.get("year_month", ""),
                "供应商": rec.get("supplier", ""),
            }
        )

    out.sort(key=lambda x: (_month_sort_key(str(x.get("最近年月", ""))), str(x.get("物品名称", ""))), reverse=True)
    return out


def _supplier_view(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for rec in records:
        supplier = rec.get("supplier", "") or "未填写供应商"
        grouped.setdefault(supplier, []).append(rec)

    out = []
    for supplier, rows in grouped.items():
        prices: list[float] = []
        for r in rows:
            price_val = r.get("unit_price")
            if isinstance(price_val, (int, float)):
                prices.append(float(price_val))

        latest_row = sorted(rows, key=lambda r: _month_sort_key(r.get("year_month", "")), reverse=True)[0]

        out.append(
            {
                "供应商": supplier,
                "记录数": len(rows),
                "最近年月": latest_row.get("year_month", ""),
                "最近单价": _fmt(latest_row.get("unit_price")),
                "最低单价": _fmt(min(prices)) if prices else "",
                "最高单价": _fmt(max(prices)) if prices else "",
                "平均单价": _fmt(sum(prices) / len(prices)) if prices else "",
                "最近物品": latest_row.get("item_name", ""),
                "最近规格": latest_row.get("spec", ""),
            }
        )

    out.sort(key=lambda x: (x.get("记录数", 0), _month_sort_key(str(x.get("最近年月", "")))), reverse=True)
    return out


def _all_detail_view(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out = []
    ordered = sorted(records, key=lambda r: _month_sort_key(r.get("year_month", "")), reverse=True)
    for rec in ordered:
        out.append(
            {
                "项目名称": rec.get("project_name", ""),
                "物品名称": rec.get("item_name", ""),
                "规格": rec.get("spec", ""),
                "单位": rec.get("unit", ""),
                "数量": _fmt(rec.get("quantity")),
                "单价": _fmt(rec.get("unit_price")),
                "年月": rec.get("year_month", ""),
                "供应商": rec.get("supplier", ""),
            }
        )
    return out


def create_app(data: dict[str, Any]) -> Flask:
    app = Flask(__name__)
    records = data.get("records", [])
    meta = data.get("meta", {})

    @app.get("/")
    def index():
        return HTML_PAGE

    @app.get("/api/meta")
    def api_meta():
        return jsonify(
            {
                "record_count": meta.get("record_count", 0),
                "year_months": meta.get("year_months", []),
                "suppliers": meta.get("suppliers", []),
                "generated_at": meta.get("generated_at", ""),
            }
        )

    @app.get("/api/query")
    def api_query():
        q = (request.args.get("q") or "").strip()
        year_month = (request.args.get("year_month") or "").strip()
        supplier = (request.args.get("supplier") or "").strip()
        mode = (request.args.get("mode") or "latest").strip()

        try:
            limit = int(request.args.get("limit", "200"))
        except ValueError:
            limit = 200
        limit = max(20, min(5000, limit))

        filtered = []
        for rec in records:
            if year_month and rec.get("year_month", "") != year_month:
                continue
            if supplier and rec.get("supplier", "") != supplier:
                continue
            if not _record_match(rec, q):
                continue
            filtered.append(rec)

        if mode == "supplier":
            rows = _supplier_view(filtered)
            mode_label = "供应商情况"
        elif mode == "all":
            rows = _all_detail_view(filtered)
            mode_label = "明细记录"
        else:
            rows = _latest_price_view(filtered)
            mode_label = "最近价格"

        rows = rows[:limit]
        columns = list(rows[0].keys()) if rows else ["提示"]
        if not rows:
            rows = [{"提示": "没有匹配结果"}]

        return jsonify(
            {
                "mode": mode,
                "mode_label": mode_label,
                "total_hits": len(filtered),
                "returned": 0 if rows and "提示" in rows[0] else len(rows),
                "columns": columns,
                "rows": rows,
                "generated_at": meta.get("generated_at", ""),
            }
        )

    return app


def main() -> None:
    print("开始构建JSON数据库...")
    payload = rebuild_json_db(EXCEL_PATH, JSON_DB_PATH)
    print(f"JSON数据库已生成: {JSON_DB_PATH}")
    print(f"记录数: {payload['meta']['record_count']}")
    print("查询系统启动中: http://127.0.0.1:5050")

    app = create_app(payload)
    app.run(host="127.0.0.1", port=5050, debug=False)


if __name__ == "__main__":
    main()
