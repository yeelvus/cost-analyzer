#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合成本分析看板（单入口 · 单文件夹 · 可每月重跑）

输出目录（固定）：
  01成本统计/Z_report/综合成本看板/
    index.html              ← 唯一入口（大型综合看板）
    data_latest.json        ← 同源数据快照（可选调试）
    生成信息.txt
    archive/index_时间戳.html

看板能力：
  · 组合层：多项目 KPI / 筛选 / 风险
  · 点击项目 → 居中弹窗，Tab 综合分析：
      1 总览（责任成本·已发生·预估·毛利）
      2 预算人材机
      3 采购材料
      4 预算 vs 采购（量价）
      5 分包 vs 责任成本
      6 财务付款

用法：
  python build_unified_cost_dashboard.py
  python build_unified_cost_dashboard.py --month 2026年06月 --fx 4.9
  python build_unified_cost_dashboard.py --open
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
import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

try:
    from settings import CODE_ROOT, DATA_ROOT, REPORT_ROOT
    from common.paths import html_data_path, std_data_path as _std_p, sum_data_path as _sum_p
except Exception:
    CODE_ROOT = Path(__file__).resolve().parents[2]  # .../code
    DATA_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
    REPORT_ROOT = DATA_ROOT / "Z_report"

    def html_data_path(*parts, create=True):
        p = DATA_ROOT.joinpath("04_html", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

    def _sum_p(*parts, create=True, legacy=None):
        p = DATA_ROOT.joinpath("0_Sum", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

    def _std_p(*parts, create=True, legacy=None):
        p = DATA_ROOT.joinpath("0_Std", *parts)
        if create:
            p.mkdir(parents=True, exist_ok=True)
        return p

OUT_DIR = html_data_path("综合成本看板")
ARCHIVE_DIR = OUT_DIR / "archive"
CONFIG_PATH = CODE_ROOT / "common" / "budget_vs_process_config.json"

COMPANY_DIR = DATA_ROOT / "1_成本报公司"
BUDGET_SUM = _sum_p("预算部分", create=False)
BUDGET_STD = _std_p("预算部分", create=False, legacy=DATA_ROOT / "3_预算部分" / "2_人材机_std")
INDEX_XLSX = DATA_ROOT / "18_output" / "项目汇总索引.xlsx"
LABOR_XLSX = _sum_p("人工费", create=False) / "人工费汇总.xlsx"
PROC_SUM = _sum_p("采购台账", create=False) / "采购部采购台账_std_sum.xlsx"
PROC_CAT = _sum_p("采购台账", create=False) / "项目采购分类汇总.xlsx"
MACH_SUM = _sum_p("机具费", create=False) / "机具费_汇总.xlsx"
SUB_SUM = _sum_p("分包合同", create=False)
FIN_SUM = _sum_p("财务付款", create=False) / "财务付款数据汇总.xlsx"
RESP_DIR = _std_p("责任成本", create=False, legacy=DATA_ROOT / "3.1_责任成本" / "2_责任成本_std")
INFO_XLSX = DATA_ROOT / "0_项目基本信息情况" / "项目基本信息表.xlsx"

DEFAULT_FX = 4.90


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
    s = str(v).strip()
    return "" if s.lower() in {"nan", "none"} else s


def thb_wan(thb: float) -> float:
    return thb / 10000.0


def find_latest(dir_path: Path, pattern: str, exclude: str = "") -> Optional[Path]:
    if not dir_path.exists():
        return None
    files = [
        p for p in dir_path.glob(pattern)
        if p.is_file() and not p.name.startswith("~$") and (not exclude or exclude not in p.name)
    ]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def load_fx() -> float:
    if CONFIG_PATH.exists():
        try:
            return float(json.loads(CONFIG_PATH.read_text(encoding="utf-8")).get("fx_cny_to_thb", DEFAULT_FX))
        except Exception:
            pass
    return DEFAULT_FX


def budget_to_thb(amount: float, currency: str, fx: float) -> float:
    cur = (currency or "THB").upper()
    if cur in {"CNY", "RMB", "￥", "元", "人民币"}:
        return amount * fx
    return amount


def key_of(code: str, name: str) -> str:
    if code:
        return code
    return f"N::{name}" if name else ""


def idx_put(d: Dict[str, Dict], code: str, name: str, payload: Dict):
    """按编码/名称双索引写入（合并到同一 dict 引用）。"""
    k = key_of(code, name)
    if not k:
        return
    # 若编码与名称分别存在，合并到编码
    node = None
    if code and code in d:
        node = d[code]
    elif name and f"N::{name}" in d:
        node = d[f"N::{name}"]
    elif name and name in d:
        node = d[name]
    if node is None:
        node = {"code": code, "name": name}
        d[k] = node
        if code:
            d[code] = node
        if name:
            d[f"N::{name}"] = node
            d[name] = node
    else:
        if code:
            node["code"] = node.get("code") or code
            d[code] = node
        if name:
            node["name"] = node.get("name") or name
            d[f"N::{name}"] = node
            d[name] = node
    node.update(payload)


def idx_get(d: Dict[str, Dict], code: str, name: str) -> Dict:
    if code and code in d:
        return d[code]
    if name and f"N::{name}" in d:
        return d[f"N::{name}"]
    if name and name in d:
        return d[name]
    return {}


# ───────────────────────── loaders ─────────────────────────

def load_company(month_hint: Optional[str]) -> Tuple[List[Dict], str, str]:
    files = [
        p for p in COMPANY_DIR.glob("*泰国子公司各项目成本统计表*.xlsx")
        if not p.name.startswith("~$") and "对比" not in p.name and "内部" not in p.name
    ]
    if not files:
        return [], "", ""
    path = max(files, key=lambda p: p.stat().st_mtime)
    if month_hint:
        for p in files:
            if re.sub(r"[年0]", "", month_hint)[:6] in re.sub(r"[年0]", "", p.name):
                path = p
                break
    m = re.search(r"(\d{4})年0?(\d{1,2})月", path.name)
    month = f"{m.group(1)}年{int(m.group(2)):02d}月" if m else path.stem

    wb = openpyxl.load_workbook(path, data_only=True)
    active: List[Dict] = []
    ws = next((wb[n] for n in wb.sheetnames if "在建" in n), None)
    if ws is not None:
        header_row = 2
        for r in range(1, 6):
            if any("项目名称" in sn(ws.cell(r, c).value) for c in range(1, 12)):
                header_row = r
                break
        headers = [sn(ws.cell(header_row, c).value).replace("\n", "") or f"c{c}" for c in range(1, 22)]
        seen = {}
        for i, h in enumerate(headers):
            if h in seen:
                headers[i] = f"{h}_{seen[h]}"
                seen[h] += 1
            else:
                seen[h] = 1

        def col(*keys):
            for i, h in enumerate(headers):
                if any(k in h for k in keys):
                    return i
            return None

        ic = {
            "code": col("项目编码"), "name": col("项目名称"), "pm": col("项目经理"),
            "contract": col("业主合同额"), "resp": col("责任成本"),
            "bid_profit": col("投标毛利"), "bid_margin": col("含税金差（投标）", "含税金差"),
            "settle": col("业主预计结算额"), "est": col("梳理成本"),
            "actual": col("已发生"), "forecast": col("预估发生"),
            "exp_profit": col("预估毛利"), "exp_margin": col("预计毛利率"),
            "progress": col("完成比例"), "area": col("建筑面积"),
            "note": col("备注"), "pnote": col("进度概述"),
        }
        for r in range(header_row + 1, ws.max_row + 1):
            name = sn(ws.cell(r, (ic["name"] or 2) + 1).value)
            if not name or name in {"合计", "小计"}:
                continue

            def g(k):
                i = ic.get(k)
                return ws.cell(r, i + 1).value if i is not None else None

            contract, resp = sf(g("contract")), sf(g("resp"))
            settle, est = sf(g("settle")), sf(g("est"))
            actual, forecast = sf(g("actual")), sf(g("forecast"))
            exp_profit, bid_profit = sf(g("exp_profit")), sf(g("bid_profit"))
            exp_m = sf(g("exp_margin"))
            bid_m = sf(g("bid_margin"))
            if exp_m > 1.5:
                exp_m /= 100
            if bid_m > 1.5:
                bid_m /= 100
            if not exp_m and settle:
                exp_m = (settle - est) / settle if settle else 0
            if not bid_m and contract and bid_profit:
                bid_m = bid_profit / contract
            prog = sf(g("progress"))
            if prog > 1.5:
                prog /= 100
            consume = actual / resp if resp else None
            active.append({
                "bucket": "在建",
                "code": sn(g("code")),
                "name": name,
                "pm": sn(g("pm")),
                "contract": contract, "resp_cost": resp,
                "bid_profit": bid_profit, "bid_margin": bid_m or None,
                "settle": settle, "est_cost": est,
                "actual": actual, "forecast": forecast,
                "exp_profit": exp_profit, "exp_margin": exp_m or None,
                "progress": prog, "consume": consume,
                "margin_erosion": (bid_m - exp_m) if bid_m and exp_m else None,
                "cost_progress_gap": (consume - prog) if consume is not None and prog else None,
                "area": sf(g("area")), "note": sn(g("note")), "progress_note": sn(g("pnote")),
            })

    done: List[Dict] = []
    ws2 = next((wb[n] for n in wb.sheetnames if "完工" in n), None)
    if ws2 is not None:
        header_row = 2
        for r in range(1, 6):
            if any(sn(ws2.cell(r, c).value) == "项目名称" for c in range(1, 6)):
                header_row = r
                break
        for r in range(header_row + 1, ws2.max_row + 1):
            seq, name = ws2.cell(r, 1).value, sn(ws2.cell(r, 2).value)
            if not name or (isinstance(seq, str) and not str(seq).replace(".", "").isdigit()):
                continue
            if name.startswith("与业主") or name.startswith("一") or name.startswith("二"):
                continue
            rev, cost_tax, cost = sf(ws2.cell(r, 3).value), sf(ws2.cell(r, 4).value), sf(ws2.cell(r, 5).value)
            profit, margin = sf(ws2.cell(r, 7).value), sf(ws2.cell(r, 8).value)
            if margin > 1.5:
                margin /= 100
            if not margin and rev:
                margin = (rev - (cost_tax or cost)) / rev
            done.append({
                "bucket": "已完工", "code": "", "name": name, "pm": "",
                "contract": rev, "resp_cost": cost_tax or cost,
                "bid_profit": 0, "bid_margin": None,
                "settle": rev, "est_cost": cost_tax or cost,
                "actual": cost or cost_tax, "forecast": 0,
                "exp_profit": profit, "exp_margin": margin or None,
                "progress": 1.0, "consume": 1.0,
                "margin_erosion": None, "cost_progress_gap": 0,
                "area": 0, "note": "已完工", "progress_note": "",
            })
    wb.close()
    return active + done, month, path.name


def load_budget(fx: float) -> Tuple[Dict[str, Dict], Dict[str, Dict], str]:
    """预算人材机 + 材料量。金额→万泰铢。"""
    summary = find_latest(BUDGET_SUM, "预算人材机汇总*.xlsx") or find_latest(
        BUDGET_STD, "预算人材机汇总*.xlsx"
    )
    if not summary:
        return {}, {}, "缺失"
    df = pd.read_excel(summary, sheet_name="总汇总")
    mat = pd.read_excel(summary, sheet_name="材料汇总")
    bud_map: Dict[str, Dict] = {}
    for _, r in df.iterrows():
        code, name, cur = sn(r.get("项目编码")), sn(r.get("项目名称")), sn(r.get("币种")) or "THB"
        if not code and not name:
            continue
        def bw(col):
            return thb_wan(budget_to_thb(sf(r.get(col)), cur, fx))
        idx_put(bud_map, code, name, {
            "budget_currency": cur,
            "budget_labor": bw("人_明细"),
            "budget_mat": bw("材类明细"),
            "budget_mach": bw("设机_明细"),
            "budget_mgmt": bw("管_明细"),
            "budget_other": bw("其他类别_明细"),
            "budget_total": bw("标段人材机汇总"),
        })
    mat_map: Dict[str, Dict] = {}
    for _, r in mat.iterrows():
        code, name, cur = sn(r.get("项目编码")), sn(r.get("项目名称")), sn(r.get("币种")) or "THB"
        if not code and not name:
            continue
        def qty(prefix):
            return sf(r.get(f"{prefix}_数量"))
        def amt(prefix):
            return thb_wan(budget_to_thb(sf(r.get(f"{prefix}_金额")), cur, fx))
        idx_put(mat_map, code, name, {
            "bud_rebar_qty_kg": qty("钢筋"),
            "bud_rebar_amt": amt("钢筋"),
            "bud_conc_qty_m3": qty("混凝土"),
            "bud_conc_amt": amt("混凝土"),
            "bud_mortar_qty": qty("砂浆"),
            "bud_mortar_amt": amt("砂浆"),
            "bud_gravel_qty": qty("碎石"),
            "bud_gravel_amt": amt("碎石"),
            "bud_brick_qty": qty("砖"),
            "bud_brick_amt": amt("砖"),
            "bud_other_mat_amt": amt("其他材料"),
            "bud_mat_currency": cur,
        })
    return bud_map, mat_map, str(summary)


def load_labor() -> Tuple[Dict[str, Dict], str]:
    path = LABOR_XLSX if LABOR_XLSX.exists() else None
    if not path:
        return {}, "缺失"
    df = pd.read_excel(path, sheet_name="按项目汇总" if "按项目汇总" in pd.ExcelFile(path).sheet_names else 0)
    d: Dict[str, Dict] = {}
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        if not code and not name:
            continue
        idx_put(d, code, name, {
            "proc_labor": thb_wan(sf(r.get("总人工费"))),
            "labor_cn_formal": thb_wan(sf(r.get("中方正式工铢"))),
            "labor_cn_temp": thb_wan(sf(r.get("中方劳务工铢"))),
            "labor_local": thb_wan(sf(r.get("子公司属地工铢"))),
            "labor_agency": thb_wan(sf(r.get("劳务公司铢"))),
            "labor_visa": thb_wan(sf(r.get("签证费"))),
        })
    return d, str(path)


def load_purchase() -> Tuple[Dict[str, Dict], Dict[str, Dict], str]:
    d: Dict[str, Dict] = {}
    src = []
    if PROC_SUM.exists():
        src.append(str(PROC_SUM))
        df = pd.read_excel(PROC_SUM, sheet_name="项目汇总")
        for _, r in df.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            if not code and not name:
                continue
            idx_put(d, code, name, {"proc_purchase": thb_wan(sf(r.get("总采购金额")))})
    # 商混
    rm = find_latest(REPORT_ROOT, "搅拌站混凝土汇总报表_*.xlsx", exclude="仅有")
    if rm:
        src.append(str(rm))
        df = pd.read_excel(rm, sheet_name=0)
        for _, r in df.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            if "金额（万泰铢）" in df.columns:
                wan = sf(r.get("金额（万泰铢）"))
            else:
                wan = thb_wan(sf(r.get("金额")))
            node = idx_get(d, code, name)
            prev = sf(node.get("proc_ready_mix")) if node else 0
            idx_put(d, code, name, {
                "proc_ready_mix": prev + wan,
                "ready_mix_m3": sf(r.get("总方量")),
            })
    for node in {id(v): v for v in d.values()}.values():
        node["proc_mat"] = sf(node.get("proc_purchase")) + sf(node.get("proc_ready_mix"))

    # 分类
    cat_map: Dict[str, Dict] = {}
    if PROC_CAT.exists():
        src.append(str(PROC_CAT))
        df = pd.read_excel(PROC_CAT, sheet_name="项目汇总")
        for _, r in df.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            idx_put(cat_map, code, name, {
                "buy_rebar_kg": sf(r.get("钢筋_kg")),
                "buy_rebar_amt": thb_wan(sf(r.get("钢筋_金额"))),
                "buy_cement_kg": sf(r.get("砂浆水泥_kg")),
                "buy_cement_amt": thb_wan(sf(r.get("砂浆水泥_金额"))),
                "buy_brick_m3": sf(r.get("砖_m3")),
                "buy_brick_amt": thb_wan(sf(r.get("砖_金额"))),
                "buy_formwork_m2": sf(r.get("模板_m2")),
                "buy_formwork_amt": thb_wan(sf(r.get("模板_金额"))),
                "buy_gravel_t": sf(r.get("碎石_t")),
                "buy_gravel_amt": thb_wan(sf(r.get("碎石_金额"))),
                "buy_other_amt": thb_wan(sf(r.get("其他_金额"))),
                "buy_cat_total": thb_wan(sf(r.get("合计金额"))),
            })
    return d, cat_map, "; ".join(src) or "缺失"


def load_machine() -> Tuple[Dict[str, Dict], str]:
    path = MACH_SUM if MACH_SUM.exists() else find_latest(DATA_ROOT / "10_机具费🚜" / "B_机具费_sum", "*汇总*.xlsx")
    if not path or not path.exists():
        return {}, "缺失"
    df = pd.read_excel(path, sheet_name="项目汇总机具费" if "项目汇总机具费" in pd.ExcelFile(path).sheet_names else 0)
    d: Dict[str, Dict] = {}
    col = "费用合计(泰铢)" if "费用合计(泰铢)" in df.columns else next((c for c in df.columns if "费用" in str(c)), None)
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        idx_put(d, code, name, {
            "proc_mach": thb_wan(sf(r.get(col))) if col else 0,
            "mach_count": sf(r.get("机具种类数")),
            "mach_uses": sf(r.get("使用次数")),
        })
    return d, str(path)


def load_subcontract() -> Tuple[Dict[str, Dict], str]:
    path = find_latest(SUB_SUM, "分包签订合同汇总_*.xlsx")
    if not path:
        return {}, "缺失"
    sheet = "万泰铢" if "万泰铢" in pd.ExcelFile(path).sheet_names else 0
    df = pd.read_excel(path, sheet_name=sheet)
    d: Dict[str, Dict] = {}
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        # 兼容列名
        contract = sf(r.get("合同额（万泰铢）") or r.get("合同额（泰铢）"))
        if "合同额（泰铢）" in df.columns and "合同额（万泰铢）" not in df.columns:
            contract = thb_wan(contract)
        paid = sf(r.get("分包已付款（万泰铢）") or r.get("分包已付款（泰铢）"))
        if "分包已付款（泰铢）" in df.columns and "分包已付款（万泰铢）" not in df.columns:
            paid = thb_wan(paid)
        ratio = sf(r.get("分包付款比例"))
        unpaid = sf(r.get("分包未付款"))
        if not unpaid and contract:
            unpaid = contract - paid
        idx_put(d, code, name, {
            "sub_contract": contract,
            "sub_paid": paid,
            "sub_unpaid": unpaid,
            "sub_pay_ratio": ratio if ratio else (paid / contract if contract else None),
        })
    return d, str(path)


def load_finance() -> Tuple[Dict[str, Dict], str]:
    path = FIN_SUM if FIN_SUM.exists() else REPORT_ROOT / "财务付款数据汇总.xlsx"
    if not path.exists():
        return {}, "缺失"
    xl = pd.ExcelFile(path)
    d: Dict[str, Dict] = {}
    # 总体：金额已是万泰铢量级（与科目表一致）
    total = pd.read_excel(path, sheet_name="总体汇总")
    for _, r in total.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        idx_put(d, code, name, {
            "fin_payable": sf(r.get("应付金额")),
            "fin_paid": sf(r.get("实付金额")),
            "fin_vat": sf(r.get("增值税额")),
            "fin_withhold": sf(r.get("预扣税额")),
            "fin_exvat": sf(r.get("不含税额")),
        })
    if "项目一级科目汇总" in xl.sheet_names:
        subj = pd.read_excel(path, sheet_name="项目一级科目汇总")
        for _, r in subj.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            cats = {}
            for c in subj.columns:
                if c in ("项目编码", "项目名称"):
                    continue
                cats[str(c)] = sf(r.get(c))
            idx_put(d, code, name, {"fin_subjects": cats})
    if "项目按月宽表" in xl.sheet_names:
        wide = pd.read_excel(path, sheet_name="项目按月宽表")
        for _, r in wide.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            months = []
            for c in wide.columns:
                m = re.match(r"金额汇总_(\d{4}-\d{2})", str(c))
                if m:
                    months.append({"month": m.group(1), "amount": sf(r.get(c))})
            months.sort(key=lambda x: x["month"])
            idx_put(d, code, name, {"fin_monthly": months})
    return d, str(path)


def load_resp_detail() -> Tuple[Dict[str, Dict], str]:
    """责任成本明细：分包类条目合计 + 全部分项。"""
    d: Dict[str, Dict] = {}
    if not RESP_DIR.exists():
        return d, "缺失"
    files = list(RESP_DIR.glob("THCC*.xlsx")) + list(RESP_DIR.glob("LSBM*.xlsx"))
    for f in files:
        m = re.match(r"([A-Z0-9-]+)_(.+?)_责任成本", f.stem)
        code = m.group(1) if m else ""
        name = m.group(2) if m else f.stem
        try:
            df = pd.read_excel(f)
        except Exception:
            continue
        df.columns = [str(c).replace("\n", "").strip() for c in df.columns]
        amt_col = next((c for c in df.columns if "金额" in c or "综合合价" in c), None)
        name_col = next((c for c in df.columns if c in ("名称", "项目") or "名称" in c), None)
        if not amt_col or not name_col:
            continue
        items = []
        sub_sum = 0.0
        total = 0.0
        for _, r in df.iterrows():
            nm = sn(r.get(name_col))
            am = sf(r.get(amt_col))
            if not nm or am == 0:
                # 仍记录分类标题无金额
                if nm and am == 0:
                    continue
                continue
            total += am
            is_sub = "分包" in nm
            if is_sub:
                sub_sum += am
            items.append({"name": nm, "amount_wan": thb_wan(am), "is_subcontract": is_sub})
        idx_put(d, code, name, {
            "resp_detail_total": thb_wan(total),
            "resp_subcontract_budget": thb_wan(sub_sum),
            "resp_items": items[:40],  # 限制体积
        })
    return d, str(RESP_DIR)


def load_info() -> Dict[str, Dict]:
    if not INFO_XLSX.exists():
        return {}
    try:
        df = pd.read_excel(INFO_XLSX, sheet_name="在建项目")
    except Exception:
        df = pd.read_excel(INFO_XLSX)
    d: Dict[str, Dict] = {}
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        idx_put(d, code, name, {
            "status": sn(r.get("项目状态")),
            "full_name": sn(r.get("项目全称")),
            "pm_info": sn(r.get("项目经理")),
        })
    return d


# ───────────────────────── assemble ─────────────────────────

def risk_of(p: Dict) -> Tuple[str, List[str]]:
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
    # 执行率
    if p.get("exec_total") and p["exec_total"] >= 1.3:
        alerts.append(f"人材机超预算{(p['exec_total']-1)*100:.0f}%")
    elif p.get("exec_total") and p["exec_total"] >= 1.15:
        alerts.append(f"人材机执行率{(p['exec_total'])*100:.0f}%")
    if p.get("sub_vs_resp") and p["sub_vs_resp"] > 1.1:
        alerts.append("分包合同额>责任成本分包口径")
    if not p.get("resp_cost"):
        alerts.append("缺责任成本")
    if not p.get("budget_total"):
        alerts.append("缺预算")
    if "负毛利" in alerts or any(a.startswith("毛利侵蚀") for a in alerts) or any("超预算" in a for a in alerts):
        risk = "danger"
    elif any(a in alerts for a in ("低毛利", "成本超进度")) or any("执行率" in a for a in alerts):
        risk = "warning"
    elif any(a.startswith("缺") for a in alerts):
        risk = "muted"
    else:
        risk = "good"
    return risk, alerts


def assemble(month_hint: Optional[str], fx: float) -> Dict[str, Any]:
    company, month, company_src = load_company(month_hint)
    bud_map, bud_mat_map, bud_src = load_budget(fx)
    labor_map, labor_src = load_labor()
    pur_map, cat_map, pur_src = load_purchase()
    mach_map, mach_src = load_machine()
    sub_map, sub_src = load_subcontract()
    fin_map, fin_src = load_finance()
    resp_map, resp_src = load_resp_detail()
    info_map = load_info()

    # universe: start from company, add any with budget
    projects: Dict[str, Dict] = {}
    for p in company:
        code, name = p.get("code") or "", p.get("name") or ""
        k = key_of(code, name)
        projects[k] = dict(p)

    # ensure budget-only projects appear
    for node in {id(v): v for v in bud_map.values()}.values():
        code, name = node.get("code") or "", node.get("name") or ""
        k = key_of(code, name)
        if k not in projects and not any(
            (code and x.get("code") == code) or (name and x.get("name") == name) for x in projects.values()
        ):
            projects[k] = {
                "bucket": "其他", "code": code, "name": name, "pm": "",
                "contract": 0, "resp_cost": 0, "bid_profit": 0, "bid_margin": None,
                "settle": 0, "est_cost": 0, "actual": 0, "forecast": 0,
                "exp_profit": 0, "exp_margin": None, "progress": 0, "consume": None,
                "margin_erosion": None, "cost_progress_gap": None,
                "area": 0, "note": "", "progress_note": "",
            }

    out_list = []
    for p in projects.values():
        code, name = p.get("code") or "", p.get("name") or ""
        b = idx_get(bud_map, code, name)
        bm = idx_get(bud_mat_map, code, name)
        lab = idx_get(labor_map, code, name)
        pur = idx_get(pur_map, code, name)
        cat = idx_get(cat_map, code, name)
        mach = idx_get(mach_map, code, name)
        sub = idx_get(sub_map, code, name)
        fin = idx_get(fin_map, code, name)
        resp = idx_get(resp_map, code, name)
        info = idx_get(info_map, code, name)

        p.update({k: v for k, v in b.items() if k not in ("code", "name")})
        p.update({k: v for k, v in bm.items() if k not in ("code", "name")})
        p.update({k: v for k, v in lab.items() if k not in ("code", "name")})
        p.update({k: v for k, v in pur.items() if k not in ("code", "name")})
        p.update({k: v for k, v in cat.items() if k not in ("code", "name")})
        p.update({k: v for k, v in mach.items() if k not in ("code", "name")})
        p.update({k: v for k, v in sub.items() if k not in ("code", "name")})
        p.update({k: v for k, v in fin.items() if k not in ("code", "name")})
        p.update({k: v for k, v in resp.items() if k not in ("code", "name")})
        if info:
            p["status"] = info.get("status") or p.get("status") or p.get("bucket")
            p["full_name"] = info.get("full_name") or ""
            if not p.get("pm"):
                p["pm"] = info.get("pm_info") or ""
        else:
            p.setdefault("status", p.get("bucket"))
            p.setdefault("full_name", "")

        # 过程人材机合计
        p["proc_labor"] = sf(p.get("proc_labor"))
        p["proc_mat"] = sf(p.get("proc_mat"))
        p["proc_mach"] = sf(p.get("proc_mach"))
        p["proc_total"] = p["proc_labor"] + p["proc_mat"] + p["proc_mach"]
        p["budget_total"] = sf(p.get("budget_total"))
        p["budget_labor"] = sf(p.get("budget_labor"))
        p["budget_mat"] = sf(p.get("budget_mat"))
        p["budget_mach"] = sf(p.get("budget_mach"))

        def er(a, b):
            return (a / b) if b > 1e-9 else None

        p["exec_labor"] = er(p["proc_labor"], p["budget_labor"])
        p["exec_mat"] = er(p["proc_mat"], p["budget_mat"])
        p["exec_mach"] = er(p["proc_mach"], p["budget_mach"])
        p["exec_total"] = er(p["proc_total"], p["budget_total"])

        # 预算 vs 采购量
        p["cmp_rebar_qty"] = {
            "budget": sf(p.get("bud_rebar_qty_kg")),
            "process": sf(p.get("buy_rebar_kg")),
            "unit": "kg",
            "ratio": er(sf(p.get("buy_rebar_kg")), sf(p.get("bud_rebar_qty_kg"))),
        }
        p["cmp_rebar_amt"] = {
            "budget": sf(p.get("bud_rebar_amt")),
            "process": sf(p.get("buy_rebar_amt")),
            "unit": "万泰铢",
            "ratio": er(sf(p.get("buy_rebar_amt")), sf(p.get("bud_rebar_amt"))),
        }
        # 混凝土：预算 m3 vs 商混方量
        p["cmp_concrete"] = {
            "budget_qty": sf(p.get("bud_conc_qty_m3")),
            "process_qty": sf(p.get("ready_mix_m3")),
            "budget_amt": sf(p.get("bud_conc_amt")),
            "process_amt": sf(p.get("proc_ready_mix")),
            "unit": "m³",
            "qty_ratio": er(sf(p.get("ready_mix_m3")), sf(p.get("bud_conc_qty_m3"))),
        }

        # 分包 vs 责任成本
        p["sub_contract"] = sf(p.get("sub_contract"))
        p["sub_paid"] = sf(p.get("sub_paid"))
        resp_cost = sf(p.get("resp_cost"))
        resp_sub = sf(p.get("resp_subcontract_budget"))
        p["sub_vs_resp"] = er(p["sub_contract"], resp_sub) if resp_sub else er(p["sub_contract"], resp_cost)
        p["sub_pay_vs_resp"] = er(p["sub_paid"], resp_cost) if resp_cost else None

        # 财务（已是万泰铢）
        p["fin_paid"] = sf(p.get("fin_paid"))
        p["fin_payable"] = sf(p.get("fin_payable"))
        p["fin_pay_ratio"] = er(p["fin_paid"], p["fin_payable"]) if p["fin_payable"] else None

        risk, alerts = risk_of(p)
        p["risk"] = risk
        p["alerts"] = alerts

        # clean non-serializable
        for k, v in list(p.items()):
            if isinstance(v, float) and pd.isna(v):
                p[k] = None
        out_list.append(p)

    # sort
    order = {"danger": 0, "warning": 1, "muted": 2, "good": 3}
    out_list.sort(key=lambda x: (order.get(x.get("risk"), 9), -(x.get("settle") or 0)))

    meta = {
        "title": "综合成本分析看板",
        "month": month,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "fx_cny_to_thb": fx,
        "unit": "万泰铢（除非另行标注）",
        "sources": {
            "报公司": company_src,
            "预算": bud_src,
            "人工": labor_src,
            "采购商混": pur_src,
            "机具": mach_src,
            "分包": sub_src,
            "财务": fin_src,
            "责任成本明细": resp_src,
        },
        "counts": {
            "projects": len(out_list),
            "active": sum(1 for p in out_list if p.get("bucket") == "在建"),
            "done": sum(1 for p in out_list if p.get("bucket") == "已完工"),
            "risk": sum(1 for p in out_list if p.get("risk") in ("danger", "warning")),
        },
    }
    return {"meta": meta, "projects": out_list}


# ───────────────────────── HTML ─────────────────────────


def _load_chart_js() -> str:
    """读取本地 Chart.js 源码，供内联进 HTML（规避 file:// 外链拦截）。"""
    candidates = [
        CODE_ROOT / "templates" / "vendor" / "chart.umd.min.js",
        Path(__file__).resolve().parent / "templates" / "vendor" / "chart.umd.min.js",
        OUT_DIR / "vendor" / "chart.umd.min.js",
    ]
    path = next((p for p in candidates if p.exists() and p.stat().st_size > 1000), None)
    if not path:
        return "/* Chart.js missing */ console.warn('Chart.js 未内联：请放置 templates/vendor/chart.umd.min.js');"
    js = path.read_text(encoding="utf-8", errors="replace")
    # 防止库代码中的 </script> 截断 HTML
    js = js.replace("</script>", "<\\/script>").replace("</Script>", "<\\/Script>")
    return js


def build_html(payload: Dict[str, Any]) -> str:
    """从模板注入 JSON + 内联 Chart.js，单文件可在 file:// 下打开。"""
    meta = payload.get("meta") or {}
    data_json = json.dumps(payload, ensure_ascii=False, default=str)
    # 防止 JSON 内出现 </script> 截断页面脚本
    data_json = data_json.replace("<", "\\u003c")

    tpl_candidates = [
        CODE_ROOT / "templates" / "unified_cost_dashboard.html",
        Path(__file__).resolve().parent / "templates" / "unified_cost_dashboard.html",
    ]
    tpl_path = next((p for p in tpl_candidates if p.exists()), None)
    if not tpl_path:
        raise FileNotFoundError("找不到看板模板 templates/unified_cost_dashboard.html")

    chart_js = _load_chart_js()
    html = tpl_path.read_text(encoding="utf-8")
    html = (
        html.replace("__CHART_JS_INLINE__", chart_js)
        .replace("__DATA_JSON__", data_json)
        .replace("__TITLE__", str(meta.get("title") or "综合成本分析看板"))
        .replace("__MONTH__", str(meta.get("month") or ""))
    )
    return html


# ───────────────────────── main ─────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="综合成本分析看板（单文件夹单入口）")
    parser.add_argument("--month", default=None)
    parser.add_argument("--fx", type=float, default=None)
    parser.add_argument("--open", action="store_true")
    args = parser.parse_args()

    fx = args.fx if args.fx is not None else load_fx()
    print(f"[1] 组装数据 (fx={fx}) …")
    payload = assemble(args.month, fx)
    meta = payload["meta"]
    print(f"    项目 {meta['counts']['projects']} | 在建 {meta['counts']['active']} | 风险 {meta['counts']['risk']}")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)

    # 本地 Chart.js（避免 CDN / Tracking Prevention / file 协议拦截）
    vendor_dir = OUT_DIR / "vendor"
    vendor_dir.mkdir(parents=True, exist_ok=True)
    chart_dst = vendor_dir / "chart.umd.min.js"
    chart_src_candidates = [
        CODE_ROOT / "templates" / "vendor" / "chart.umd.min.js",
        Path(__file__).resolve().parent / "templates" / "vendor" / "chart.umd.min.js",
        chart_dst,
    ]
    chart_src = next((p for p in chart_src_candidates if p.exists() and p.stat().st_size > 1000), None)
    if chart_src and chart_src.resolve() != chart_dst.resolve():
        shutil.copy2(chart_src, chart_dst)
    if not chart_dst.exists() or chart_dst.stat().st_size < 1000:
        print("[warn] 本地 Chart.js 缺失，图表可能不可用。请检查 templates/vendor/chart.umd.min.js")
    else:
        print(f"[2] 本地图表库: {chart_dst} ({chart_dst.stat().st_size} bytes)")

    # clean old multi-html confusion: only index in root
    html = build_html(payload)
    index = OUT_DIR / "index.html"
    index.write_text(html, encoding="utf-8")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arch = ARCHIVE_DIR / f"index_{stamp}.html"
    arch.write_text(html, encoding="utf-8")

    (OUT_DIR / "data_latest.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )

    info = [
        f"综合成本分析看板",
        f"生成时间: {meta['generated']}",
        f"数据月: {meta['month']}",
        f"汇率 CNY→THB: {fx}",
        f"入口: {index}",
        f"归档: {arch}",
        "",
        "每月更新:",
        "  cd Desktop/code && python3 build_unified_cost_dashboard.py --month YYYY年MM月",
        "",
        "数据源:",
    ]
    for k, v in meta["sources"].items():
        info.append(f"  - {k}: {v}")
    (OUT_DIR / "生成信息.txt").write_text("\n".join(info), encoding="utf-8")

    # 在 Z_report 放一个跳转说明，避免多入口
    jump = html_data_path() / "打开综合成本看板.html"
    jump.write_text(
        f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<meta http-equiv="refresh" content="0;url=综合成本看板/index.html"/>
<title>跳转综合看板</title></head>
<body style="font-family:sans-serif;padding:40px">
<p>正在进入综合成本分析看板…</p>
<p><a href="综合成本看板/index.html">若未自动跳转请点击这里</a></p>
</body></html>""",
        encoding="utf-8",
    )

    print(f"[OK] 唯一入口: {index}")
    print(f"     归档: {arch}")
    print(f"     快捷: {jump}")

    if args.open:
        import subprocess
        subprocess.run(["open", str(index)], check=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
