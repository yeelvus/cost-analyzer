#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
预算 vs 过程 · 人材机结构对比（可每月重复运行）

做什么
------
把「广联达预算人材机」与「过程台账实际（人/材/机）」对齐到同一项目、同一单位（万泰铢），
输出：
  1) 数据表 Excel（带时间戳 + 最新快照，便于每月更新）
  2) 交互 HTML 对比页（结构占比、执行率、超耗预警）

数据源（自动找最新）
------------------
预算:
  01成本统计/3_预算部分/2_人材机_std/【总汇总】所有标段.xlsx
  回退: 18_output/项目汇总索引.xlsx · 预算金额

过程:
  人: Z_report/人工费汇总.xlsx · 按项目汇总 · 总人工费
  材: 9_采购…/D_…/采购部采购台账_std_merged.xlsx · 项目汇总 · 总采购金额
      + Z_report/搅拌站混凝土汇总报表_*.xlsx · 金额
  机: 10_机具…/B_机具费_sum/机具费汇总_全量标准化_汇总.xlsx · 项目汇总机具费

用法
----
  # 每月常规：用当前盘里最新 std/sum 刷新对比
  python build_budget_vs_process.py

  # 指定人民币→泰铢汇率（预算 CNY 项目）
  python build_budget_vs_process.py --fx 4.85

  # 仅出 Excel、或指定输出目录
  python build_budget_vs_process.py --no-html
  python build_budget_vs_process.py --output-dir /path/to/Z_report

  # 业务月标签（写入报告标题与文件名，不改变数据过滤逻辑）
  python build_budget_vs_process.py --month 2026-06

配置
----
  Desktop/code/common/budget_vs_process_config.json  （首次运行自动生成）
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
import html as html_lib
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# ── paths ──────────────────────────────────────────────────────────
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

BUDGET_SUM_DIR = _sum_p("预算部分", create=False)
BUDGET_STD_DIR = _std_p("预算部分", create=False, legacy=DATA_ROOT / "3_预算部分" / "2_人材机_std")
INDEX_XLSX = DATA_ROOT / "18_output" / "项目汇总索引.xlsx"
LABOR_XLSX = _sum_p("人工费", create=False) / "人工费_汇总.xlsx"
PROC_SUM = _sum_p("采购台账", create=False) / "采购部采购台账_std_sum.xlsx"
MACH_SUM = _sum_p("机具费", create=False) / "机具费_汇总.xlsx"
READY_MIX_DIR = _sum_p("商混站", create=False)
CONFIG_PATH = CODE_ROOT / "common" / "budget_vs_process_config.json"

DEFAULT_FX_CNY_TO_THB = 4.90  # 可在 config 或 --fx 覆盖；约 1 CNY ≈ 4.9 THB


# ── utils ──────────────────────────────────────────────────────────

def sf(v: Any) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "")
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
    if s.lower() in {"nan", "none"}:
        return ""
    return s


def esc(s: Any) -> str:
    return html_lib.escape("" if s is None else str(s))


def find_latest(dir_path: Path, pattern: str) -> Optional[Path]:
    if not dir_path.exists():
        return None
    files = [p for p in dir_path.glob(pattern) if p.is_file() and not p.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


def load_or_init_config() -> Dict[str, Any]:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    cfg = {
        "fx_cny_to_thb": DEFAULT_FX_CNY_TO_THB,
        "warn_exec_ratio": 1.15,       # 过程/预算 > 115% 黄
        "danger_exec_ratio": 1.30,     # > 130% 红
        "warn_struct_gap_pct": 8.0,    # 结构占比差（百分点）
        "notes": "修改 fx_cny_to_thb 后重新运行本脚本即可刷新看板",
    }
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return cfg


def thb_to_wan(thb: float) -> float:
    return thb / 10000.0


def budget_to_thb(amount: float, currency: str, fx: float) -> float:
    cur = (currency or "").upper().strip()
    if not amount:
        return 0.0
    if cur in {"THB", "฿", "铢", "泰铢"}:
        return amount
    if cur in {"CNY", "RMB", "￥", "元", "人民币"}:
        return amount * fx
    # 未知币种：按原值当作 THB，并在备注标记
    return amount


# ── loaders ────────────────────────────────────────────────────────

def load_budget(fx: float) -> Tuple[pd.DataFrame, str]:
    """返回预算表（金额已转为万泰铢）与来源说明。"""
    summary = find_latest(BUDGET_SUM_DIR, "预算人材机汇总*.xlsx") or find_latest(
        BUDGET_STD_DIR, "预算人材机汇总*.xlsx"
    )
    source = ""
    if summary and summary.exists():
        df = pd.read_excel(summary, sheet_name="总汇总")
        source = str(summary)
    elif INDEX_XLSX.exists():
        df = pd.read_excel(INDEX_XLSX, sheet_name="预算金额")
        source = f"{INDEX_XLSX}::预算金额"
    else:
        raise FileNotFoundError("未找到预算总汇总，请先运行广联达人材机统计脚本")

    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        code = sn(r.get("项目编码"))
        name = sn(r.get("项目名称"))
        if not code and not name:
            continue
        cur = sn(r.get("币种")) or "THB"
        b_total = budget_to_thb(sf(r.get("标段人材机汇总")), cur, fx)
        b_labor = budget_to_thb(sf(r.get("人_明细")), cur, fx)
        b_mat = budget_to_thb(sf(r.get("材类明细")), cur, fx)
        b_mach = budget_to_thb(sf(r.get("设机_明细")), cur, fx)
        b_mgmt = budget_to_thb(sf(r.get("管_明细")), cur, fx)
        b_other = budget_to_thb(sf(r.get("其他类别_明细")), cur, fx)
        rows.append({
            "项目编码": code,
            "项目名称": name,
            "预算币种": cur,
            "预算_人_万泰铢": thb_to_wan(b_labor),
            "预算_材_万泰铢": thb_to_wan(b_mat),
            "预算_机_万泰铢": thb_to_wan(b_mach),
            "预算_管_万泰铢": thb_to_wan(b_mgmt),
            "预算_其他_万泰铢": thb_to_wan(b_other),
            "预算_人材机合计_万泰铢": thb_to_wan(b_total),
            "预算_钢筋_原币": sf(r.get("钢筋")),
            "预算_混凝土_原币": sf(r.get("混凝土")),
            "fx_cny_to_thb": fx if cur.upper() in {"CNY", "RMB", "￥", "元", "人民币"} else None,
        })
    out = pd.DataFrame(rows)
    return out, source


def load_labor() -> Tuple[pd.DataFrame, str]:
    path = LABOR_XLSX
    if not path.exists():
        for p in DATA_ROOT.rglob("人工费_汇总.xlsx"):
            if p.name.startswith("~$"):
                continue
            path = p
            break
    if not path.exists():
        return pd.DataFrame(columns=["项目编码", "项目名称", "过程_人_万泰铢"]), "缺失"

    # prefer 按项目汇总
    xl = pd.ExcelFile(path)
    sheet = "按项目汇总" if "按项目汇总" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    col = "总人工费" if "总人工费" in df.columns else None
    if col is None:
        for c in df.columns:
            if "总人工" in c or "人工费汇总" in c:
                col = c
                break
    rows = []
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        if not code and not name:
            continue
        amt = sf(r.get(col)) if col else 0.0
        # 若已是万泰铢 sheet
        if "万泰铢" in sheet or (col and "万泰铢" in str(col)):
            wan = amt
        else:
            wan = thb_to_wan(amt)
        rows.append({"项目编码": code, "项目名称": name, "过程_人_万泰铢": wan})
    return pd.DataFrame(rows), str(path)


def load_materials() -> Tuple[pd.DataFrame, str]:
    sources = []
    # 采购
    proc_map: Dict[str, float] = {}
    name_map: Dict[str, str] = {}
    if PROC_SUM.exists():
        sources.append(str(PROC_SUM))
        df = pd.read_excel(PROC_SUM, sheet_name="项目汇总")
        df.columns = [str(c).strip() for c in df.columns]
        for _, r in df.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            amt = thb_to_wan(sf(r.get("总采购金额")))
            key = code or name
            if not key:
                continue
            proc_map[key] = proc_map.get(key, 0.0) + amt
            if code:
                proc_map[code] = proc_map.get(code, 0.0)  # ensure
            if name:
                name_map[key] = name
                if code:
                    name_map[code] = name
    # 商混（优先完整报表，避开「仅有项目编码」精简版）
    rm_candidates = [
        p for p in READY_MIX_DIR.glob("搅拌站混凝土汇总报表_*.xlsx")
        if p.is_file() and not p.name.startswith("~$") and "仅有" not in p.name
    ]
    rm_path = max(rm_candidates, key=lambda p: p.stat().st_mtime) if rm_candidates else find_latest(
        READY_MIX_DIR, "搅拌站混凝土汇总报表_*.xlsx"
    )
    rm_map: Dict[str, float] = {}
    if rm_path:
        sources.append(str(rm_path))
        df = pd.read_excel(rm_path, sheet_name=0)
        df.columns = [str(c).strip() for c in df.columns]
        amt_col = "金额（万泰铢）" if "金额（万泰铢）" in df.columns else ("金额" if "金额" in df.columns else None)
        for _, r in df.iterrows():
            code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
            if amt_col == "金额（万泰铢）":
                wan = sf(r.get(amt_col))
            else:
                wan = thb_to_wan(sf(r.get(amt_col))) if amt_col else 0.0
            for key in filter(None, [code, name]):
                rm_map[key] = rm_map.get(key, 0.0) + wan
                if name:
                    name_map[key] = name

    keys = set(proc_map) | set(rm_map)
    rows = []
    for key in keys:
        # prefer code-like keys
        code = key if re.match(r"^[A-Z]{2,}", key) or key.startswith("THCC") or key.startswith("LSBM") else ""
        name = name_map.get(key, key if not code else "")
        if not code and re.match(r"THCC|LSBM|XCY", key):
            code = key
        mat = proc_map.get(key, 0.0) + rm_map.get(key, 0.0)
        # also try merge by both if we only have name
        rows.append({
            "项目编码": code,
            "项目名称": name or key,
            "过程_采购_万泰铢": proc_map.get(key, 0.0),
            "过程_商混_万泰铢": rm_map.get(key, 0.0),
            "过程_材_万泰铢": mat,
        })
    # consolidate by code if possible
    out = pd.DataFrame(rows)
    if out.empty:
        return out, "; ".join(sources) or "缺失"
    # group by code or name
    def gkey(r):
        return r["项目编码"] if r["项目编码"] else r["项目名称"]

    out["_k"] = out.apply(gkey, axis=1)
    agg = out.groupby("_k", as_index=False).agg({
        "项目编码": "first",
        "项目名称": "first",
        "过程_采购_万泰铢": "sum",
        "过程_商混_万泰铢": "sum",
        "过程_材_万泰铢": "sum",
    })
    return agg.drop(columns=["_k"], errors="ignore"), "; ".join(sources) or "缺失"


def load_machines() -> Tuple[pd.DataFrame, str]:
    path = MACH_SUM
    if not path.exists():
        alt = find_latest(DATA_ROOT / "10_机具费🚜" / "B_机具费_sum", "*汇总*.xlsx")
        path = alt or path
    if not path or not path.exists():
        return pd.DataFrame(columns=["项目编码", "项目名称", "过程_机_万泰铢"]), "缺失"

    xl = pd.ExcelFile(path)
    sheet = "项目汇总机具费" if "项目汇总机具费" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    amt_col = "费用合计(泰铢)" if "费用合计(泰铢)" in df.columns else None
    if amt_col is None:
        for c in df.columns:
            if "费用" in c or "合计" in c:
                amt_col = c
                break
    rows = []
    for _, r in df.iterrows():
        code, name = sn(r.get("项目编码")), sn(r.get("项目名称"))
        if not code and not name:
            continue
        rows.append({
            "项目编码": code,
            "项目名称": name,
            "过程_机_万泰铢": thb_to_wan(sf(r.get(amt_col))) if amt_col else 0.0,
        })
    return pd.DataFrame(rows), str(path)


def try_match_code(name: str, code: str, lookup) -> Tuple[str, str]:
    """可选：用 common.project 补全编码。"""
    if code:
        return code, name
    if not name or lookup is None:
        return code, name
    try:
        res = lookup(name) if callable(lookup) else None
        if isinstance(res, dict):
            return sn(res.get("项目编码") or res.get("code") or code), sn(res.get("标准名称") or res.get("项目名称") or name)
        if isinstance(res, (list, tuple)) and res:
            return sn(res[0]), name
    except Exception:
        pass
    return code, name


def get_matcher():
    try:
        from common.project import load_or_init_db, match_project
        from settings import MATCH_DB_PATH

        db = load_or_init_db(str(MATCH_DB_PATH))

        def _match(name: str):
            return match_project(name, db)

        return _match
    except Exception:
        return None


# ── merge & metrics ────────────────────────────────────────────────

def merge_all(
    budget: pd.DataFrame,
    labor: pd.DataFrame,
    materials: pd.DataFrame,
    machines: pd.DataFrame,
) -> pd.DataFrame:
    matcher = get_matcher()

    def norm_frame(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        df = df.copy()
        df["项目编码"] = df["项目编码"].map(sn)
        df["项目名称"] = df["项目名称"].map(sn)
        if matcher:
            codes, names = [], []
            for _, r in df.iterrows():
                c, n = try_match_code(r["项目名称"], r["项目编码"], matcher)
                codes.append(c)
                names.append(n or r["项目名称"])
            df["项目编码"] = codes
            df["项目名称"] = names
        return df

    budget = norm_frame(budget)
    labor = norm_frame(labor)
    materials = norm_frame(materials)
    machines = norm_frame(machines)

    # build universe of keys: prefer 编码
    def key_series(df: pd.DataFrame) -> pd.Series:
        return df.apply(lambda r: r["项目编码"] if r["项目编码"] else f"NAME::{r['项目名称']}", axis=1)

    frames = []
    for df, tag in [(budget, "b"), (labor, "l"), (materials, "m"), (machines, "h")]:
        if df.empty:
            continue
        d = df.copy()
        d["_key"] = key_series(d)
        frames.append((d, tag))

    if not frames:
        return pd.DataFrame()

    # start from budget if available else labor
    base = frames[0][0]
    for d, tag in frames[1:]:
        cols_keep = [c for c in d.columns if c not in {"项目编码", "项目名称"} or c == "_key"]
        # avoid duplicate name cols
        rename = {}
        for c in d.columns:
            if c in base.columns and c not in {"_key", "项目编码", "项目名称"}:
                rename[c] = f"{c}_{tag}"
        d2 = d.rename(columns=rename)
        # merge on _key
        base = base.merge(
            d2.drop(columns=[c for c in ["项目编码", "项目名称"] if c in d2.columns], errors="ignore"),
            on="_key",
            how="outer",
        )

    # recover code/name
    def pick_code(row):
        if sn(row.get("项目编码")):
            return sn(row.get("项目编码"))
        k = sn(row.get("_key"))
        if k.startswith("NAME::"):
            return ""
        return k

    def pick_name(row):
        if sn(row.get("项目名称")):
            return sn(row.get("项目名称"))
        k = sn(row.get("_key"))
        if k.startswith("NAME::"):
            return k.replace("NAME::", "", 1)
        return k

    # After outer merges, 项目编码/名称 may be missing on some — re-extract from keys and any left
    # Re-merge identity from each source
    id_map = {}
    for d, _ in frames:
        for _, r in d.iterrows():
            k = r["_key"]
            id_map.setdefault(k, {"项目编码": sn(r.get("项目编码")), "项目名称": sn(r.get("项目名称"))})
            if sn(r.get("项目编码")):
                id_map[k]["项目编码"] = sn(r.get("项目编码"))
            if sn(r.get("项目名称")):
                id_map[k]["项目名称"] = sn(r.get("项目名称"))

    base["项目编码"] = base["_key"].map(lambda k: id_map.get(k, {}).get("项目编码", ""))
    base["项目名称"] = base["_key"].map(lambda k: id_map.get(k, {}).get("项目名称", ""))

    # fill numeric
    num_cols = [
        "预算_人_万泰铢", "预算_材_万泰铢", "预算_机_万泰铢", "预算_管_万泰铢", "预算_其他_万泰铢", "预算_人材机合计_万泰铢",
        "过程_人_万泰铢", "过程_材_万泰铢", "过程_机_万泰铢", "过程_采购_万泰铢", "过程_商混_万泰铢",
    ]
    for c in num_cols:
        if c not in base.columns:
            base[c] = 0.0
        base[c] = base[c].map(sf)

    # 过程人材机合计
    base["过程_人材机合计_万泰铢"] = base["过程_人_万泰铢"] + base["过程_材_万泰铢"] + base["过程_机_万泰铢"]

    # 结构占比（预算 / 过程 各自合计为 100%）
    def ratio(part, total):
        t = sf(total)
        return (sf(part) / t) if t > 1e-9 else None

    for side, prefix in [("预算", "预算"), ("过程", "过程")]:
        tot = f"{prefix}_人材机合计_万泰铢"
        base[f"{prefix}_人占比"] = base.apply(lambda r: ratio(r[f"{prefix}_人_万泰铢"], r[tot]), axis=1)
        base[f"{prefix}_材占比"] = base.apply(lambda r: ratio(r[f"{prefix}_材_万泰铢"], r[tot]), axis=1)
        base[f"{prefix}_机占比"] = base.apply(lambda r: ratio(r[f"{prefix}_机_万泰铢"], r[tot]), axis=1)

    # 执行率 = 过程 / 预算
    def exec_r(proc, bud):
        b = sf(bud)
        if b <= 1e-9:
            return None
        return sf(proc) / b

    base["执行率_人"] = base.apply(lambda r: exec_r(r["过程_人_万泰铢"], r["预算_人_万泰铢"]), axis=1)
    base["执行率_材"] = base.apply(lambda r: exec_r(r["过程_材_万泰铢"], r["预算_材_万泰铢"]), axis=1)
    base["执行率_机"] = base.apply(lambda r: exec_r(r["过程_机_万泰铢"], r["预算_机_万泰铢"]), axis=1)
    base["执行率_合计"] = base.apply(lambda r: exec_r(r["过程_人材机合计_万泰铢"], r["预算_人材机合计_万泰铢"]), axis=1)

    # 超耗金额
    base["超耗_人_万泰铢"] = base["过程_人_万泰铢"] - base["预算_人_万泰铢"]
    base["超耗_材_万泰铢"] = base["过程_材_万泰铢"] - base["预算_材_万泰铢"]
    base["超耗_机_万泰铢"] = base["过程_机_万泰铢"] - base["预算_机_万泰铢"]
    base["超耗_合计_万泰铢"] = base["过程_人材机合计_万泰铢"] - base["预算_人材机合计_万泰铢"]

    # 结构差（百分点）
    def gap(a, b):
        if a is None or b is None or (isinstance(a, float) and pd.isna(a)) or (isinstance(b, float) and pd.isna(b)):
            return None
        return (sf(a) - sf(b)) * 100.0

    base["结构差_人_pct"] = base.apply(lambda r: gap(r["过程_人占比"], r["预算_人占比"]), axis=1)
    base["结构差_材_pct"] = base.apply(lambda r: gap(r["过程_材占比"], r["预算_材占比"]), axis=1)
    base["结构差_机_pct"] = base.apply(lambda r: gap(r["过程_机占比"], r["预算_机占比"]), axis=1)

    # 数据完整度标记
    def coverage(r):
        has_b = r["预算_人材机合计_万泰铢"] > 0
        has_p = r["过程_人材机合计_万泰铢"] > 0
        if has_b and has_p:
            return "双边齐全"
        if has_b:
            return "仅预算"
        if has_p:
            return "仅过程"
        return "无数据"

    base["数据覆盖"] = base.apply(coverage, axis=1)

    # drop helper
    base = base.drop(columns=["_key"], errors="ignore")
    # sort: 双边齐全 first, then by 执行率
    order = {"双边齐全": 0, "仅预算": 1, "仅过程": 2, "无数据": 3}
    base["_ord"] = base["数据覆盖"].map(order)
    base = base.sort_values(["_ord", "执行率_合计"], ascending=[True, False], na_position="last")
    base = base.drop(columns=["_ord"])
    return base.reset_index(drop=True)


def risk_of_row(r: pd.Series, cfg: Dict) -> str:
    er = r.get("执行率_合计")
    if er is None or (isinstance(er, float) and pd.isna(er)):
        if r.get("数据覆盖") != "双边齐全":
            return "muted"
        return "muted"
    if er >= cfg.get("danger_exec_ratio", 1.30):
        return "danger"
    if er >= cfg.get("warn_exec_ratio", 1.15):
        return "warning"
    # 结构扭曲
    for col in ("结构差_人_pct", "结构差_材_pct", "结构差_机_pct"):
        g = r.get(col)
        if g is not None and not (isinstance(g, float) and pd.isna(g)):
            if abs(sf(g)) >= cfg.get("warn_struct_gap_pct", 8.0) * 1.5:
                return "warning"
    return "good"


def alerts_of_row(r: pd.Series, cfg: Dict) -> List[str]:
    a = []
    if r.get("数据覆盖") == "仅预算":
        a.append("无过程台账")
    elif r.get("数据覆盖") == "仅过程":
        a.append("无预算")
    er = r.get("执行率_合计")
    if er is not None and not (isinstance(er, float) and pd.isna(er)):
        if er >= cfg.get("danger_exec_ratio", 1.30):
            a.append(f"合计超预算{(er-1)*100:.0f}%")
        elif er >= cfg.get("warn_exec_ratio", 1.15):
            a.append(f"合计接近/超预算{(er-1)*100:.0f}%")
    for lab, col in [("人", "执行率_人"), ("材", "执行率_材"), ("机", "执行率_机")]:
        e = r.get(col)
        if e is not None and not (isinstance(e, float) and pd.isna(e)) and e >= cfg.get("danger_exec_ratio", 1.30):
            a.append(f"{lab}超耗{(e-1)*100:.0f}%")
    thr = cfg.get("warn_struct_gap_pct", 8.0)
    for lab, col in [("人", "结构差_人_pct"), ("材", "结构差_材_pct"), ("机", "结构差_机_pct")]:
        g = r.get(col)
        if g is not None and not (isinstance(g, float) and pd.isna(g)) and abs(sf(g)) >= thr:
            a.append(f"{lab}占比偏差{sf(g):+.1f}pct")
    return a


# ── export excel ───────────────────────────────────────────────────

def export_excel(df: pd.DataFrame, meta: Dict, out_dir: Path, stamp: str) -> Tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    latest = out_dir / "预算vs过程_人材机对比_最新.xlsx"
    stamped = out_dir / f"预算vs过程_人材机对比_{stamp}.xlsx"

    # 展示用列顺序
    cols = [
        "项目编码", "项目名称", "数据覆盖", "预算币种",
        "预算_人_万泰铢", "过程_人_万泰铢", "执行率_人", "超耗_人_万泰铢", "预算_人占比", "过程_人占比", "结构差_人_pct",
        "预算_材_万泰铢", "过程_材_万泰铢", "过程_采购_万泰铢", "过程_商混_万泰铢", "执行率_材", "超耗_材_万泰铢", "预算_材占比", "过程_材占比", "结构差_材_pct",
        "预算_机_万泰铢", "过程_机_万泰铢", "执行率_机", "超耗_机_万泰铢", "预算_机占比", "过程_机占比", "结构差_机_pct",
        "预算_人材机合计_万泰铢", "过程_人材机合计_万泰铢", "执行率_合计", "超耗_合计_万泰铢",
        "风险", "预警",
    ]
    export = df.copy()
    for c in cols:
        if c not in export.columns:
            export[c] = None
    export = export[cols]

    meta_df = pd.DataFrame([{"项": k, "值": v} for k, v in meta.items()])

    # 组合结构
    def sumcol(c):
        return float(export[c].map(sf).sum())

    portfolio = pd.DataFrame([
        {"维度": "人", "预算_万泰铢": sumcol("预算_人_万泰铢"), "过程_万泰铢": sumcol("过程_人_万泰铢")},
        {"维度": "材", "预算_万泰铢": sumcol("预算_材_万泰铢"), "过程_万泰铢": sumcol("过程_材_万泰铢")},
        {"维度": "机", "预算_万泰铢": sumcol("预算_机_万泰铢"), "过程_万泰铢": sumcol("过程_机_万泰铢")},
        {"维度": "合计", "预算_万泰铢": sumcol("预算_人材机合计_万泰铢"), "过程_万泰铢": sumcol("过程_人材机合计_万泰铢")},
    ])
    portfolio["执行率"] = portfolio.apply(
        lambda r: (r["过程_万泰铢"] / r["预算_万泰铢"]) if r["预算_万泰铢"] > 1e-9 else None, axis=1
    )
    portfolio["超耗_万泰铢"] = portfolio["过程_万泰铢"] - portfolio["预算_万泰铢"]

    for path in (latest, stamped):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            export.to_excel(w, sheet_name="项目对比明细", index=False)
            portfolio.to_excel(w, sheet_name="组合结构", index=False)
            meta_df.to_excel(w, sheet_name="生成信息", index=False)
            # 仅双边
            both = export[export["数据覆盖"] == "双边齐全"]
            both.to_excel(w, sheet_name="双边齐全项目", index=False)

    return latest, stamped


# ── HTML ───────────────────────────────────────────────────────────

def build_html(df: pd.DataFrame, meta: Dict, cfg: Dict) -> str:
    generated = meta.get("生成时间", "")
    month = meta.get("业务月", "")
    fx = meta.get("汇率_CNY_to_THB", "")

    records = []
    for _, r in df.iterrows():
        risk = risk_of_row(r, cfg)
        alerts = alerts_of_row(r, cfg)
        rec = {c: (None if (isinstance(r.get(c), float) and pd.isna(r.get(c))) else r.get(c)) for c in df.columns}
        rec["风险"] = risk
        rec["预警"] = alerts
        records.append(rec)

    both = [x for x in records if x.get("数据覆盖") == "双边齐全"]
    n_both = len(both)
    n_over = sum(1 for x in both if (x.get("执行率_合计") or 0) >= cfg.get("warn_exec_ratio", 1.15))

    def ssum(key, rows=None):
        rows = rows if rows is not None else both
        return sum(sf(x.get(key)) for x in rows)

    b_labor, p_labor = ssum("预算_人_万泰铢"), ssum("过程_人_万泰铢")
    b_mat, p_mat = ssum("预算_材_万泰铢"), ssum("过程_材_万泰铢")
    b_mach, p_mach = ssum("预算_机_万泰铢"), ssum("过程_机_万泰铢")
    b_tot, p_tot = ssum("预算_人材机合计_万泰铢"), ssum("过程_人材机合计_万泰铢")

    # chart data — top projects by budget among 双边
    top = sorted(both, key=lambda x: sf(x.get("预算_人材机合计_万泰铢")), reverse=True)[:18]
    chart_names = [x.get("项目名称") or x.get("项目编码") for x in top]
    chart_b = [round(sf(x.get("预算_人材机合计_万泰铢")), 1) for x in top]
    chart_p = [round(sf(x.get("过程_人材机合计_万泰铢")), 1) for x in top]
    chart_er = [round(sf(x.get("执行率_合计")) * 100, 1) if x.get("执行率_合计") is not None else None for x in top]

    def fmt_pct(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        return f"{sf(v)*100:.1f}%"

    def fmt_n(v, d=1):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        return f"{sf(v):,.{d}f}"

    rows_html = []
    for x in records:
        risk = x.get("风险") or "muted"
        er = x.get("执行率_合计")
        alerts = x.get("预警") or []
        tags = " ".join(f'<span class="tag">{esc(a)}</span>' for a in alerts) or '<span class="tag ok">—</span>'
        rows_html.append(f"""
<tr class="row risk-{esc(risk)}" data-cover="{esc(x.get('数据覆盖'))}" data-risk="{esc(risk)}"
    data-name="{esc(x.get('项目名称'))}" data-json="{esc(json.dumps(x, ensure_ascii=False, default=str))}">
  <td><span class="dot {esc(risk)}"></span>{esc(x.get('项目编码') or '—')}</td>
  <td><strong>{esc(x.get('项目名称') or '—')}</strong><div class="sub">{esc(x.get('数据覆盖'))} · 预算币种 {esc(x.get('预算币种') or '—')}</div></td>
  <td class="num">{fmt_n(x.get('预算_人_万泰铢'))}</td>
  <td class="num">{fmt_n(x.get('过程_人_万泰铢'))}</td>
  <td class="num">{fmt_pct(x.get('执行率_人'))}</td>
  <td class="num">{fmt_n(x.get('预算_材_万泰铢'))}</td>
  <td class="num">{fmt_n(x.get('过程_材_万泰铢'))}</td>
  <td class="num">{fmt_pct(x.get('执行率_材'))}</td>
  <td class="num">{fmt_n(x.get('预算_机_万泰铢'))}</td>
  <td class="num">{fmt_n(x.get('过程_机_万泰铢'))}</td>
  <td class="num">{fmt_pct(x.get('执行率_机'))}</td>
  <td class="num"><b>{fmt_n(x.get('预算_人材机合计_万泰铢'))}</b></td>
  <td class="num"><b>{fmt_n(x.get('过程_人材机合计_万泰铢'))}</b></td>
  <td class="num er-{esc(risk)}"><b>{fmt_pct(er)}</b></td>
  <td>{tags}</td>
</tr>""")

    src_block = "<br/>".join(f"· {esc(k)}: {esc(v)}" for k, v in meta.items() if str(k).startswith("源_"))

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>预算 vs 过程 · 人材机结构对比 · {esc(month)}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg:#f3f6f5; --card:#fff; --text:#1a2420; --muted:#6a7a72;
  --good:#1f8a65; --warn:#c4892a; --danger:#c2473d; --muted-risk:#8a9390;
  --line:#e2e9e5; --hero1:#163d5c; --hero2:#1f6f8b;
}}
* {{box-sizing:border-box;}}
body {{margin:0;font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;background:var(--bg);color:var(--text);}}
.wrap {{max-width:1500px;margin:0 auto;padding:18px;}}
.hero {{
  background:linear-gradient(135deg,var(--hero1),var(--hero2));color:#fff;
  border-radius:18px;padding:22px 26px;margin-bottom:14px;
  box-shadow:0 8px 24px rgba(22,61,92,.18);
}}
.hero h1 {{margin:0 0 6px;font-size:22px;}}
.hero p {{margin:3px 0;font-size:13px;opacity:.92;}}
.hero a {{color:#dff;}}
.kpis {{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px;}}
.kpi {{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px 14px;}}
.kpi .l {{font-size:12px;color:var(--muted);}}
.kpi .v {{font-size:20px;font-weight:700;margin-top:4px;}}
.kpi .u {{font-size:11px;color:var(--muted);}}
.panel {{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:14px 16px;margin-bottom:14px;}}
.panel h2 {{margin:0 0 10px;font-size:16px;}}
.charts {{display:grid;grid-template-columns:1.2fr 1fr;gap:12px;}}
@media (max-width:960px) {{.charts {{grid-template-columns:1fr;}}}}
.chart-box {{height:320px;background:#fafbfc;border:1px solid var(--line);border-radius:12px;padding:8px;}}
.filters {{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px;}}
.filters select,.filters input {{border:1px solid var(--line);border-radius:8px;padding:6px 10px;font-size:13px;}}
.table-scroll {{max-height:580px;overflow:auto;border:1px solid var(--line);border-radius:10px;}}
table {{width:100%;border-collapse:collapse;font-size:12px;}}
th,td {{padding:7px 6px;border-bottom:1px solid var(--line);text-align:left;}}
th {{position:sticky;top:0;background:#eef4f7;z-index:1;}}
.num {{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}}
.sub {{font-size:11px;color:var(--muted);margin-top:2px;}}
.row {{cursor:pointer;}} .row:hover {{background:#f0f7fa;}}
.dot {{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px;}}
.dot.good {{background:var(--good);}} .dot.warning {{background:var(--warn);}}
.dot.danger {{background:var(--danger);}} .dot.muted {{background:var(--muted-risk);}}
.er-danger {{color:var(--danger);}} .er-warning {{color:var(--warn);}} .er-good {{color:var(--good);}}
.tag {{display:inline-block;background:#eef2f0;border-radius:6px;padding:1px 6px;margin:1px;font-size:10.5px;}}
.tag.ok {{background:#e8f6ef;color:var(--good);}}
.note {{font-size:12px;color:var(--muted);line-height:1.55;margin-top:8px;}}
.mask {{
  position:fixed;inset:0;background:rgba(15,25,30,.48);display:none;z-index:40;
  align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(2px);
}}
.mask.show {{display:flex;}}
.drawer {{
  position:relative;width:min(720px,96vw);max-height:min(88vh,900px);background:#fff;
  border-radius:16px;box-shadow:0 20px 50px rgba(0,0,0,.22);
  overflow:auto;padding:20px 22px 28px;
  transform:translateY(12px) scale(.98);opacity:0;transition:transform .2s ease, opacity .2s ease;
  z-index:50;
}}
.mask.show .drawer {{transform:translateY(0) scale(1);opacity:1;}}
.drawer .close {{
  position:sticky;top:0;float:right;z-index:2;
  border:none;background:#eef2f0;border-radius:8px;padding:6px 12px;cursor:pointer;font-size:13px;
}}
.drawer .close:hover {{background:#e2e8e5;}}
.drawer h3 {{margin:0 40px 4px 0;padding-right:8px;}}
.bars {{margin:12px 0;}}
.bar-row {{display:grid;grid-template-columns:40px 1fr 70px;gap:8px;align-items:center;margin:6px 0;font-size:12px;}}
.track {{height:10px;background:#e8eef0;border-radius:5px;overflow:hidden;position:relative;}}
.track i {{display:block;height:100%;position:absolute;left:0;top:0;}}
.track i.b {{background:rgba(31,111,139,.45);}}
.track i.p {{background:rgba(31,138,101,.85);}}
.kv {{display:grid;grid-template-columns:1fr 1fr;gap:8px;}}
.kv .item {{background:#f7fafb;border:1px solid var(--line);border-radius:10px;padding:8px 10px;}}
.kv .t {{font-size:11px;color:var(--muted);}} .kv .n {{font-size:15px;font-weight:600;margin-top:2px;}}
.struct {{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <h1>预算 vs 过程 · 人材机结构对比</h1>
    <p>业务月标签：<b>{esc(month)}</b> ｜ 生成：{esc(generated)} ｜ CNY→THB 汇率：<b>{esc(fx)}</b></p>
    <p>预算=广联达人材机（人/材/机）；过程=人工费台账 + 采购台账 + 商混 + 机具租赁。单位默认 <b>万泰铢</b>。</p>
    <p>可与 <a href="全过程成本分析看板_最新.html">全过程成本分析看板</a> 联看。每月更新：重新跑 Python 即可覆盖「_最新」文件。</p>
  </div>

  <div class="kpis">
    <div class="kpi"><div class="l">双边齐全项目</div><div class="v">{n_both}</div><div class="u">全表 {len(records)} 行</div></div>
    <div class="kpi"><div class="l">合计执行率偏高</div><div class="v" style="color:var(--danger)">{n_over}</div><div class="u">过程/预算 ≥ {cfg.get('warn_exec_ratio',1.15):.0%}</div></div>
    <div class="kpi"><div class="l">预算人材机合计</div><div class="v">{b_tot:,.0f}</div><div class="u">万泰铢 · 仅双边</div></div>
    <div class="kpi"><div class="l">过程人材机合计</div><div class="v">{p_tot:,.0f}</div><div class="u">万泰铢 · 仅双边</div></div>
    <div class="kpi"><div class="l">人 预算→过程</div><div class="v">{b_labor:,.0f}→{p_labor:,.0f}</div><div class="u">执行 {(p_labor/b_labor*100 if b_labor else 0):.0f}%</div></div>
    <div class="kpi"><div class="l">材 预算→过程</div><div class="v">{b_mat:,.0f}→{p_mat:,.0f}</div><div class="u">执行 {(p_mat/b_mat*100 if b_mat else 0):.0f}% · 含采购+商混</div></div>
    <div class="kpi"><div class="l">机 预算→过程</div><div class="v">{b_mach:,.0f}→{p_mach:,.0f}</div><div class="u">执行 {(p_mach/b_mach*100 if b_mach else 0):.0f}%</div></div>
    <div class="kpi"><div class="l">组合超耗</div><div class="v">{(p_tot-b_tot):,.0f}</div><div class="u">过程 − 预算（万泰铢）</div></div>
  </div>

  <div class="panel">
    <h2>组合结构 &amp; 项目对比</h2>
    <div class="charts">
      <div class="chart-box"><canvas id="c1"></canvas></div>
      <div class="chart-box"><canvas id="c2"></canvas></div>
    </div>
  </div>

  <div class="panel">
    <h2>项目明细（点击行钻取结构）</h2>
    <div class="filters">
      <label>覆盖
        <select id="fCover">
          <option value="">全部</option>
          <option value="双边齐全" selected>双边齐全</option>
          <option value="仅预算">仅预算</option>
          <option value="仅过程">仅过程</option>
        </select>
      </label>
      <label>风险
        <select id="fRisk">
          <option value="">全部</option>
          <option value="danger">高风险</option>
          <option value="warning">需关注</option>
          <option value="good">健康</option>
          <option value="muted">数据缺口</option>
        </select>
      </label>
      <label>搜索 <input id="fSearch" placeholder="编码/名称"/></label>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th>编码</th><th>项目</th>
            <th class="num">预算人</th><th class="num">过程人</th><th class="num">人执行</th>
            <th class="num">预算材</th><th class="num">过程材</th><th class="num">材执行</th>
            <th class="num">预算机</th><th class="num">过程机</th><th class="num">机执行</th>
            <th class="num">预算合计</th><th class="num">过程合计</th><th class="num">合计执行</th>
            <th>预警</th>
          </tr>
        </thead>
        <tbody id="tbody">
          {''.join(rows_html)}
        </tbody>
      </table>
    </div>
    <div class="note">
      <b>口径说明</b><br/>
      · 执行率 = 过程金额 / 预算金额；&gt;100% 表示过程累计已超过预算该分项。<br/>
      · 过程「材」= 采购台账含税金额 + 搅拌站商混金额（均转万泰铢）。预算「材」= 广联达材类明细。<br/>
      · 预算 CNY 已按汇率换算为泰铢再÷10000；THB 预算直接÷10000。汇率可在 config 或 --fx 调整后重跑。<br/>
      · 「仅过程」项目无广联达预算文件；「仅预算」可能是新立项尚未进台账或名称未对齐。<br/>
      <b>数据源</b><br/>{src_block}
    </div>
  </div>
</div>

<div class="mask" id="mask" role="dialog" aria-modal="true">
  <div class="drawer" id="drawer" onclick="event.stopPropagation()">
    <button class="close" id="btnClose" type="button">关闭</button>
    <h3 id="dTitle">项目</h3>
    <div id="dSub" class="note" style="margin-top:0"></div>
    <div class="struct">
      <div>
        <h4>预算结构</h4>
        <canvas id="pieB" height="180"></canvas>
      </div>
      <div>
        <h4>过程结构</h4>
        <canvas id="pieP" height="180"></canvas>
      </div>
    </div>
    <div class="bars" id="dBars"></div>
    <div class="kv" id="dKv"></div>
    <h4>预警</h4>
    <div id="dAlerts"></div>
  </div>
</div>

<script>
const ROWS = {json.dumps(records, ensure_ascii=False, default=str)};
const chartNames = {json.dumps(chart_names, ensure_ascii=False)};
const chartB = {json.dumps(chart_b)};
const chartP = {json.dumps(chart_p)};
const chartEr = {json.dumps(chart_er)};
const portfolio = {{
  labels: ['人','材','机'],
  budget: [{b_labor:.2f},{b_mat:.2f},{b_mach:.2f}],
  process: [{p_labor:.2f},{p_mat:.2f},{p_mach:.2f}],
}};

new Chart(document.getElementById('c1'), {{
  type: 'bar',
  data: {{
    labels: chartNames,
    datasets: [
      {{ label: '预算合计(万泰铢)', data: chartB, backgroundColor: 'rgba(31,111,139,.55)' }},
      {{ label: '过程合计(万泰铢)', data: chartP, backgroundColor: 'rgba(31,138,101,.7)' }},
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ title: {{ display: true, text: '项目预算 vs 过程（人材机合计，按预算 Top）' }}, legend: {{ position: 'bottom' }} }},
    scales: {{ x: {{ ticks: {{ maxRotation: 60, minRotation: 30, font: {{ size: 10 }} }} }} }}
  }}
}});

new Chart(document.getElementById('c2'), {{
  type: 'bar',
  data: {{
    labels: portfolio.labels,
    datasets: [
      {{ label: '预算', data: portfolio.budget, backgroundColor: 'rgba(31,111,139,.55)' }},
      {{ label: '过程', data: portfolio.process, backgroundColor: 'rgba(31,138,101,.7)' }},
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ title: {{ display: true, text: '双边齐全项目 · 组合人/材/机结构' }}, legend: {{ position: 'bottom' }} }}
  }}
}});

function applyFilter() {{
  const cover = document.getElementById('fCover').value;
  const risk = document.getElementById('fRisk').value;
  const q = document.getElementById('fSearch').value.trim().toLowerCase();
  document.querySelectorAll('#tbody tr').forEach(tr => {{
    let ok = true;
    if (cover && tr.dataset.cover !== cover) ok = false;
    if (risk && tr.dataset.risk !== risk) ok = false;
    if (q) {{
      const t = (tr.dataset.name + ' ' + tr.children[0].innerText).toLowerCase();
      if (!t.includes(q)) ok = false;
    }}
    tr.style.display = ok ? '' : 'none';
  }});
}}
['fCover','fRisk','fSearch'].forEach(id => document.getElementById(id).addEventListener('input', applyFilter));
applyFilter();

const mask = document.getElementById('mask');
const drawer = document.getElementById('drawer');
let pieB, pieP;
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

function fmt(n, d=1) {{
  if (n === null || n === undefined || n === '') return '—';
  const x = Number(n); if (Number.isNaN(x)) return '—';
  return x.toLocaleString('zh-CN', {{ maximumFractionDigits: d, minimumFractionDigits: d }});
}}
function pct(n) {{
  if (n === null || n === undefined) return '—';
  return (Number(n)*100).toFixed(1) + '%';
}}

function openRow(p) {{
  document.getElementById('dTitle').textContent = (p['项目编码'] ? p['项目编码'] + ' · ' : '') + (p['项目名称'] || '');
  document.getElementById('dSub').textContent = (p['数据覆盖'] || '') + ' ｜ 预算币种 ' + (p['预算币种'] || '—');
  const bb = [p['预算_人_万泰铢']||0, p['预算_材_万泰铢']||0, p['预算_机_万泰铢']||0];
  const pp = [p['过程_人_万泰铢']||0, p['过程_材_万泰铢']||0, p['过程_机_万泰铢']||0];
  if (pieB) pieB.destroy(); if (pieP) pieP.destroy();
  pieB = new Chart(document.getElementById('pieB'), {{
    type: 'doughnut',
    data: {{ labels: ['人','材','机'], datasets: [{{ data: bb, backgroundColor: ['#3a7bd5','#1f8a65','#c4892a'] }}] }},
    options: {{ plugins: {{ legend: {{ position: 'bottom' }} }} }}
  }});
  pieP = new Chart(document.getElementById('pieP'), {{
    type: 'doughnut',
    data: {{ labels: ['人','材','机'], datasets: [{{ data: pp, backgroundColor: ['#3a7bd5','#1f8a65','#c4892a'] }}] }},
    options: {{ plugins: {{ legend: {{ position: 'bottom' }} }} }}
  }});

  const maxv = Math.max(...bb, ...pp, 1);
  const labs = ['人','材','机'];
  let bars = '';
  for (let i=0;i<3;i++) {{
    const bw = (bb[i]/maxv*100).toFixed(1);
    const pw = (pp[i]/maxv*100).toFixed(1);
    bars += `<div class="bar-row"><span>${{labs[i]}}</span><div class="track"><i class="b" style="width:${{bw}}%"></i><i class="p" style="width:${{pw}}%;opacity:.85"></i></div><span>${{fmt(bb[i])}} / ${{fmt(pp[i])}}</span></div>`;
  }}
  document.getElementById('dBars').innerHTML = '<div class="note">条形：浅蓝=预算，绿色=过程（同尺度）</div>' + bars;

  document.getElementById('dKv').innerHTML = `
    <div class="item"><div class="t">合计执行率</div><div class="n">${{pct(p['执行率_合计'])}}</div></div>
    <div class="item"><div class="t">合计超耗</div><div class="n">${{fmt(p['超耗_合计_万泰铢'])}} 万泰铢</div></div>
    <div class="item"><div class="t">人/材/机 执行率</div><div class="n">${{pct(p['执行率_人'])}} / ${{pct(p['执行率_材'])}} / ${{pct(p['执行率_机'])}}</div></div>
    <div class="item"><div class="t">结构差 人/材/机 (pct)</div><div class="n">${{p['结构差_人_pct']==null?'—':Number(p['结构差_人_pct']).toFixed(1)}} / ${{p['结构差_材_pct']==null?'—':Number(p['结构差_材_pct']).toFixed(1)}} / ${{p['结构差_机_pct']==null?'—':Number(p['结构差_机_pct']).toFixed(1)}}</div></div>
    <div class="item"><div class="t">过程采购 / 商混</div><div class="n">${{fmt(p['过程_采购_万泰铢'])}} / ${{fmt(p['过程_商混_万泰铢'])}}</div></div>
    <div class="item"><div class="t">预算管+其他</div><div class="n">${{fmt(p['预算_管_万泰铢'])}} / ${{fmt(p['预算_其他_万泰铢'])}}</div></div>
  `;
  const alerts = (p['预警'] && p['预警'].length) ? p['预警'].map(a => `<span class="tag">${{a}}</span>`).join(' ') : '<span class="tag ok">无</span>';
  document.getElementById('dAlerts').innerHTML = alerts;
  openDrawer();
}}

document.querySelectorAll('#tbody tr').forEach(tr => {{
  tr.addEventListener('click', () => {{
    try {{ openRow(JSON.parse(tr.getAttribute('data-json'))); }} catch(e) {{ console.error(e); }}
  }});
}});
</script>
</body>
</html>
"""


# ── main ───────────────────────────────────────────────────────────

def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="预算 vs 过程 · 人材机结构对比（可每月重跑）")
    parser.add_argument("--fx", type=float, default=None, help="CNY→THB 汇率，覆盖 config")
    parser.add_argument("--month", default=None, help="业务月标签，如 2026-06 或 2026年06月")
    parser.add_argument("--output-dir", default=None, help="输出目录，默认 0_html")
    parser.add_argument("--no-html", action="store_true", help="只生成 Excel")
    parser.add_argument("--open", action="store_true", help="生成后尝试打开 HTML")
    args = parser.parse_args(argv)

    cfg = load_or_init_config()
    fx = args.fx if args.fx is not None else float(cfg.get("fx_cny_to_thb", DEFAULT_FX_CNY_TO_THB))
    out_dir = Path(args.output_dir) if args.output_dir else html_data_path()
    out_dir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    month_label = args.month or datetime.now().strftime("%Y年%m月")

    print(f"[1/5] 加载预算人材机 (fx={fx}) …")
    budget, src_b = load_budget(fx)
    print(f"      预算项目 {len(budget)} | {src_b}")

    print("[2/5] 加载过程人工 …")
    labor, src_l = load_labor()
    print(f"      人工 {len(labor)} | {src_l}")

    print("[3/5] 加载过程材料（采购+商混）…")
    materials, src_m = load_materials()
    print(f"      材料 {len(materials)} | {src_m}")

    print("[4/5] 加载过程机具 …")
    machines, src_h = load_machines()
    print(f"      机具 {len(machines)} | {src_h}")

    print("[5/5] 合并指标并导出 …")
    df = merge_all(budget, labor, materials, machines)
    if df.empty:
        print("[ERR] 无可用数据", file=sys.stderr)
        return 1

    # risk columns for excel
    df["风险"] = df.apply(lambda r: risk_of_row(r, cfg), axis=1)
    df["预警"] = df.apply(lambda r: "；".join(alerts_of_row(r, cfg)), axis=1)

    meta = {
        "业务月": month_label,
        "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "汇率_CNY_to_THB": fx,
        "配置文件": str(CONFIG_PATH),
        "源_预算": src_b,
        "源_人工": src_l,
        "源_材料": src_m,
        "源_机具": src_h,
        "行数": len(df),
        "双边齐全": int((df["数据覆盖"] == "双边齐全").sum()),
        "脚本": str(Path(__file__).resolve()),
    }

    latest_xlsx, stamped_xlsx = export_excel(df, meta, out_dir, stamp)
    print(f"      Excel: {latest_xlsx}")
    print(f"      Excel(归档): {stamped_xlsx}")

    html_path = out_dir / "预算vs过程_人材机对比_最新.html"
    if not args.no_html:
        html = build_html(df, meta, cfg)
        html_path.write_text(html, encoding="utf-8")
        # also stamped copy
        (out_dir / f"预算vs过程_人材机对比_{stamp}.html").write_text(html, encoding="utf-8")
        print(f"      HTML: {html_path}")

    # soft-link note into full lifecycle page if exists
    life = out_dir / "全过程成本分析看板_最新.html"
    if life.exists() and not args.no_html:
        try:
            text = life.read_text(encoding="utf-8")
            needle = "预算vs过程_人材机对比_最新.html"
            if needle not in text:
                text = text.replace(
                    "设计说明见：Desktop/code/S_数据分析模块/全过程成本分析看板_设计说明.md",
                    "设计说明见：Desktop/code/S_数据分析模块/全过程成本分析看板_设计说明.md<br/>"
                    f'· <a href="{needle}" target="_blank">预算 vs 过程 · 人材机结构对比</a>',
                )
                life.write_text(text, encoding="utf-8")
                print("      已在全过程看板底部增加对比页链接")
        except Exception as e:
            print(f"      [warn] 未能写入全过程看板链接: {e}")

    n_both = int((df["数据覆盖"] == "双边齐全").sum())
    n_over = int(df.loc[df["数据覆盖"] == "双边齐全", "执行率_合计"].map(
        lambda x: False if x is None or (isinstance(x, float) and pd.isna(x)) else x >= cfg.get("warn_exec_ratio", 1.15)
    ).sum()) if n_both else 0
    print(f"[OK] 完成 | 双边齐全 {n_both} | 执行率偏高 {n_over}")
    print("     每月更新：更新 raw/std 后执行 python build_budget_vs_process.py")
    print(f"     汇率配置：{CONFIG_PATH}")

    if args.open and html_path.exists():
        import subprocess
        subprocess.run(["open", str(html_path)], check=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
