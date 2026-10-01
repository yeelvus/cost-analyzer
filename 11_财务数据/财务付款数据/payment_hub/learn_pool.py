# -*- coding: utf-8 -*-
"""
D_财务付款_学习池（02_Std/12_财务付款/D_财务付款_学习池）

- 00_汇总/全部摘要_学习池.xlsx + 00_处理优先级_总览.xlsx
- 01_按项目/{优先级}_{编码}_{名称}.xlsx  （有正式项目编码，按付款金额大→小）
- 01b_待核对_无项目编码/待核对_无项目编码汇总.xlsx
- 02_按月份/{YYYY-MM}_财务数据.xlsx
- 03_已确认/已确认分类.xlsx

摘要去重；行内按「应付金额占比」优先排序（大额先确认，控比例误差）。
手改或确认后 apply → 按业务摘要写记忆 + 剥离 + 重写 A2 并重建 B_std。
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .category_dropdowns import apply_category_dropdowns
from .config import Paths
from .pkg import ensure_code_root
from .utils import (
    STD_SPEND_BUCKETS,
    has_actual_payment,
    parse_std_project_filename,
    safe_filename,
    safe_text,
)

POOL_SHEET = "学习池"
STATUS_PENDING = "待确认"
STATUS_CONFIRMED = "确认"
STATUS_VALUES = [STATUS_PENDING, STATUS_CONFIRMED, "已确认", "是", "Y", "y"]

# 核心列（apply 依赖）+ 金额优先级列（展示/排序，不影响确认逻辑）
COLUMNS = [
    "确认状态",
    "处理优先级",
    "摘要",
    "一级科目",
    "二级科目",
    "三级科目",
    "供应商名称",
    "出现次数",
    "应付金额合计",
    "金额占比%",
    "累计占比%",
    "来源分组",
    "支出归类",
]

# 00_汇总 四象限文件名模板（有/无编码 × 来源分组）
# 实际文件名在导出时按 std 中出现的业务分组动态生成；常见为 流水/银行账 → 4 个文件
PERCENT_NUMBER_FORMAT = '0.00"%"'  # 数值 0.21 显示为 0.21%

SNAPSHOT_META = "_meta"

# 无正式编码归堆时的显示名
NO_CODE_LABEL = "无项目编码"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _is_confirmed(val: Any) -> bool:
    s = safe_text(val)
    if not s:
        return False
    s2 = s.replace(" ", "").replace("　", "")
    return s2 in {
        STATUS_CONFIRMED,
        "已确认",
        "是",
        "Y",
        "y",
        "1",
        "true",
        "True",
        "OK",
        "ok",
        "√",
        "✓",
    } or s2.startswith("确认")


def _norm_cat(v: Any) -> str:
    s = safe_text(v)
    if s in {"(空白)", "-", "—"}:
        return ""
    return s


def _triple(l1: Any, l2: Any, l3: Any) -> Tuple[str, str, str]:
    return _norm_cat(l1), _norm_cat(l2), _norm_cat(l3)


def _month_from_date(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "未知月份"
    if isinstance(val, datetime):
        return val.strftime("%Y-%m")
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts):
            return "未知月份"
        return ts.strftime("%Y-%m")
    except Exception:
        s = safe_text(val)
        m = re.match(r"(20\d{2})[-/]?(\d{1,2})", s)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}"
        return "未知月份"


def _pick_sheet(path: Path) -> str:
    try:
        xls = pd.ExcelFile(path)
        if "明细数据" in xls.sheet_names:
            return "明细数据"
        return xls.sheet_names[0]
    except Exception:
        return "Sheet1"


def load_confirmed_summaries(confirmed_path: Path) -> Dict[str, Dict[str, str]]:
    """摘要 -> {一级,二级,三级,供应商,确认状态,...}"""
    path = Path(confirmed_path)
    if not path.is_file():
        return {}
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception:
        return {}
    out: Dict[str, Dict[str, str]] = {}
    for _, row in df.iterrows():
        summ = safe_text(row.get("摘要"))
        if not summ:
            continue
        out[summ] = {
            "确认状态": STATUS_CONFIRMED,
            "摘要": summ,
            "一级科目": _norm_cat(row.get("一级科目")),
            "二级科目": _norm_cat(row.get("二级科目")),
            "三级科目": _norm_cat(row.get("三级科目")),
            "供应商名称": safe_text(row.get("供应商名称")),
            "出现次数": str(row.get("出现次数") or ""),
            "确认时间": safe_text(row.get("确认时间")) or _now(),
            "来源": safe_text(row.get("来源")),
        }
    return out


def load_snapshot(path: Path) -> Dict[str, Dict[str, Dict[str, str]]]:
    """
    返回: { 相对路径或文件名: { 摘要: {一级,二级,三级} } }
    兼容旧版全局 summaries 映射（视为 __global__）。
    """
    p = Path(path)
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data.get("files"), dict):
            out: Dict[str, Dict[str, Dict[str, str]]] = {}
            for fk, mp in data["files"].items():
                if isinstance(mp, dict):
                    out[str(fk)] = {
                        str(sk): sv for sk, sv in mp.items() if isinstance(sv, dict)
                    }
            return out
        # 旧格式
        rows = data.get("summaries") or {}
        if isinstance(rows, dict):
            return {
                "__global__": {
                    str(k): v for k, v in rows.items() if isinstance(v, dict)
                }
            }
        return {}
    except Exception:
        return {}


def save_snapshot(path: Path, files_map: Dict[str, Dict[str, Dict[str, str]]]) -> None:
    """files_map: {文件名: {摘要: {一级,二级,三级}}}"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    n = sum(len(v) for v in files_map.values())
    payload = {
        "exported_at": _now(),
        "file_count": len(files_map),
        "row_count": n,
        "files": files_map,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def is_thcc_project_code(code: str) -> bool:
    """是否为 THCC 正式项目编码（用于按月份池过滤）。"""
    s = safe_text(code)
    if not s:
        return False
    return bool(re.match(r"^THCC-", s, re.I))


def has_formal_project_code(code: str) -> bool:
    """
    是否具备可核对的正式项目编码（进入 01_按项目 优先处理）。

    无编码 / 未识别编码 / 需手动核对 → False，归入 01b 待核对。
    识别：THCC- / LSBM- / XCY- 等「字母+编号」形态。
    """
    s = safe_text(code)
    if not s:
        return False
    low = s.lower()
    if any(
        k in low
        for k in (
            "未识别",
            "需手动",
            "无编码",
            "未知",
            "nan",
            "none",
            "null",
        )
    ):
        return False
    if re.match(r"^[A-Za-z]{2,}[-_]\d", s):
        return True
    if re.match(r"^[A-Za-z]{2,}\d", s):
        return True
    return False


def _line_payable(row: Any) -> float:
    """行金额：优先应付；否则预扣+实付；再否则实付。"""

    def _num(key: str) -> float:
        try:
            if hasattr(row, "index") and key not in row.index:
                return 0.0
            v = row[key] if hasattr(row, "index") else row.get(key)
            n = pd.to_numeric(v, errors="coerce")
            return float(n) if pd.notna(n) else 0.0
        except Exception:
            return 0.0

    pay = _num("应付金额")
    if abs(pay) > 1e-12:
        return pay
    combo = _num("预扣税额") + _num("实付金额")
    if abs(combo) > 1e-12:
        return combo
    return _num("实付金额")


def _std_path_meta(file_path: Path, std_dir: Path) -> Tuple[str, str]:
    """
    从 B_std 路径解析 (来源分组, 支出归类)。

    新结构：B_std/{流水|银行账}/最新月份有支出|未支出/xxx.xlsx
    旧结构：B_std/最新月份有支出|未支出/xxx.xlsx → 来源分组=未分组
    """
    try:
        parts = list(Path(file_path).resolve().relative_to(Path(std_dir).resolve()).parts)
    except ValueError:
        parts = list(Path(file_path).parts)
    dirs = [p for p in parts[:-1] if p and p not in {".", ".."}]
    spend = ""
    group = "未分组"
    for p in dirs:
        if p in STD_SPEND_BUCKETS:
            spend = p
        else:
            group = p
    if not spend:
        spend = "未归类支出"
    return group, spend


def collect_raw_lines_from_std(std_dir: Path) -> List[Dict[str, Any]]:
    """从 B_std 收集明细行。"""
    ensure_code_root()
    from summary_recognize import recognize_summary

    std_dir = Path(std_dir)
    lines: List[Dict[str, Any]] = []
    for f in sorted(std_dir.rglob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        sheet = _pick_sheet(f)
        try:
            df = pd.read_excel(f, sheet_name=sheet)
        except Exception as e:
            print(f"  ⚠️ 跳过 {f.name}: {e}")
            continue
        if df.empty:
            continue
        # 文件名兜底项目（编码_名称_财务付款数据 / 旧编码_名称_(日期)）
        file_code, file_name = parse_std_project_filename(f)
        src_group, spend_bucket = _std_path_meta(f, std_dir)

        for _, row in df.iterrows():
            raw_summ = safe_text(row.get("摘要"))
            biz_summ = safe_text(row.get("业务摘要"))
            if not raw_summ and not biz_summ:
                continue
            # 仅学习「有实付金额」的摘要；空实付不进入学习池
            if not has_actual_payment(row.get("实付金额")):
                continue
            code = safe_text(row.get("项目编码")) or file_code or "未识别编码"
            pname = safe_text(row.get("项目名称")) or file_name or "未命名项目"
            if biz_summ:
                summ = biz_summ
            else:
                summ = recognize_summary(raw_summ, pname).business_summary or raw_summ
            if not summ:
                continue
            amt = pd.to_numeric(row.get("实付金额"), errors="coerce")
            payable = _line_payable(row)
            lines.append(
                {
                    "摘要": summ,
                    "一级科目": _norm_cat(row.get("一级科目")),
                    "二级科目": _norm_cat(row.get("二级科目")),
                    "三级科目": _norm_cat(row.get("三级科目")),
                    "供应商名称": safe_text(row.get("供应商名称")),
                    "项目编码": code,
                    "项目名称": pname,
                    "月份": _month_from_date(row.get("日期")),
                    "实付金额": float(amt) if pd.notna(amt) else 0.0,
                    "应付金额": float(payable),
                    "is_thcc": is_thcc_project_code(code),
                    "has_code": has_formal_project_code(code),
                    "来源分组": src_group,
                    "支出归类": spend_bucket,
                }
            )
    return lines


def _dedupe_rows(
    rows: List[Dict[str, Any]],
    *,
    sort_by_amount: bool = True,
) -> List[Dict[str, Any]]:
    """
    按摘要去重：众数科目/供应商，累计出现次数与应付金额。

    sort_by_amount=True（默认）：按应付金额合计降序，并写处理优先级/占比。
    大额摘要优先确认 → 误分对整体金额比例影响更小。
    """
    buckets: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in rows:
        buckets[r["摘要"]].append(r)

    out: List[Dict[str, Any]] = []
    for summ, items in buckets.items():
        l1s = Counter(_norm_cat(i["一级科目"]) for i in items)
        l2s = Counter(_norm_cat(i["二级科目"]) for i in items)
        l3s = Counter(_norm_cat(i["三级科目"]) for i in items)
        vendors = Counter(safe_text(i["供应商名称"]) for i in items if safe_text(i["供应商名称"]))
        l1 = l1s.most_common(1)[0][0] if l1s else ""
        l2 = l2s.most_common(1)[0][0] if l2s else ""
        l3 = l3s.most_common(1)[0][0] if l3s else ""
        vend = vendors.most_common(1)[0][0] if vendors else ""
        payable_sum = sum(float(i.get("应付金额") or 0) for i in items)
        groups = Counter(safe_text(i.get("来源分组")) or "未分组" for i in items)
        spends = Counter(safe_text(i.get("支出归类")) or "" for i in items)
        src_g = groups.most_common(1)[0][0] if groups else ""
        spend_g = spends.most_common(1)[0][0] if spends else ""
        out.append(
            {
                "确认状态": STATUS_PENDING,
                "处理优先级": 0,
                "摘要": summ,
                "一级科目": l1,
                "二级科目": l2,
                "三级科目": l3,
                "供应商名称": vend,
                "出现次数": len(items),
                "应付金额合计": round(payable_sum, 4),
                "金额占比%": 0.0,
                "累计占比%": 0.0,
                "来源分组": src_g,
                "支出归类": spend_g,
            }
        )

    total = sum(float(r.get("应付金额合计") or 0) for r in out)
    if sort_by_amount:
        out.sort(
            key=lambda r: (
                -float(r.get("应付金额合计") or 0),
                -int(r.get("出现次数") or 0),
                safe_text(r.get("摘要")),
            )
        )
    else:
        out = _sort_by_category(out)

    cum = 0.0
    for i, r in enumerate(out, start=1):
        amt = float(r.get("应付金额合计") or 0)
        share = (100.0 * amt / total) if total > 1e-12 else 0.0
        cum += share
        r["处理优先级"] = i
        r["金额占比%"] = round(share, 2)
        r["累计占比%"] = round(min(cum, 100.0), 2)
    return out


def _sort_by_category(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """一级 → 二级 → 三级 → 摘要 排序。"""
    blank = "\uffff"  # 空科目排最后

    def key(r: Dict[str, Any]):
        return (
            _norm_cat(r.get("一级科目")) or blank,
            _norm_cat(r.get("二级科目")) or blank,
            _norm_cat(r.get("三级科目")) or blank,
            safe_text(r.get("摘要")),
        )

    return sorted(rows, key=key)


def _project_payable_total(items: List[Dict[str, Any]]) -> float:
    return sum(float(i.get("应付金额") or 0) for i in items)


def _amt_value(v: Any) -> float:
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _resort_pool_rows_by_amount(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """按应付金额合计重算处理优先级 / 占比 / 累计占比。"""
    if not rows:
        return rows
    out = list(rows)
    total = sum(_amt_value(r.get("应付金额合计")) for r in out)
    out.sort(
        key=lambda x: (
            -_amt_value(x.get("应付金额合计")),
            -int(float(x.get("出现次数") or 0) or 0)
            if str(x.get("出现次数") or "").replace(".", "", 1).isdigit()
            else 0,
            safe_text(x.get("摘要")),
        )
    )
    cum = 0.0
    for i, r in enumerate(out, start=1):
        a = _amt_value(r.get("应付金额合计"))
        share = (100.0 * a / total) if total > 1e-12 else 0.0
        cum += share
        r["处理优先级"] = i
        r["金额占比%"] = round(share, 2)
        r["累计占比%"] = round(min(cum, 100.0), 2)
    return out


def _parse_project_pool_filename(stem: str) -> Tuple[int, str, str]:
    """解析 01_THCC-xxx_项目名 → (优先级, 编码, 名称)。"""
    stem = safe_text(stem)
    rank = 9999
    rest = stem
    m = re.match(r"^(\d+)_(.+)$", stem)
    if m:
        rank = int(m.group(1))
        rest = m.group(2)
    m2 = re.match(r"^([A-Za-z0-9\-]+)[_\-\s]+(.+)$", rest)
    if m2:
        return rank, safe_text(m2.group(1)), safe_text(m2.group(2))
    return rank, "", rest


def rebuild_priority_overview_from_disk(paths: Paths) -> Dict[str, Any]:
    """
    根据当前磁盘上的 01/01b/00 学习池重建「00_处理优先级_总览」。
    apply 回写后调用，避免总览丢失、结构「跑一次就没了」。
    """
    project_rank: List[Dict[str, Any]] = []
    if paths.learn_pool_by_project.is_dir():
        files = [
            f
            for f in paths.learn_pool_by_project.glob("*.xlsx")
            if not f.name.startswith("~$")
        ]
        parsed = []
        for f in files:
            rank, code, pname = _parse_project_pool_filename(f.stem)
            rows = _read_pool_file(f)
            amt = sum(_amt_value(r.get("应付金额合计")) for r in rows)
            parsed.append((rank, code, pname, f.name, len(rows), amt))
        # 有序号按序号；否则按金额
        if any(p[0] < 9000 for p in parsed):
            parsed.sort(key=lambda x: (x[0], -x[5]))
        else:
            parsed.sort(key=lambda x: -x[5])
            parsed = [
                (i, code, pname, fn, n, amt)
                for i, (_, code, pname, fn, n, amt) in enumerate(parsed, start=1)
            ]
        grand = sum(p[5] for p in parsed) or 0.0
        for rank, code, pname, fname, n_rows, amt in parsed:
            share = (100.0 * amt / grand) if grand > 1e-12 else 0.0
            project_rank.append(
                {
                    "项目优先级": rank,
                    "项目编码": code,
                    "项目名称": pname,
                    "未确认摘要数": n_rows,
                    "应付金额合计": round(amt, 2),
                    "占有编码金额%": round(share, 2),
                    "学习池文件": fname,
                }
            )

    global_summaries: List[Dict[str, Any]] = []
    if paths.learn_pool_all.is_dir():
        for f in sorted(paths.learn_pool_all.glob("*_学习池.xlsx")):
            if f.name.startswith("~$") or "处理优先级" in f.name:
                continue
            global_summaries.extend(_read_pool_file(f))
        # 同摘要合并：取金额最大
        by_s: Dict[str, Dict[str, Any]] = {}
        for r in global_summaries:
            s = safe_text(r.get("摘要"))
            if not s or s.startswith("（本分类暂无"):
                continue
            prev = by_s.get(s)
            if prev is None or _amt_value(r.get("应付金额合计")) > _amt_value(
                prev.get("应付金额合计")
            ):
                by_s[s] = r
        global_summaries = _resort_pool_rows_by_amount(list(by_s.values()))
    elif paths.learn_pool_all_xlsx.is_file():
        global_summaries = _read_pool_file(paths.learn_pool_all_xlsx)

    no_code_projects: List[Dict[str, Any]] = []
    if paths.learn_pool_no_code_xlsx.is_file():
        rows = _read_pool_file(paths.learn_pool_no_code_xlsx)
        no_code_projects.append(
            {
                "项目编码": NO_CODE_LABEL,
                "项目名称": "合并汇总",
                "明细行数": len(rows),
                "应付金额合计": round(
                    sum(_amt_value(r.get("应付金额合计")) for r in rows), 2
                ),
                "说明": "见 01b_待核对_无项目编码/待核对_无项目编码汇总.xlsx",
            }
        )

    _write_priority_overview(
        paths.learn_pool_priority_xlsx,
        project_rank=project_rank,
        global_summaries=global_summaries,
        no_code_projects=no_code_projects,
        latest_note=(
            f"重建总览 · 有编码文件 {len(project_rank)} · "
            f"全量摘要 {len(global_summaries)}"
        ),
    )
    return {
        "project_files": len(project_rank),
        "global_summaries": len(global_summaries),
        "priority_xlsx": str(paths.learn_pool_priority_xlsx),
    }


def _format_pool_workbook(
    path: Path,
    rows: List[Dict[str, Any]],
    data_dir: Path,
    *,
    title: str = "",
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = POOL_SHEET

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    for c, h in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, row in enumerate(rows, start=2):
        for c, h in enumerate(COLUMNS, start=1):
            val = row.get(h, "")
            if h == "确认状态" and not val:
                val = STATUS_PENDING
            ws.cell(row=i, column=c, value=val if val != "" else None)

    last = max(2, len(rows) + 1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"
    widths = [10, 10, 36, 14, 16, 14, 20, 10, 14, 12, 12, 12, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # 金额列数字格式；占比显示为 0.21% 这种带百分号样式（存的是百分点 0.21）
    for c, h in enumerate(COLUMNS, start=1):
        if h == "应付金额合计":
            for r in range(2, last + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
        elif h in {"金额占比%", "累计占比%"}:
            for r in range(2, last + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = PERCENT_NUMBER_FORMAT

    # 确认状态下拉
    dv = DataValidation(
        type="list",
        formula1=f'"{STATUS_PENDING},{STATUS_CONFIRMED}"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    dv.add(f"A2:A{last + 200}")
    ws.add_data_validation(dv)

    # 一/二/三级下拉
    try:
        apply_category_dropdowns(wb, data_dir=data_dir, sheet_names=[POOL_SHEET])
    except Exception as e:
        print(f"  ⚠️ 下拉写入失败 {path.name}: {e}")

    if title:
        # 说明放最后，打开文件默认先看到「学习池」数据
        note = wb.create_sheet("说明")
        note["A1"] = title
        note["A2"] = "① 按「处理优先级」从上往下确认：金额占比越大越靠前，先搞定大头再管零碎"
        note["A3"] = "② 改「一级/二级/三级」，或把「确认状态」改为「确认」后保存"
        note["A4"] = "③ 必须运行：python 财务付款一体化.py apply"
        note["A5"] = "④ 写入 03_已确认、剥离已确认摘要，并泛化抬升其余未确认行的预分类"
        note["A6"] = "⑤ 推荐：先做 01_按项目（金额大的项目优先）→ 再扫 00_汇总；无编码在 01b 单独核对"
        note["A7"] = "⑥ 不要只运行 pool 重导（可能覆盖未 apply 的修改）"
        note["A8"] = "⑦ 金额列仅辅助优先级，不改变 apply/记忆学习逻辑"
        note["A9"] = "⑧ 累计占比%：从第1名累加到当前行的金额占比，可快速看到「确认前N条覆盖多少钱」"
        note["A10"] = "⑨ 可填新二级/一级名：apply 默认自动登记进 taxonomy（先备份）；--no-auto-taxonomy 可关"
        note["A11"] = "⑩ 待确认但已改科目也会学习；仅确认才学请用 apply --only-confirmed"
        note.column_dimensions["A"].width = 96

    # 保证活动表是学习池
    try:
        wb.active = wb[POOL_SHEET]
    except Exception:
        pass

    wb.save(path)


def _write_priority_overview(
    path: Path,
    *,
    project_rank: List[Dict[str, Any]],
    global_summaries: List[Dict[str, Any]],
    no_code_projects: List[Dict[str, Any]],
    latest_note: str = "",
) -> None:
    """综合：项目优先级 + 全量摘要优先级 + 无编码清单。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    proj_df = pd.DataFrame(project_rank) if project_rank else pd.DataFrame()
    sum_df = pd.DataFrame(global_summaries) if global_summaries else pd.DataFrame()
    # 总览只要前若干大额摘要便于决策（完整列表在全部摘要表）
    if not sum_df.empty and len(sum_df) > 300:
        sum_df = sum_df.head(300).copy()
    nocode_df = pd.DataFrame(no_code_projects) if no_code_projects else pd.DataFrame()

    guide = pd.DataFrame(
        [
            {
                "建议顺序": 1,
                "动作": "打开 00_处理优先级_总览 →「项目优先级」",
                "说明": "付款金额大的正式编码项目排前，先处理这些项目的学习池",
            },
            {
                "建议顺序": 2,
                "动作": "进入 01_按项目 对应文件，按处理优先级改科目/点确认",
                "说明": "文件内摘要已按应付金额占比降序；累计占比%可看到覆盖了多少金额",
            },
            {
                "建议顺序": 3,
                "动作": "扫 00_汇总 四个表：有编码/无编码 × 流水/银行账",
                "说明": "按来源分组与是否有编码拆分；表内金额占比/累计占比显示为 0.21% 样式",
            },
            {
                "建议顺序": 4,
                "动作": "01b_待核对_无项目编码 单独核对",
                "说明": "无正式编码项目合并在此，不冲淡正式项目的优先级队列",
            },
            {
                "建议顺序": 5,
                "动作": "python 财务付款一体化.py apply",
                "说明": "写入记忆并剥离已确认；分类学习逻辑与原先一致",
            },
            {
                "建议顺序": 0,
                "动作": latest_note or "金额口径：优先应付金额；缺省时用预扣+实付",
                "说明": "金额列仅用于排序与占比，不改变 apply/记忆规则",
            },
        ]
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        guide.to_excel(writer, sheet_name="使用指南", index=False)
        if not proj_df.empty:
            proj_df.to_excel(writer, sheet_name="项目优先级", index=False)
        else:
            pd.DataFrame([{"提示": "无正式编码项目可导出"}]).to_excel(
                writer, sheet_name="项目优先级", index=False
            )
        if not sum_df.empty:
            sum_df.to_excel(writer, sheet_name="摘要优先级_全量Top", index=False)
        if not nocode_df.empty:
            nocode_df.to_excel(writer, sheet_name="无编码待核对清单", index=False)
        else:
            pd.DataFrame([{"提示": "无「无编码」项目"}]).to_excel(
                writer, sheet_name="无编码待核对清单", index=False
            )


def export_learn_pool(paths: Paths) -> Dict[str, Any]:
    """
    从 B_std 导出学习池（排除已确认摘要）。

    - 有正式项目编码 → 01_按项目（项目按金额大→小，摘要按占比大→小）
    - 无正式编码 → 01b 合并待核对
    - 00 全量汇总 + 处理优先级总览（综合项目/摘要）
    - 02 按月份仍仅 THCC（逻辑不变）
    """
    paths.ensure_output_parents()
    confirmed = load_confirmed_summaries(paths.learn_pool_confirmed)
    confirmed_keys = set(confirmed.keys())

    print(f"已确认摘要: {len(confirmed_keys)}")
    print(f"扫描 std: {paths.std_dir}")
    lines = collect_raw_lines_from_std(paths.std_dir)
    print(f"明细行: {len(lines)}")

    # 过滤已确认
    lines = [r for r in lines if r["摘要"] not in confirmed_keys]
    print(f"未确认明细行: {len(lines)}")

    by_project_coded: Dict[Tuple[str, str], List[Dict]] = defaultdict(list)
    by_project_nocode: Dict[Tuple[str, str], List[Dict]] = defaultdict(list)
    by_month: Dict[str, List[Dict]] = defaultdict(list)
    thcc_lines = 0
    coded_lines = 0
    nocode_lines = 0
    for r in lines:
        if r.get("has_code"):
            by_project_coded[(r["项目编码"], r["项目名称"])].append(r)
            coded_lines += 1
        else:
            by_project_nocode[(r["项目编码"], r["项目名称"])].append(r)
            nocode_lines += 1
        if r.get("is_thcc"):
            by_month[r["月份"]].append(r)
            thcc_lines += 1
    print(
        f"有编码明细 {coded_lines} 行 / 无编码 {nocode_lines} 行 · "
        f"按月份 THCC {thcc_lines} 行"
    )

    # 清空旧的 00/01/01b/02 文件（避免残留摘要；总览最后重写）
    for d in (
        paths.learn_pool_all,
        paths.learn_pool_by_project,
        paths.learn_pool_no_code,
        paths.learn_pool_by_month,
    ):
        d.mkdir(parents=True, exist_ok=True)
        for f in d.glob("*.xlsx"):
            if f.name.startswith("~$"):
                continue
            try:
                f.unlink()
            except OSError:
                pass

    files_snap: Dict[str, Dict[str, Dict[str, str]]] = {}
    project_files = 0
    month_files = 0
    unique_summ: set = set()
    project_rank: List[Dict[str, Any]] = []

    def _snap_rows(key: str, rows: List[Dict[str, Any]]) -> None:
        files_snap[key] = {}
        for r in rows:
            unique_summ.add(r["摘要"])
            files_snap[key][r["摘要"]] = {
                "一级科目": r["一级科目"],
                "二级科目": r["二级科目"],
                "三级科目": r["三级科目"],
            }

    # 00_汇总：固定 4 个 Excel = 有/无编码 × 来源分组(流水/银行账)
    # 来源分组来自 B_std 一级目录；若实际还有其他组则额外追加文件
    present_groups = {safe_text(r.get("来源分组")) or "未分组" for r in lines}
    forced_groups = ["流水", "银行账"]
    for g in sorted(present_groups):
        if g not in forced_groups:
            forced_groups.append(g)

    all_rows: List[Dict[str, Any]] = []  # 合并用于总览 Top
    summary_files = 0
    summary_file_list: List[str] = []

    def _summary_fname(has_code: bool, group: str) -> str:
        code_tag = "有编码" if has_code else "无编码"
        return f"{code_tag}_{safe_filename(group)}_学习池.xlsx"

    for group in forced_groups:
        for has_code, code_label in ((True, "有编码"), (False, "无编码")):
            subset = [
                r
                for r in lines
                if bool(r.get("has_code")) == has_code
                and (safe_text(r.get("来源分组")) or "未分组") == group
            ]
            fname = _summary_fname(has_code, group)
            out = paths.learn_pool_all / fname
            if subset:
                rows = _dedupe_rows(subset, sort_by_amount=True)
                all_rows.extend(rows)
                _format_pool_workbook(
                    out,
                    rows,
                    paths.classify_data,
                    title=(
                        f"00汇总 · {code_label} · 来源「{group}」 · "
                        f"金额优先 · 占比单元格显示为 0.21%"
                    ),
                )
                _snap_rows(f"00_汇总/{fname}", rows)
                summary_files += 1
                summary_file_list.append(fname)
                print(
                    f"  汇总 {fname}: {len(rows)} 摘要 · {code_label}/{group}"
                )
            elif group in ("流水", "银行账"):
                # 仅流水/银行账写空壳，保证 4 个固定文件
                empty_note = [
                    {
                        "确认状态": STATUS_PENDING,
                        "处理优先级": 1,
                        "摘要": f"（本分类暂无未确认摘要：{code_label}/{group}）",
                        "一级科目": "",
                        "二级科目": "",
                        "三级科目": "",
                        "供应商名称": "",
                        "出现次数": 0,
                        "应付金额合计": 0,
                        "金额占比%": 0,
                        "累计占比%": 0,
                        "来源分组": group,
                        "支出归类": "",
                    }
                ]
                _format_pool_workbook(
                    out,
                    empty_note,
                    paths.classify_data,
                    title=f"00汇总 · {code_label} · {group} · 当前无数据",
                )
                summary_file_list.append(fname)
                print(f"  汇总 {fname}: 0 摘要（空壳占位）")

    # 总览用：跨 4 表同摘要取应付金额最大的一条
    if all_rows:
        by_s: Dict[str, Dict[str, Any]] = {}
        for r in all_rows:
            s = safe_text(r.get("摘要"))
            if not s or s.startswith("（本分类暂无"):
                continue
            prev = by_s.get(s)
            if prev is None or _amt_value(r.get("应付金额合计")) > _amt_value(
                prev.get("应付金额合计")
            ):
                by_s[s] = dict(r)
        all_rows = _resort_pool_rows_by_amount(list(by_s.values()))
    else:
        all_rows = []

    # 01_按项目：仅正式编码，项目按金额大→小
    coded_items = sorted(
        by_project_coded.items(),
        key=lambda x: -_project_payable_total(x[1]),
    )
    grand_coded = sum(_project_payable_total(v) for _, v in coded_items) or 0.0
    for rank, ((code, pname), items) in enumerate(coded_items, start=1):
        rows = _dedupe_rows(items, sort_by_amount=True)
        if not rows:
            continue
        proj_amt = _project_payable_total(items)
        share = (100.0 * proj_amt / grand_coded) if grand_coded > 1e-12 else 0.0
        # 文件名带优先级序号，便于从资源管理器按名称排序处理
        fname = (
            f"{rank:02d}_{safe_filename(code)}_{safe_filename(pname)}.xlsx"
        )
        out = paths.learn_pool_by_project / fname
        _format_pool_workbook(
            out,
            rows,
            paths.classify_data,
            title=(
                f"项目学习池 · 优先级#{rank} · {code} {pname} · "
                f"项目应付合计 {proj_amt:,.2f}（占有编码 {share:.2f}%）"
            ),
        )
        project_files += 1
        _snap_rows(f"01_按项目/{fname}", rows)
        project_rank.append(
            {
                "项目优先级": rank,
                "项目编码": code,
                "项目名称": pname,
                "未确认摘要数": len(rows),
                "应付金额合计": round(proj_amt, 2),
                "占有编码金额%": round(share, 2),
                "学习池文件": fname,
            }
        )
        print(
            f"  项目 #{rank:02d} {fname}: {len(rows)} 摘要 · 应付 {proj_amt:,.0f}"
        )

    # 01b 无编码：合并一张表 + 清单
    no_code_lines: List[Dict] = []
    no_code_projects: List[Dict[str, Any]] = []
    for (code, pname), items in sorted(
        by_project_nocode.items(),
        key=lambda x: -_project_payable_total(x[1]),
    ):
        amt = _project_payable_total(items)
        no_code_lines.extend(items)
        no_code_projects.append(
            {
                "项目编码": code,
                "项目名称": pname,
                "明细行数": len(items),
                "应付金额合计": round(amt, 2),
                "说明": "无正式编码，合并待核对",
            }
        )
    no_code_rows: List[Dict[str, Any]] = []
    if no_code_lines:
        no_code_rows = _dedupe_rows(no_code_lines, sort_by_amount=True)
        _format_pool_workbook(
            paths.learn_pool_no_code_xlsx,
            no_code_rows,
            paths.classify_data,
            title=(
                f"待核对 · 无正式项目编码汇总 · {len(no_code_projects)} 个来源 · "
                f"请后续补编码后再并入 01_按项目"
            ),
        )
        _snap_rows(
            f"01b_待核对_无项目编码/{paths.learn_pool_no_code_xlsx.name}",
            no_code_rows,
        )
        print(
            f"  无编码汇总: {len(no_code_rows)} 摘要 · 来源项目 {len(no_code_projects)}"
        )

    # 02 按月份：仅 THCC（逻辑不变）
    for month, items in sorted(by_month.items()):
        rows = _dedupe_rows(items, sort_by_amount=True)
        if not rows:
            continue
        if month == "未知月份":
            fname = "未知月份_财务数据.xlsx"
        else:
            fname = f"{month}_财务数据.xlsx"
        out = paths.learn_pool_by_month / fname
        _format_pool_workbook(
            out,
            rows,
            paths.classify_data,
            title=f"月份学习池 · {month} · THCC · 金额优先",
        )
        month_files += 1
        _snap_rows(f"02_按月份/{fname}", rows)
        print(f"  月份 {fname}: {len(rows)} 摘要")

    # 00 处理优先级总览
    _write_priority_overview(
        paths.learn_pool_priority_xlsx,
        project_rank=project_rank,
        global_summaries=all_rows,
        no_code_projects=no_code_projects,
        latest_note=(
            f"有编码项目 {len(project_rank)} · 无编码来源 {len(no_code_projects)} · "
            f"未确认摘要 {len(unique_summ)}"
        ),
    )
    print(f"  总览 {paths.learn_pool_priority_xlsx.name}")

    save_snapshot(paths.learn_pool_snapshot, files_snap)

    if not paths.learn_pool_confirmed.is_file():
        _write_confirmed_excel(paths.learn_pool_confirmed, [])

    stats = {
        "unconfirmed_detail_lines": len(lines),
        "all_summary_rows": len(all_rows),
        "summary_00_files": summary_files,
        "summary_00_list": summary_file_list,
        "project_files": project_files,
        "no_code_sources": len(no_code_projects),
        "no_code_summaries": len(no_code_rows),
        "month_files": month_files,
        "unique_summaries": len(unique_summ),
        "confirmed_excluded": len(confirmed_keys),
        "pool_root": str(paths.learn_pool_root),
        "priority_xlsx": str(paths.learn_pool_priority_xlsx),
        "no_code_xlsx": str(paths.learn_pool_no_code_xlsx),
    }
    print(
        f"✅ 导出学习池: 00汇总文件 {len(summary_file_list)} · 去重摘要 {len(all_rows)} · "
        f"有编码项目 {project_files} · 无编码来源 {len(no_code_projects)} · "
        f"月份 {month_files}"
    )
    print(f"   目录: {paths.learn_pool_root}")
    print(f"   先看优先级: {paths.learn_pool_priority_xlsx}")
    print(f"   00_汇总四表: {', '.join(summary_file_list[:4])}")
    print(f"   无编码待核: {paths.learn_pool_no_code_xlsx}")
    return stats


def _write_confirmed_excel(path: Path, rows: List[Dict[str, Any]]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        "确认状态",
        "摘要",
        "一级科目",
        "二级科目",
        "三级科目",
        "供应商名称",
        "出现次数",
        "确认时间",
        "来源",
    ]
    # 归类排序：一级 → 二级 → 三级 → 摘要
    rows = _sort_by_category(list(rows)) if rows else []
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="已确认", index=False)
        ws = writer.book["已确认"]
        ws.freeze_panes = "A2"
        if len(df):
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df) + 1}"
        for i, w in enumerate([10, 36, 16, 18, 16, 22, 10, 18, 24], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def _read_pool_file(path: Path) -> List[Dict[str, Any]]:
    """优先读「学习池」sheet；兼容旧文件把说明放第一页的情况。"""
    path = Path(path)
    df = None
    try:
        xls = pd.ExcelFile(path)
        sheet = POOL_SHEET if POOL_SHEET in xls.sheet_names else None
        if sheet is None:
            # 找含「摘要」列的 sheet
            for sn in xls.sheet_names:
                if sn in {"说明", "_分类选项"}:
                    continue
                tmp = pd.read_excel(xls, sheet_name=sn, nrows=1)
                cols = {str(c).strip() for c in tmp.columns}
                if "摘要" in cols:
                    sheet = sn
                    break
        if sheet is None:
            sheet = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)
    except Exception as e:
        print(f"  ⚠️ 无法读取 {path.name}: {e}")
        return []

    # 兼容列名空格
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        summ = safe_text(row.get("摘要"))
        if not summ:
            continue
        status = safe_text(row.get("确认状态")) or STATUS_PENDING
        rows.append(
            {
                "确认状态": status,
                "处理优先级": row.get("处理优先级") or "",
                "摘要": summ,
                "一级科目": _norm_cat(row.get("一级科目")),
                "二级科目": _norm_cat(row.get("二级科目")),
                "三级科目": _norm_cat(row.get("三级科目")),
                "供应商名称": safe_text(row.get("供应商名称")),
                "出现次数": row.get("出现次数") or "",
                "应付金额合计": row.get("应付金额合计")
                if "应付金额合计" in df.columns
                else "",
                "金额占比%": row.get("金额占比%") if "金额占比%" in df.columns else "",
                "累计占比%": row.get("累计占比%") if "累计占比%" in df.columns else "",
                "来源分组": safe_text(row.get("来源分组"))
                if "来源分组" in df.columns
                else "",
                "支出归类": safe_text(row.get("支出归类"))
                if "支出归类" in df.columns
                else "",
                "_file": path.name,
                "_path": str(path),
                "_mtime": path.stat().st_mtime,
            }
        )
    return rows


def _scan_pool_dirs(paths: Paths) -> List[Dict[str, Any]]:
    """扫描 00_汇总 / 01_按项目 / 01b_无编码 / 02_按月份（不含优先级总览）。"""
    rows: List[Dict[str, Any]] = []
    for d in (
        paths.learn_pool_all,
        paths.learn_pool_by_project,
        paths.learn_pool_no_code,
        paths.learn_pool_by_month,
    ):
        if not d.is_dir():
            continue
        for f in sorted(d.glob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            # 优先级总览不是学习池确认表
            if "处理优先级" in f.name or f.name.startswith("00_处理"):
                continue
            rows.extend(_read_pool_file(f))
    return rows


def _backup_taxonomy_file(tax_path: Path) -> Path:
    """备份 taxonomy.json，返回备份路径。"""
    import shutil
    from datetime import datetime as _dt

    tax_path = Path(tax_path)
    stamp = _dt.now().strftime("%Y%m%d_%H%M%S")
    bak = tax_path.with_name(f"taxonomy_backup_{stamp}.json")
    shutil.copy2(tax_path, bak)
    return bak


def _patch_subject_map(data_dir: Path, name_to_l2: Dict[str, str]) -> int:
    """把新二级名写入 subject_map.json（可选加速）。"""
    path = Path(data_dir) / "subject_map.json"
    if not name_to_l2:
        return 0
    data: Dict[str, str] = {}
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    n = 0
    for name, lid in name_to_l2.items():
        if name and lid and data.get(name) != lid:
            data[name] = lid
            n += 1
    if n:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return n


def apply_learn_pool(
    paths: Paths,
    *,
    reclassify: bool = True,
    auto_taxonomy: bool = True,
    learn_unconfirmed_edits: bool = True,
) -> Dict[str, Any]:
    """
    检测手改/确认 → 按业务摘要写记忆 → 合并已确认表 → 剥离 → 刷新未确认建议。

    不改 A2 / B_std。把科目填进月结文件是财务付款流程的事。

    auto_taxonomy:
        学习池中填写了体系里不存在的一/二/三级时，自动登记到 taxonomy（先备份）。
    learn_unconfirmed_edits:
        待确认但相对导出快照已改科目的行也写入记忆（不仅「确认」状态）。
    """
    ensure_code_root()
    from learn_classify.engine import ClassifyEngine
    from learn_classify.learn import learn_one
    from learn_classify.store import CorrectionStore, MemoryStore
    from learn_classify.taxonomy import Taxonomy
    from summary_recognize import recognize_summary

    paths.ensure_output_parents()
    snapshot = load_snapshot(paths.learn_pool_snapshot)
    pool_rows = _scan_pool_dirs(paths)
    snap_rows = sum(len(v) for v in snapshot.values())
    print(f"扫描学习池行: {len(pool_rows)} · 快照文件 {len(snapshot)} · 快照行 {snap_rows}")
    print(
        f"选项: auto_taxonomy={auto_taxonomy} · "
        f"learn_unconfirmed_edits={learn_unconfirmed_edits}"
    )

    def _snap_key_for_row(it: Dict[str, Any]) -> str:
        """根据文件所在目录还原导出时的 snapshot key。"""
        p = Path(it.get("_path") or "")
        name = it.get("_file") or p.name
        parent = p.parent.name if p.parent else ""
        if parent == "00_汇总" or "00_汇总" in str(p):
            return f"00_汇总/{name}"
        if parent == "01_按项目" or "01_按项目" in str(p):
            return f"01_按项目/{name}"
        if "01b" in parent or "无项目编码" in parent or "01b" in str(p):
            return f"01b_待核对_无项目编码/{name}"
        if parent == "02_按月份" or "02_按月份" in str(p):
            return f"02_按月份/{name}"
        for prefix in (
            "00_汇总",
            "01_按项目",
            "01b_待核对_无项目编码",
            "02_按月份",
        ):
            k = f"{prefix}/{name}"
            if k in snapshot:
                return k
        return name

    # 按摘要聚合候选锁定
    by_summ: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in pool_rows:
        by_summ[r["摘要"]].append(r)

    locked: Dict[str, Dict[str, Any]] = {}
    conflicts: List[Dict[str, Any]] = []

    for summ, items in by_summ.items():
        candidates = []
        for it in items:
            sk = _snap_key_for_row(it)
            file_snap = snapshot.get(sk) or snapshot.get("__global__") or {}
            snap = file_snap.get(summ) or {}
            snap_t = _triple(snap.get("一级科目"), snap.get("二级科目"), snap.get("三级科目"))
            cur_t = _triple(it["一级科目"], it["二级科目"], it["三级科目"])
            confirmed = _is_confirmed(it["确认状态"])
            has_snap = bool(snap)
            changed = has_snap and cur_t != snap_t and (cur_t[0] or cur_t[1] or cur_t[2])
            # 无该文件快照时：仅「确认」才锁定（避免整表误学）
            if not has_snap and not confirmed:
                continue
            # 确认 或（可选）待确认但相对导出快照已手改科目
            if confirmed or (learn_unconfirmed_edits and changed):
                if not cur_t[1] and not cur_t[2]:
                    continue
                candidates.append(
                    {
                        **it,
                        "_confirmed": confirmed,
                        "_changed": changed,
                        "_triple": cur_t,
                    }
                )

        if not candidates:
            continue

        # 优先级：确认 > 手改；同级取 mtime 最新
        candidates.sort(
            key=lambda x: (1 if x["_confirmed"] else 0, x["_mtime"]),
            reverse=True,
        )
        best = candidates[0]
        conf_set = {c["_triple"] for c in candidates if c["_confirmed"]}
        if len(conf_set) > 1:
            conflicts.append(
                {
                    "摘要": summ,
                    "冲突说明": "多处确认且科目不一致",
                    "采用": f"{best['_triple'][0]}/{best['_triple'][1]}/{best['_triple'][2]}",
                    "来源": best.get("_file"),
                }
            )
        locked[summ] = best

    print(f"锁定摘要: {len(locked)} · 冲突: {len(conflicts)}")

    # 学习（可选自动登记新科目）
    tax_path = paths.classify_data / "taxonomy.json"
    tax = Taxonomy.load(tax_path)
    memory = MemoryStore(paths.classify_data / "memory.jsonl")
    corrections = CorrectionStore(paths.classify_data / "corrections.jsonl")
    learned = 0
    learn_errors = 0
    learned_texts: set[str] = set()
    tax_created_all: List[str] = []
    tax_backup: Optional[str] = None
    subject_map_hits: Dict[str, str] = {}

    for summ, it in locked.items():
        l1, l2, l3 = it["_triple"]
        try:
            # 先探测是否需要登记（避免无新类时也备份）
            need_register = False
            if auto_taxonomy:
                _a, _b, _c = tax.resolve_triple(l1 or None, l2 or None, l3 or None)
                if not _b and not _c and (l2 or l3):
                    need_register = True
            if need_register and tax_backup is None:
                bak = _backup_taxonomy_file(tax_path)
                tax_backup = str(bak)
                print(f"  📦 taxonomy 已备份: {bak.name}")

            learn_text = (
                recognize_summary(summ, it.get("项目名称") or "").business_summary
                or summ
            )
            result = learn_one(
                taxonomy=tax,
                memory=memory,
                corrections=corrections,
                vendor_stats=None,
                subject="",
                summary=learn_text,
                vendor=it.get("供应商名称") or "",
                new_l1=l1,
                new_l2=l2 or l3,
                new_l3=l3 if l2 else "",
                source="learn_pool",
                notes=(
                    f"file={it.get('_file')};confirmed={it.get('_confirmed')};"
                    f"changed={it.get('_changed')}"
                ),
                also_summary_only=True,
                auto_register=auto_taxonomy,
            )
            # 兼容旧 learn_one 三元组 / 新四元组
            if len(result) >= 4:
                l1n, l2n, l3n, clog = result[0], result[1], result[2], result[3]
            else:
                l1n, l2n, l3n, clog = result[0], result[1], result[2], []
            learned_texts.add(learn_text)
            if clog:
                tax_created_all.extend(clog)
                print(f"  ＋ 新科目 {summ[:24]}… → {clog}")
            # subject_map
            leaf = tax.resolve_l3_id(l3n) or tax.resolve_l2_id(l2n)
            if leaf and l2n:
                subject_map_hits[l2n] = leaf
            learned += 1
        except Exception as e:
            learn_errors += 1
            print(f"  ⚠️ 学习失败 {summ[:30]}: {e}")

    # 有自动登记则落盘 taxonomy + 科目体系 + subject_map
    if tax_created_all:
        # 版本尾巴标记
        if "auto" not in (tax.version or ""):
            tax.version = f"{tax.version}+auto"
        tax.description = (
            (tax.description or "")
            + f" | auto_register { _now() }: {len(tax_created_all)} nodes"
        ).strip(" |")
        tax.save(tax_path)
        sm_n = _patch_subject_map(paths.classify_data, subject_map_hits)
        print(f"  💾 taxonomy 已保存（新增 {len(tax_created_all)} 项标记）")
        if sm_n:
            print(f"  💾 subject_map 更新 {sm_n} 条")
        try:
            from learn_classify.taxonomy_io import export_taxonomy_excel

            export_taxonomy_excel(paths.classify_data, paths.taxonomy_xlsx)
            print(f"  💾 科目体系.xlsx 已刷新")
        except Exception as e:
            print(f"  ⚠️ 导出科目体系失败: {e}")

    print(f"写入记忆: {learned} · 失败: {learn_errors} · 新科目登记: {len(tax_created_all)}")

    # 合并已确认表
    confirmed = load_confirmed_summaries(paths.learn_pool_confirmed)
    for summ, it in locked.items():
        l1, l2, l3 = it["_triple"]
        learn_text = (
            recognize_summary(summ, it.get("项目名称") or "").business_summary or summ
        )
        confirmed[learn_text] = {
            "确认状态": STATUS_CONFIRMED,
            "摘要": learn_text,
            "一级科目": l1,
            "二级科目": l2,
            "三级科目": l3,
            "供应商名称": it.get("供应商名称") or "",
            "出现次数": it.get("出现次数") or "",
            "确认时间": _now(),
            "来源": it.get("_file") or "",
        }
    conf_rows = list(confirmed.values())
    conf_rows.sort(key=lambda r: r.get("确认时间") or "", reverse=True)
    _write_confirmed_excel(paths.learn_pool_confirmed, conf_rows)
    print(f"已确认表: {paths.learn_pool_confirmed} ({len(conf_rows)} 行)")

    if conflicts:
        try:
            pd.DataFrame(conflicts).to_excel(paths.learn_pool_conflicts, index=False)
            print(f"冲突记录: {paths.learn_pool_conflicts}")
        except Exception as e:
            print(f"  ⚠️ 写冲突表失败: {e}")

    # 从 00/01/01b/02 剥离锁定摘要
    locked_keys = set(locked.keys()) | learned_texts
    stripped_files = 0
    for d in (
        paths.learn_pool_all,
        paths.learn_pool_by_project,
        paths.learn_pool_no_code,
        paths.learn_pool_by_month,
    ):
        if not d.is_dir():
            continue
        for f in sorted(d.glob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            if "处理优先级" in f.name or f.name.startswith("00_处理"):
                continue
            rows = _read_pool_file(f)
            if not rows:
                continue
            kept = []
            for r in rows:
                cell = r["摘要"]
                biz = recognize_summary(cell).business_summary or cell
                if cell in locked_keys or biz in locked_keys:
                    continue
                kept.append(r)
            removed = len(rows) - len(kept)
            if removed == 0:
                continue
            # 写回：保留金额列并按金额重排优先级（不得丢列）
            out_rows = []
            for r in kept:
                out_rows.append(
                    {
                        "确认状态": r.get("确认状态") or STATUS_PENDING,
                        "处理优先级": r.get("处理优先级") or "",
                        "摘要": r["摘要"],
                        "一级科目": r["一级科目"],
                        "二级科目": r["二级科目"],
                        "三级科目": r["三级科目"],
                        "供应商名称": r.get("供应商名称") or "",
                        "出现次数": r.get("出现次数") or "",
                        "应付金额合计": r.get("应付金额合计")
                        if r.get("应付金额合计") not in (None, "")
                        else 0,
                        "金额占比%": r.get("金额占比%") or 0,
                        "累计占比%": r.get("累计占比%") or 0,
                        "来源分组": r.get("来源分组") or "",
                        "支出归类": r.get("支出归类") or "",
                    }
                )
            out_rows = _resort_pool_rows_by_amount(out_rows)
            if not out_rows:
                try:
                    f.unlink()
                    print(f"  🗑️ 已空删除 {f.name}")
                except OSError:
                    _format_pool_workbook(f, [], paths.classify_data)
            else:
                _format_pool_workbook(
                    f,
                    out_rows,
                    paths.classify_data,
                    title="学习池 · apply 后保留金额优先级排序",
                )
            stripped_files += 1
            print(f"  剥离 {f.name}: -{removed} 行，剩 {len(out_rows)}")

    # 用新记忆 + 新科目体系刷新未确认池建议（含 01b 无编码）
    # 即使本次 0 条锁定，若刚登记了科目也重载引擎
    pool_refresh = {}
    eng = None
    if learned or tax_created_all:
        print("刷新未确认池建议（memory / core / similar / 新科目）…")
        eng = ClassifyEngine.from_data_dir(paths.classify_data)
        pool_refresh = refresh_pool_suggestions(paths, eng)
        print(f"池建议刷新: {pool_refresh}")

    # 更新按文件快照（剩余行，含刷新后科目）
    remaining = _scan_pool_dirs(paths)
    new_snap: Dict[str, Dict[str, Dict[str, str]]] = defaultdict(dict)
    for r in remaining:
        sk = _snap_key_for_row(r)
        new_snap[sk][r["摘要"]] = {
            "一级科目": r["一级科目"],
            "二级科目": r["二级科目"],
            "三级科目": r["三级科目"],
        }
    save_snapshot(paths.learn_pool_snapshot, dict(new_snap))

    # apply 后重建优先级总览，避免 00_处理优先级_总览 丢失
    try:
        ov = rebuild_priority_overview_from_disk(paths)
        print(f"  ↻ 已重建处理优先级总览: {ov.get('priority_xlsx')}")
    except Exception as e:
        print(f"  ⚠️ 重建优先级总览失败: {e}")

    if reclassify:
        print(
            "应用确认不再改财务文件。本次只写入科目记忆和学习池。"
            "要把科目填进月结结果，请另运行: python 财务付款一体化.py reclassify"
        )
    else:
        print("本次只写入科目记忆和学习池，不改财务分类文件。")
    reclass_stats = {"skipped": "pipeline_fill_is_separate"}

    stats = {
        "locked": len(locked),
        "learned": learned,
        "learn_errors": learn_errors,
        "conflicts": len(conflicts),
        "confirmed_total": len(conf_rows),
        "stripped_files": stripped_files,
        "pool_refresh": pool_refresh,
        "reclassify": reclass_stats,
        "auto_taxonomy": auto_taxonomy,
        "taxonomy_created": tax_created_all,
        "taxonomy_backup": tax_backup,
        "learn_unconfirmed_edits": learn_unconfirmed_edits,
    }
    print(f"✅ apply-learn-pool 完成: locked={learned}")
    if tax_created_all:
        print(f"   新科目: {', '.join(tax_created_all[:12])}"
              + ("…" if len(tax_created_all) > 12 else ""))
    return stats


# 可接受的引擎层（用于抬升未确认建议）
_REFRESH_LAYERS = {
    "memory",
    "memory_core",
    "similar_memory",
    "rule",
    "rule_hi",
    "pattern",
    "subject_map",
    "vendor",
}

def refresh_pool_suggestions(
    paths: Paths,
    eng: Any,
    *,
    min_confidence: float = 0.55,
) -> Dict[str, Any]:
    """
    用当前引擎重算 00/01/02 未确认行的 L1/L2/L3。

    - 只抬升：空科目 / 待复核类 / 或新建议置信更高且层更强
    - 不改「确认状态」
    - 写回后与快照对齐，避免下次误判为手改
    """
    ensure_code_root()
    from summary_recognize import recognize_summary

    files_updated = 0
    rows_scanned = 0
    rows_changed = 0
    by_layer: Counter = Counter()

    for d in (
        paths.learn_pool_all,
        paths.learn_pool_by_project,
        paths.learn_pool_no_code,
        paths.learn_pool_by_month,
    ):
        if not d.is_dir():
            continue
        for f in sorted(d.glob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            if "处理优先级" in f.name or f.name.startswith("00_处理"):
                continue
            rows = _read_pool_file(f)
            if not rows:
                continue
            changed = False
            out_rows: List[Dict[str, Any]] = []
            for r in rows:
                rows_scanned += 1
                status = r.get("确认状态") or STATUS_PENDING

                def _base_row(src: Dict[str, Any], st: str) -> Dict[str, Any]:
                    # 保留金额/优先级/来源列，仅刷新科目
                    return {
                        "确认状态": st,
                        "处理优先级": src.get("处理优先级") or "",
                        "摘要": src["摘要"],
                        "一级科目": src.get("一级科目") or "",
                        "二级科目": src.get("二级科目") or "",
                        "三级科目": src.get("三级科目") or "",
                        "供应商名称": src.get("供应商名称") or "",
                        "出现次数": src.get("出现次数") or "",
                        "应付金额合计": src.get("应付金额合计") or "",
                        "金额占比%": src.get("金额占比%") or "",
                        "累计占比%": src.get("累计占比%") or "",
                        "来源分组": src.get("来源分组") or "",
                        "支出归类": src.get("支出归类") or "",
                    }

                if _is_confirmed(status):
                    out_rows.append(_base_row(r, status))
                    continue

                old_t = _triple(r["一级科目"], r["二级科目"], r["三级科目"])
                try:
                    text = recognize_summary(r["摘要"]).business_summary or r["摘要"]
                    res = eng.classify(
                        "",
                        text,
                        r.get("供应商名称") or "",
                    )
                except Exception:
                    res = None

                new_row = _base_row(r, STATUS_PENDING)

                if res is not None and res.layer in _REFRESH_LAYERS:
                    conf = float(res.confidence or 0)
                    if conf >= min_confidence:
                        new_t = (
                            _norm_cat(res.l1_name),
                            _norm_cat(res.l2_name),
                            _norm_cat(res.l3_name),
                        )
                        if new_t[1] or new_t[2]:
                            should = False
                            # 原无有效二级/三级 → 直接抬升
                            if not old_t[1] and not old_t[2]:
                                should = True
                            elif new_t != old_t:
                                # 原像待复核 / 空白一级
                                if not old_t[0] or "待" in old_t[0] or "待" in old_t[1]:
                                    should = True
                                # 相似/记忆/规则 且置信足够时覆盖旧建议
                                elif res.layer in {
                                    "memory",
                                    "memory_core",
                                    "similar_memory",
                                    "rule",
                                    "rule_hi",
                                    "pattern",
                                } and conf >= 0.65:
                                    should = True
                            if should and new_t != old_t:
                                new_row["一级科目"] = new_t[0]
                                new_row["二级科目"] = new_t[1]
                                new_row["三级科目"] = new_t[2]
                                rows_changed += 1
                                by_layer[res.layer] += 1
                                changed = True

                out_rows.append(new_row)

            if changed:
                # 始终按金额重排，保留优先级列（无金额时占比为 0 仍保持列结构）
                out_rows = _resort_pool_rows_by_amount(out_rows)
                _format_pool_workbook(
                    f,
                    out_rows,
                    paths.classify_data,
                    title="学习池 · 已根据确认结果刷新预分类建议（金额优先）",
                )
                files_updated += 1
                print(f"  ↻ 建议更新 {f.name}")

    return {
        "files_updated": files_updated,
        "rows_scanned": rows_scanned,
        "rows_changed": rows_changed,
        "by_layer": dict(by_layer),
    }
