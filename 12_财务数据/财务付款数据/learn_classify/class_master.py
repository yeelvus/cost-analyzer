# -*- coding: utf-8 -*-
"""
付款分类总表（学习专用，按摘要去重）

核心列：摘要、一级科目、二级科目、供应商名称
辅助列：出现次数、金额合计、示例项目、命中层、备注 等

用法：
  export_class_master  → 导出去重总表
  apply_class_master   → 按「摘要→一二级」学习并重刷 B_std
"""

from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .engine import ClassifyEngine
from .learn import learn_one
from .review import reclassify_std_dir
from .store import CorrectionStore, MemoryStore, VendorStats
from .taxonomy import Taxonomy

MASTER_SHEET = "分类总表"
README_SHEET = "使用说明"

# 核心列（学习焦点）+ 辅助列
MASTER_COLUMNS = [
    "摘要",
    "一级科目",
    "二级科目",
    "三级科目",
    "供应商名称",
    # —— 以下为辅助 ——
    "出现次数",
    "金额合计",
    "示例项目",
    "示例科目",
    "命中层",
    "分类ID",
    "备注",
]


def _safe_str(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if pd.isna(v):
        return ""
    return str(v).strip()


def _pick_sheet(path: Path) -> str:
    try:
        xls = pd.ExcelFile(path)
        if "明细数据" in xls.sheet_names:
            return "明细数据"
        if "明细数据_万泰铢" in xls.sheet_names:
            return "明细数据_万泰铢"
        return xls.sheet_names[0]
    except Exception:
        return "Sheet1"


def _collect_raw_lines(std_dir: Path) -> List[Dict[str, Any]]:
    """收集明细行（未去重）。"""
    std_dir = Path(std_dir)
    lines: List[Dict[str, Any]] = []

    for f in sorted(std_dir.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        sheet = _pick_sheet(f)
        try:
            df = pd.read_excel(f, sheet_name=sheet)
        except Exception as e:
            print(f"  ⚠️ 跳过 {f.name}: {e}")
            continue

        cols = {str(c).strip(): c for c in df.columns}

        def col(*names):
            for n in names:
                if n in cols:
                    return cols[n]
            return None

        c_proj = col("项目名称")
        c_subj = col("科目")
        c_summ = col("摘要")
        c_vend = col("供应商名称", "供应商")
        c_amt = col("实付金额")
        c_l1 = col("一级科目")
        c_l2 = col("二级科目")
        c_l3 = col("三级科目")
        c_id = col("分类ID")
        c_layer = col("命中层")

        for _, row in df.iterrows():
            summ = _safe_str(row.get(c_summ)) if c_summ else ""
            if not summ:
                continue
            amt = 0.0
            if c_amt is not None:
                try:
                    amt = float(row.get(c_amt) or 0)
                except Exception:
                    amt = 0.0
            lines.append(
                {
                    "摘要": summ,
                    "供应商名称": _safe_str(row.get(c_vend)) if c_vend else "",
                    "一级科目": _safe_str(row.get(c_l1)) if c_l1 else "",
                    "二级科目": _safe_str(row.get(c_l2)) if c_l2 else "",
                    "三级科目": _safe_str(row.get(c_l3)) if c_l3 else "",
                    "分类ID": _safe_str(row.get(c_id)) if c_id else "",
                    "命中层": _safe_str(row.get(c_layer)) if c_layer else "",
                    "科目": _safe_str(row.get(c_subj)) if c_subj else "",
                    "项目名称": _safe_str(row.get(c_proj)) if c_proj else "",
                    "实付金额": amt,
                }
            )
    return lines


def build_deduped_master(
    std_dir: Path,
    engine: Optional[ClassifyEngine] = None,
) -> pd.DataFrame:
    """
    按「摘要」完全相同去重。
    一级/二级：取该摘要下出现最多的组合；若文件无分类则用引擎补全。
    供应商：取出现最多的非空供应商。
    """
    lines = _collect_raw_lines(std_dir)
    if not lines:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for ln in lines:
        groups[ln["摘要"]].append(ln)

    rows: List[Dict[str, Any]] = []
    for summ, items in sorted(groups.items(), key=lambda x: (-len(x[1]), x[0])):
        n = len(items)
        amt_sum = sum(float(x.get("实付金额") or 0) for x in items)

        # 供应商众数
        vend_counter = Counter(
            x["供应商名称"] for x in items if x.get("供应商名称")
        )
        vendor = vend_counter.most_common(1)[0][0] if vend_counter else ""

        # 一/二/三级众数
        pair_counter = Counter()
        for x in items:
            l1, l2, l3 = (
                x.get("一级科目") or "",
                x.get("二级科目") or "",
                x.get("三级科目") or "",
            )
            if l1 or l2 or l3:
                pair_counter[(l1, l2, l3)] += 1
        if pair_counter:
            (l1, l2, l3), _ = pair_counter.most_common(1)[0]
        else:
            l1, l2, l3 = "", "", ""

        # 示例
        proj_counter = Counter(x["项目名称"] for x in items if x.get("项目名称"))
        subj_counter = Counter(x["科目"] for x in items if x.get("科目"))
        layer_counter = Counter(x["命中层"] for x in items if x.get("命中层"))
        id_counter = Counter(x["分类ID"] for x in items if x.get("分类ID"))

        # 引擎补全
        cid = id_counter.most_common(1)[0][0] if id_counter else ""
        layer = layer_counter.most_common(1)[0][0] if layer_counter else ""
        if engine is not None and (not l1 or not l2):
            r = engine.classify(
                subj_counter.most_common(1)[0][0] if subj_counter else "",
                summ,
                vendor,
            )
            l1 = l1 or r.l1_name
            l2 = l2 or r.l2_name
            l3 = l3 or r.l3_name
            cid = cid or (r.l3_id or r.l2_id)
            layer = layer or r.layer
        elif engine is not None and l2 and not l3:
            # 有二级无三级时尝试引擎补三级
            r = engine.classify(
                subj_counter.most_common(1)[0][0] if subj_counter else "",
                summ,
                vendor,
            )
            if r.l2_name == l2 and r.l3_name:
                l3 = r.l3_name
                cid = r.l3_id or cid

        rows.append(
            {
                "摘要": summ,
                "一级科目": l1,
                "二级科目": l2,
                "三级科目": l3,
                "供应商名称": vendor,
                "出现次数": n,
                "金额合计": round(amt_sum, 2),
                "示例项目": proj_counter.most_common(1)[0][0] if proj_counter else "",
                "示例科目": subj_counter.most_common(1)[0][0] if subj_counter else "",
                "命中层": layer,
                "分类ID": cid,
                "备注": "",
            }
        )

    # 出现次数多的排前面，方便先改高频摘要
    df = pd.DataFrame(rows, columns=MASTER_COLUMNS)
    return df


def _apply_master_format(path: Path, data_dir: Path) -> None:
    wb = load_workbook(path)

    if README_SHEET in wb.sheetnames:
        del wb[README_SHEET]
    ws_r = wb.create_sheet(README_SHEET, 0)
    tips = [
        "【付款分类总表 · 按摘要去重 · 学习专用】",
        "",
        "【三级分类】一级 → 二级 → 三级（均可下拉修改）",
        "  例：6_管理费 → 办公费 → 网络通讯费 / 邮寄快递费 / 办公物品采购 / 其他办公",
        "  无三级的二级：三级列可留空。",
        "",
        "【粒度】业务大类，不到型号规格（C30/HRB400 等）。",
        "",
        "核心列（请维护）：",
        "  · 摘要、一级科目、二级科目、三级科目、供应商名称",
        "",
        "辅助列：出现次数 / 金额合计 / 示例项目 / 示例科目 / 命中层 / 分类ID / 备注",
        "",
        "操作：",
        "  1. 下拉修改一/二/三级后保存",
        "  2. python 财务付款一体化.py apply-class",
        "  3. python 财务付款一体化.py export-class  （可选，刷新总表）",
        "",
        "类目维护：export-taxonomy → 编辑科目体系 → import-taxonomy",
    ]
    for i, t in enumerate(tips, 1):
        ws_r.cell(i, 1, t)
    ws_r.column_dimensions["A"].width = 92

    if MASTER_SHEET not in wb.sheetnames:
        wb.worksheets[-1].title = MASTER_SHEET
    ws = wb[MASTER_SHEET]
    try:
        wb.move_sheet(MASTER_SHEET, offset=1 - wb.sheetnames.index(MASTER_SHEET))
    except Exception:
        pass

    ws.freeze_panes = "A2"
    if ws.max_row and ws.max_column:
        ws.auto_filter.ref = ws.dimensions

    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True)
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    headers = {
        str(ws.cell(1, c).value).strip(): c
        for c in range(1, ws.max_column + 1)
        if ws.cell(1, c).value is not None
    }
    widths = {
        "摘要": 48,
        "一级科目": 14,
        "二级科目": 14,
        "三级科目": 14,
        "供应商名称": 22,
        "出现次数": 10,
        "金额合计": 14,
        "示例项目": 16,
        "示例科目": 12,
        "命中层": 12,
        "分类ID": 14,
        "备注": 18,
    }
    for name, w in widths.items():
        if name in headers:
            ws.column_dimensions[get_column_letter(headers[name])].width = w

    # 核心列高亮
    core_fill = PatternFill("solid", fgColor="FFF2CC")
    for name in ("摘要", "一级科目", "二级科目", "三级科目", "供应商名称", "备注"):
        if name not in headers:
            continue
        col = headers[name]
        for r in range(2, ws.max_row + 1):
            ws.cell(r, col).fill = core_fill

    # 下拉
    try:
        root = Path(__file__).resolve().parent.parent
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))
        from payment_hub.category_dropdowns import apply_category_dropdowns

        apply_category_dropdowns(wb, data_dir=data_dir, sheet_names=[MASTER_SHEET])
    except Exception as e:
        print(f"  ⚠️ 联动下拉失败，扁平下拉: {e}")
        try:
            tax = Taxonomy.load(Path(data_dir) / "taxonomy.json")
            l1s = [x.name for x in sorted(tax.level1.values(), key=lambda z: z.order) if x.active]
            l2s = [x.name for x in sorted(tax.level2.values(), key=lambda z: z.order) if x.active]
            opt = wb.create_sheet("_opts")
            for i, n in enumerate(l1s, 1):
                opt.cell(i, 1, n)
            for i, n in enumerate(l2s, 1):
                opt.cell(i, 2, n)
            opt.sheet_state = "hidden"
            last = max(ws.max_row, 2) + 200
            c1, c2 = headers.get("一级科目"), headers.get("二级科目")
            if c1 and l1s:
                dv1 = DataValidation(
                    type="list",
                    formula1=f"=_opts!$A$1:$A${len(l1s)}",
                    allow_blank=True,
                    showDropDown=False,
                    showErrorMessage=False,
                )
                dv1.add(f"{get_column_letter(c1)}2:{get_column_letter(c1)}{last}")
                ws.add_data_validation(dv1)
            if c2 and l2s:
                dv2 = DataValidation(
                    type="list",
                    formula1=f"=_opts!$B$1:$B${len(l2s)}",
                    allow_blank=True,
                    showDropDown=False,
                    showErrorMessage=False,
                )
                dv2.add(f"{get_column_letter(c2)}2:{get_column_letter(c2)}{last}")
                ws.add_data_validation(dv2)
        except Exception as e2:
            print(f"  ⚠️ 下拉写入失败: {e2}")

    wb.active = wb[MASTER_SHEET]
    wb.save(path)


def export_class_master(
    std_dir: Path,
    out_path: Path,
    data_dir: Path,
    *,
    engine: Optional[ClassifyEngine] = None,
) -> Tuple[Path, int]:
    data_dir = Path(data_dir)
    if engine is None:
        engine = ClassifyEngine.from_data_dir(data_dir)

    df = build_deduped_master(std_dir, engine=engine)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=MASTER_SHEET, index=False)

    _apply_master_format(out_path, data_dir)
    return out_path, len(df)


def apply_class_master(
    master_path: Path,
    std_dir: Path,
    data_dir: Path,
    *,
    reclassify: bool = True,
    reexport: bool = True,
    out_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    从分类总表学习：摘要 → 一级/二级（去重后每摘要一行）。
    """
    master_path = Path(master_path)
    data_dir = Path(data_dir)
    if not master_path.is_file():
        raise FileNotFoundError(f"分类总表不存在: {master_path}")

    try:
        df = pd.read_excel(master_path, sheet_name=MASTER_SHEET)
    except Exception:
        df = pd.read_excel(master_path, sheet_name=0)

    cols = {str(c).strip(): c for c in df.columns}
    if "摘要" not in cols or "二级科目" not in cols:
        raise ValueError(
            f"分类总表需含「摘要」「二级科目」列，当前={list(df.columns)}"
        )

    tax = Taxonomy.load(data_dir / "taxonomy.json")
    memory = MemoryStore(data_dir / "memory.jsonl")
    corrections = CorrectionStore(data_dir / "corrections.jsonl")
    vendor_stats = VendorStats(data_dir / "vendor_stats.json")

    learned = 0
    skipped = 0
    errors = 0
    seen = set()

    for _, row in df.iterrows():
        summ = _safe_str(row.get(cols["摘要"]))
        l2 = _safe_str(row.get(cols["二级科目"]))
        l1 = _safe_str(row.get(cols["一级科目"])) if "一级科目" in cols else ""
        l3 = _safe_str(row.get(cols["三级科目"])) if "三级科目" in cols else ""
        vend = ""
        if "供应商名称" in cols:
            vend = _safe_str(row.get(cols["供应商名称"]))
        elif "供应商" in cols:
            vend = _safe_str(row.get(cols["供应商"]))
        subj = _safe_str(row.get(cols["示例科目"])) if "示例科目" in cols else ""

        if not summ or not l2:
            skipped += 1
            continue
        if summ in seen:
            continue
        seen.add(summ)

        try:
            learn_one(
                taxonomy=tax,
                memory=memory,
                corrections=corrections,
                vendor_stats=vendor_stats,
                subject=subj,
                summary=summ,
                vendor=vend,
                new_l1=l1,
                new_l2=l2,
                new_l3=l3,
                source="class_master",
                notes="分类总表(摘要去重/三级)",
                also_summary_only=True,
            )
            learned += 1
        except Exception:
            errors += 1

    stats: Dict[str, Any] = {
        "learned": learned,
        "skipped": skipped,
        "errors": errors,
        "unique_summaries": len(seen),
    }

    if reclassify:
        engine = ClassifyEngine.from_data_dir(data_dir)
        stats["reclassify"] = reclassify_std_dir(Path(std_dir), engine)

    if reexport:
        dest = Path(out_path) if out_path else master_path
        engine = ClassifyEngine.from_data_dir(data_dir)
        path, n = export_class_master(std_dir, dest, data_dir, engine=engine)
        stats["reexport"] = str(path)
        stats["reexport_rows"] = n

    return stats
