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
from .normalize import memory_key
from .review import reclassify_std_dir
from .store import CorrectionStore, MemoryStore, VendorStats
from .taxonomy import Taxonomy

MASTER_SHEET = "分类总表"
README_SHEET = "使用说明"

# 核心列（学习焦点）+ 辅助列
# 确认状态：确认 = 已认可，参与学习；待分类 = 需人工处理，不写入记忆
MASTER_COLUMNS = [
    "确认状态",
    "摘要",
    "一级科目",
    "二级科目",
    "三级科目",
    "供应商名称",
    # —— 以下为辅助 ——
    "出现次数",
    "命中层",
    "分类ID",
    "备注",
]

STATUS_CONFIRMED = "确认"
STATUS_PENDING = "待分类"
_CONFIRM_VALUES = {STATUS_CONFIRMED, "已确认", "是", "Y", "y", "1", "true", "True"}
# 不算有效分类、不写入记忆的二级名
_SKIP_L2 = {"", "待复核", "其他", "(空白)", "待分类", "—", "-"}


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

    for f in sorted(std_dir.rglob("*.xlsx")):
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
            # 仅学习「有实付金额」的摘要；空实付不进入分类总表/学习
            amt_raw = row.get(c_amt) if c_amt is not None else None
            try:
                from payment_hub.utils import has_actual_payment

                if not has_actual_payment(amt_raw):
                    continue
            except Exception:
                # 回退：空单元格跳过
                if amt_raw is None or (isinstance(amt_raw, float) and pd.isna(amt_raw)):
                    continue
                if str(amt_raw).strip() == "":
                    continue
            amt = 0.0
            if c_amt is not None:
                try:
                    amt = float(amt_raw if amt_raw is not None else 0)
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


def _is_confirmed(val: Any) -> bool:
    s = _safe_str(val)
    return s in _CONFIRM_VALUES or s == STATUS_CONFIRMED


def _is_learnable_l2(l2: str) -> bool:
    """二级是否可作为有效分类写入记忆。"""
    return _safe_str(l2) not in _SKIP_L2


def resolve_leaf_from_row(
    tax: Taxonomy,
    l1: str = "",
    l2: str = "",
    l3: str = "",
) -> str:
    """
    只解析「最后一级叶子科目」ID。

    优先级：三级 > 二级列里的三级名 > 二级 > 其它可解析叶子。
    一二级显示名由 taxonomy.display_triple(leaf) 向上补全，不必单独学习。
    """
    l1, l2, l3 = _safe_str(l1), _safe_str(l2), _safe_str(l3)
    # 1) 明确的三级
    if l3 and l3 not in _SKIP_L2:
        lid = tax.resolve_l3_id(l3) or tax.resolve_leaf_id(l3)
        if lid:
            return lid
    # 2) 二级列：可能是三级正式名，也可能是二级
    if l2 and l2 not in _SKIP_L2:
        lid = tax.resolve_l3_id(l2) or tax.resolve_l2_id(l2) or tax.resolve_leaf_id(l2)
        if lid:
            return lid
    # 3) 一级单独不能当叶子
    return ""


def _is_learnable_leaf(tax: Taxonomy, leaf_id: str) -> bool:
    if not leaf_id:
        return False
    if leaf_id == getattr(tax, "pending_l2_id", "") or leaf_id == "pending":
        return False
    l1, l2, l3 = tax.display_triple(leaf_id)
    leaf_name = l3 or l2
    if leaf_name in _SKIP_L2:
        return False
    return True


def load_labels_from_master(path: Path) -> Dict[str, Dict[str, str]]:
    """
    从已有总表读取**有效分类**行（有二级科目即可，不要求「确认」）：
    摘要 → 一/二/三级/供应商/备注/确认状态。
    再次 export 时保留，避免被自动分类覆盖。
    """
    path = Path(path)
    if not path.is_file():
        return {}
    try:
        df = pd.read_excel(path, sheet_name=MASTER_SHEET)
    except Exception:
        try:
            df = pd.read_excel(path, sheet_name=0)
        except Exception:
            return {}
    cols = {str(c).strip(): c for c in df.columns}
    if "摘要" not in cols:
        return {}
    status_col = cols.get("确认状态") or cols.get("确认")
    out: Dict[str, Dict[str, str]] = {}
    for _, row in df.iterrows():
        summ = _safe_str(row.get(cols["摘要"]))
        if not summ:
            continue
        l2 = _safe_str(row.get(cols["二级科目"])) if "二级科目" in cols else ""
        l3 = _safe_str(row.get(cols["三级科目"])) if "三级科目" in cols else ""
        if not _is_learnable_l2(l2) and not l3:
            continue
        st = _safe_str(row.get(status_col)) if status_col else ""
        out[summ] = {
            "一级科目": _safe_str(row.get(cols["一级科目"])) if "一级科目" in cols else "",
            "二级科目": l2,
            "三级科目": l3,
            "供应商名称": _safe_str(row.get(cols["供应商名称"])) if "供应商名称" in cols else "",
            "备注": _safe_str(row.get(cols["备注"])) if "备注" in cols else "",
            "确认状态": st if st else STATUS_CONFIRMED,
        }
    return out


# 兼容旧名
load_confirmed_from_master = load_labels_from_master


def build_deduped_master(
    std_dir: Path,
    engine: Optional[ClassifyEngine] = None,
    *,
    previous_confirmed: Optional[Dict[str, Dict[str, str]]] = None,
) -> pd.DataFrame:
    """
    按「摘要」完全相同去重。

    - 上一版总表里已有有效二级/三级：沿用（人工调整不被自动覆盖）
    - 其余：用文件/引擎建议
    """
    previous_labels = previous_confirmed or {}
    lines = _collect_raw_lines(std_dir)
    if not lines:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for ln in lines:
        groups[ln["摘要"]].append(ln)

    rows: List[Dict[str, Any]] = []
    for summ, items in groups.items():
        n = len(items)

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

        layer_counter = Counter(x["命中层"] for x in items if x.get("命中层"))
        id_counter = Counter(x["分类ID"] for x in items if x.get("分类ID"))
        subj_counter = Counter(x["科目"] for x in items if x.get("科目"))

        cid = id_counter.most_common(1)[0][0] if id_counter else ""
        layer = layer_counter.most_common(1)[0][0] if layer_counter else ""
        sample_subj = subj_counter.most_common(1)[0][0] if subj_counter else ""
        note = ""
        status = STATUS_PENDING

        # 上一版总表已有有效分类：保留人工调整
        if summ in previous_labels:
            prev = previous_labels[summ]
            if prev.get("二级科目") or prev.get("三级科目"):
                l1 = prev.get("一级科目") or l1
                l2 = prev.get("二级科目") or l2
                l3 = prev.get("三级科目") or l3
                if prev.get("供应商名称"):
                    vendor = prev["供应商名称"]
                note = prev.get("备注") or ""
                status = prev.get("确认状态") or STATUS_CONFIRMED
                if not _is_confirmed(status):
                    status = STATUS_CONFIRMED  # 来自总表的有效分类视为已采用
                layer = "master"
                cid = cid or ""
        else:
            # 引擎补全建议
            if engine is not None and (not l1 or not l2):
                r = engine.classify(sample_subj, summ, vendor)
                l1 = l1 or r.l1_name
                l2 = l2 or r.l2_name
                l3 = l3 or r.l3_name
                cid = cid or (r.l3_id or r.l2_id)
                layer = layer or r.layer
            elif engine is not None and l2 and not l3:
                r = engine.classify(sample_subj, summ, vendor)
                if r.l2_name == l2 and r.l3_name:
                    l3 = r.l3_name
                    cid = r.l3_id or cid
            # 无有效二级 → 待分类
            if not _is_learnable_l2(l2):
                status = STATUS_PENDING

        rows.append(
            {
                "确认状态": status,
                "摘要": summ,
                "一级科目": l1,
                "二级科目": l2,
                "三级科目": l3,
                "供应商名称": vendor,
                "出现次数": n,
                "命中层": layer,
                "分类ID": cid,
                "备注": note,
            }
        )

    df = pd.DataFrame(rows, columns=MASTER_COLUMNS)
    # 待分类优先，再按出现次数降序
    df["_ord"] = df["确认状态"].map({STATUS_PENDING: 0, STATUS_CONFIRMED: 1}).fillna(0)
    df = df.sort_values(["_ord", "出现次数"], ascending=[True, False]).drop(columns=["_ord"])
    return df.reset_index(drop=True)


def _apply_master_format(path: Path, data_dir: Path) -> None:
    wb = load_workbook(path)

    if README_SHEET in wb.sheetnames:
        del wb[README_SHEET]
    ws_r = wb.create_sheet(README_SHEET, 0)
    tips = [
        "【付款分类总表 · 叶子科目变更才学习】",
        "",
        "核心：只看「最后一级科目」是否变化。一二级随叶子自动带出，不必单独记。",
        "· 有三级时：叶子 = 三级（如 网络通讯费）",
        "· 无三级时：叶子 = 二级（如 差旅费）",
        "· 叶子与记忆相同 → 跳过；叶子变了 → 写入记忆",
        "",
        "1. 怎么改分类",
        "   · 优先改最末级（三级有则改三级，否则改二级）",
        "   · 一级/二级下拉会跟着体系联动，仅辅助查看",
        "   · 叶子为空 / 待复核 → 不学习",
        "",
        "2. 命令（默认入口 = apply-class）",
        "   · python 财务付款一体化.py              # 等同 apply-class",
        "   · export-class  导出/刷新本表（表内有效分类会保留）",
        "   · apply-class   仅学叶子变更 → 写学习文件 → 重刷 B_std → 再导出",
        "",
        "3. 学习生成的文件",
        "   · 学习记录.xlsx   本次叶子变更 + 记忆库摘要",
        "   · memory.jsonl    引擎记忆（下次自动命中）",
        "   · corrections.jsonl 纠错日志",
        "",
        "4. 建议节奏",
        "   打开表 → 改高频摘要的末级科目 → 保存 → 直接运行一体化.py → 循环",
        "",
        "5. 三级示例：管理费 → 办公费 → 网络通讯费（学的是「网络通讯费」这个叶子）",
        "   无三级的二级：三级列可留空。粒度到业务大类，不到型号规格。",
        "",
        "6. 类目字典：export-taxonomy / import-taxonomy（科目体系.xlsx）",
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
        "确认状态": 10,
        "摘要": 48,
        "一级科目": 14,
        "二级科目": 14,
        "三级科目": 14,
        "供应商名称": 22,
        "出现次数": 10,
        "命中层": 12,
        "分类ID": 14,
        "备注": 18,
    }
    for name, w in widths.items():
        if name in headers:
            ws.column_dimensions[get_column_letter(headers[name])].width = w

    # 核心列高亮
    core_fill = PatternFill("solid", fgColor="FFF2CC")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")  # 确认-绿
    pending_fill = PatternFill("solid", fgColor="FCE4D6")  # 待分类-橙
    for name in ("确认状态", "摘要", "一级科目", "二级科目", "三级科目", "供应商名称", "备注"):
        if name not in headers:
            continue
        col = headers[name]
        for r in range(2, ws.max_row + 1):
            if name == "确认状态":
                val = str(ws.cell(r, col).value or "")
                ws.cell(r, col).fill = ok_fill if val == STATUS_CONFIRMED else pending_fill
            else:
                ws.cell(r, col).fill = core_fill

    last = max(ws.max_row, 2) + 200

    # 确认状态下拉
    if "确认状态" in headers:
        c_st = headers["确认状态"]
        dv_st = DataValidation(
            type="list",
            formula1=f'"{STATUS_CONFIRMED},{STATUS_PENDING}"',
            allow_blank=False,
            showDropDown=False,
            showErrorMessage=True,
            errorTitle="确认状态",
            error="请选择：确认 或 待分类",
            promptTitle="确认状态",
            prompt="确认=写入学习；待分类=仅人工暂存",
            showInputMessage=True,
        )
        dv_st.add(f"{get_column_letter(c_st)}2:{get_column_letter(c_st)}{last}")
        ws.add_data_validation(dv_st)

    # 一/二/三级下拉
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
    preserve_path: Optional[Path] = None,
) -> Tuple[Path, int]:
    """
    导出总表。preserve_path / out_path 上已有**有效二级**的行会保留其分类标签。
    """
    data_dir = Path(data_dir)
    out_path = Path(out_path)
    if engine is None:
        engine = ClassifyEngine.from_data_dir(data_dir)

    # 合并多来源总表标签（有效二级即保留，不要求「确认」）
    prev: Dict[str, Dict[str, str]] = {}
    for p in (preserve_path, out_path):
        if p and Path(p).is_file():
            prev.update(load_labels_from_master(Path(p)))

    df = build_deduped_master(std_dir, engine=engine, previous_confirmed=prev)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    n_ok = int((df["确认状态"] == STATUS_CONFIRMED).sum()) if len(df) else 0
    n_pending = int((df["确认状态"] == STATUS_PENDING).sum()) if len(df) else 0
    n_learnable = 0
    if len(df):
        n_learnable = int(df["二级科目"].map(_is_learnable_l2).sum())
    print(f"   有效分类 {n_learnable} 条 · 表内确认标记 {n_ok} · 待分类标记 {n_pending}")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=MASTER_SHEET, index=False)

    _apply_master_format(out_path, data_dir)
    return out_path, len(df)


def _export_learn_report(
    data_dir: Path,
    out_path: Path,
    *,
    session_rows: List[Dict[str, Any]],
    memory: MemoryStore,
    tax: Taxonomy,
    stats: Dict[str, Any],
) -> Path:
    """生成可读的学习记录 Excel（本次学习 + 记忆库摘要 + 文件路径）。"""
    from datetime import datetime

    data_dir = Path(data_dir)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df_session = pd.DataFrame(
        session_rows
        or [
            {
                "摘要": "（本次无确认行写入）",
                "一级科目": "",
                "二级科目": "",
                "三级科目": "",
                "供应商名称": "",
                "分类ID": "",
                "来源": "",
                "学习时间": ts,
            }
        ]
    )

    mem_rows: List[Dict[str, Any]] = []
    for e in sorted(memory.entries.values(), key=lambda x: x.updated_at or "", reverse=True):
        l1, l2, l3 = tax.display_triple(e.l2_id)
        mem_rows.append(
            {
                "摘要": e.summary,
                "供应商名称": e.vendor,
                "一级科目": l1,
                "二级科目": l2,
                "三级科目": l3,
                "分类ID": e.l2_id,
                "来源": e.source,
                "置信度": e.confidence,
                "命中次数": e.hit_count,
                "更新时间": e.updated_at,
                "备注": e.notes,
            }
        )
    df_mem = pd.DataFrame(mem_rows)

    paths_info = [
        {"说明": "本次学习时间", "路径或内容": ts},
        {"说明": "本次叶子变更条数", "路径或内容": stats.get("learned", 0)},
        {"说明": "叶子未变跳过", "路径或内容": stats.get("skipped_unchanged", 0)},
        {"说明": "无效叶子跳过", "路径或内容": stats.get("skipped_invalid", 0)},
        {"说明": "记忆库(jsonl)", "路径或内容": str(data_dir / "memory.jsonl")},
        {"说明": "纠错日志(jsonl)", "路径或内容": str(data_dir / "corrections.jsonl")},
        {"说明": "供应商统计", "路径或内容": str(data_dir / "vendor_stats.json")},
        {"说明": "科目体系", "路径或内容": str(data_dir / "taxonomy.json")},
        {"说明": "分类总表", "路径或内容": str(data_dir / "付款分类总表.xlsx")},
        {"说明": "本学习记录", "路径或内容": str(out_path)},
        {
            "说明": "用法",
            "路径或内容": "只改末级科目 → 保存 → apply-class（仅叶子变化才写入记忆，一二级自动带出）",
        },
    ]
    df_paths = pd.DataFrame(paths_info)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_session.to_excel(writer, sheet_name="本次学习", index=False)
        df_mem.to_excel(writer, sheet_name="记忆库摘要", index=False)
        df_paths.to_excel(writer, sheet_name="文件路径", index=False)

    return out_path


def apply_class_master(
    master_path: Path,
    std_dir: Path,
    data_dir: Path,
    *,
    reclassify: bool = True,
    reexport: bool = True,
    out_path: Optional[Path] = None,
    only_confirmed: bool = False,
    learn_log_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    从分类总表学习：只关心 **最后一级叶子科目是否变化**。

    - 有三级 → 叶子=三级；无三级 → 叶子=二级
    - 一二级由科目体系按叶子向上补全，不单独作为学习条件
    - 叶子与现有记忆相同 → 跳过；不同 → 写入记忆
    - only_confirmed=True 时额外要求「确认状态=确认」
    """
    from datetime import datetime

    master_path = Path(master_path)
    data_dir = Path(data_dir)
    if not master_path.is_file():
        raise FileNotFoundError(f"分类总表不存在: {master_path}")

    try:
        df = pd.read_excel(master_path, sheet_name=MASTER_SHEET)
    except Exception:
        df = pd.read_excel(master_path, sheet_name=0)

    cols = {str(c).strip(): c for c in df.columns}
    if "摘要" not in cols:
        raise ValueError(f"分类总表需含「摘要」列，当前={list(df.columns)}")
    if "二级科目" not in cols and "三级科目" not in cols:
        raise ValueError(
            f"分类总表需含「二级科目」或「三级科目」，当前={list(df.columns)}"
        )

    tax = Taxonomy.load(data_dir / "taxonomy.json")
    memory = MemoryStore(data_dir / "memory.jsonl")
    corrections = CorrectionStore(data_dir / "corrections.jsonl")
    vendor_stats = VendorStats(data_dir / "vendor_stats.json")

    learned = 0
    skipped = 0
    skipped_invalid = 0
    skipped_unchanged = 0
    skipped_pending = 0
    errors = 0
    seen = set()
    session_rows: List[Dict[str, Any]] = []
    status_col = cols.get("确认状态") or cols.get("确认")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for _, row in df.iterrows():
        summ = _safe_str(row.get(cols["摘要"]))
        l2 = _safe_str(row.get(cols["二级科目"])) if "二级科目" in cols else ""
        l1 = _safe_str(row.get(cols["一级科目"])) if "一级科目" in cols else ""
        l3 = _safe_str(row.get(cols["三级科目"])) if "三级科目" in cols else ""
        vend = ""
        if "供应商名称" in cols:
            vend = _safe_str(row.get(cols["供应商名称"]))
        elif "供应商" in cols:
            vend = _safe_str(row.get(cols["供应商"]))
        status = _safe_str(row.get(status_col)) if status_col else ""

        if not summ:
            skipped += 1
            continue
        if only_confirmed and not _is_confirmed(status):
            skipped_pending += 1
            continue
        if summ in seen:
            continue
        seen.add(summ)

        # 只解析叶子：三级优先，否则二级（一二级跟着叶子走）
        leaf_id = resolve_leaf_from_row(tax, l1, l2, l3)
        if not _is_learnable_leaf(tax, leaf_id):
            skipped_invalid += 1
            continue

        # 与记忆比较：叶子未变则不学
        mkey = memory_key("", summ, None)
        old_entry = memory.get(mkey)
        old_leaf = (old_entry.l2_id if old_entry else "") or ""
        if old_leaf == leaf_id:
            skipped_unchanged += 1
            continue

        l1_name, l2_name, l3_name = tax.display_triple(leaf_id)
        leaf_name = l3_name or l2_name
        old_l1, old_l2, old_l3 = tax.display_triple(old_leaf) if old_leaf else ("", "", "")
        old_leaf_name = (old_l3 or old_l2) if old_leaf else ""

        try:
            learn_one(
                taxonomy=tax,
                memory=memory,
                corrections=corrections,
                vendor_stats=vendor_stats,
                subject="",
                summary=summ,
                vendor=vend,
                new_l1=l1_name,
                new_l2=l2_name,
                new_l3=l3_name,
                old_l2_id=old_leaf,
                source="class_master_leaf",
                notes=f"叶子变更: {old_leaf_name or '∅'} → {leaf_name}",
                also_summary_only=True,
            )
            session_rows.append(
                {
                    "摘要": summ,
                    "原叶子科目": old_leaf_name or "（无）",
                    "新叶子科目": leaf_name,
                    "一级科目": l1_name,
                    "二级科目": l2_name,
                    "三级科目": l3_name,
                    "供应商名称": vend,
                    "分类ID": leaf_id,
                    "原分类ID": old_leaf,
                    "来源": "class_master_leaf",
                    "学习时间": ts,
                }
            )
            learned += 1
        except Exception:
            errors += 1

    stats: Dict[str, Any] = {
        "learned": learned,
        "skipped": skipped,
        "skipped_invalid": skipped_invalid,
        "skipped_unchanged": skipped_unchanged,
        "skipped_pending": skipped_pending,
        "errors": errors,
        "unique_scanned": len(seen),
        "only_confirmed": only_confirmed,
        "learn_mode": "leaf_change",
        "memory_file": str(data_dir / "memory.jsonl"),
        "corrections_file": str(data_dir / "corrections.jsonl"),
        "vendor_stats_file": str(data_dir / "vendor_stats.json"),
    }
    print(
        f"   叶子变更学习 {learned} 条 · 叶子未变跳过 {skipped_unchanged} · "
        f"无效叶子 {skipped_invalid} · 确认过滤 {skipped_pending} · "
        f"缺摘要 {skipped} · 失败 {errors}"
    )
    print(f"   记忆文件: {stats['memory_file']}")
    print(f"   纠错日志: {stats['corrections_file']}")

    # 可读学习报告（始终生成，便于核对）
    log_path = Path(learn_log_path) if learn_log_path else data_dir / "学习记录.xlsx"
    try:
        report = _export_learn_report(
            data_dir,
            log_path,
            session_rows=session_rows,
            memory=memory,
            tax=tax,
            stats=stats,
        )
        stats["learn_log"] = str(report)
        print(f"   学习记录: {report}")
    except Exception as e:
        print(f"  ⚠️ 学习记录导出失败: {e}")
        stats["learn_log_error"] = str(e)

    if reclassify:
        engine = ClassifyEngine.from_data_dir(data_dir)
        stats["reclassify"] = reclassify_std_dir(Path(std_dir), engine)

    if reexport:
        dest = Path(out_path) if out_path else master_path
        engine = ClassifyEngine.from_data_dir(data_dir)
        # 保留本次总表里所有确认行（含刚学的 + 未改动的确认）
        path, n = export_class_master(
            std_dir,
            dest,
            data_dir,
            engine=engine,
            preserve_path=master_path,
        )
        stats["reexport"] = str(path)
        stats["reexport_rows"] = n

    return stats


PENDING_SHEET = "仍待人工_高频"


def apply_dropdowns_to_pending_report(
    report_path: Path,
    data_dir: Path,
    *,
    sheet_name: str = PENDING_SHEET,
) -> int:
    """
    给修正记录里的「仍待人工_高频」加确认状态 + 一/二/三级联动下拉。
    返回应用了下拉的工作表数量。
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    report_path = Path(report_path)
    data_dir = Path(data_dir)
    if not report_path.is_file():
        raise FileNotFoundError(f"报告不存在: {report_path}")

    wb = load_workbook(report_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"未找到工作表: {sheet_name}，现有={wb.sheetnames}")

    ws = wb[sheet_name]
    ws.freeze_panes = "A2"
    if ws.max_row and ws.max_column:
        ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    core_fill = PatternFill("solid", fgColor="FFF2CC")
    pending_fill = PatternFill("solid", fgColor="FCE4D6")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")

    headers = {
        str(ws.cell(1, c).value).strip(): c
        for c in range(1, (ws.max_column or 1) + 1)
        if ws.cell(1, c).value is not None
    }
    for c in range(1, (ws.max_column or 1) + 1):
        cell = ws.cell(1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    widths = {
        "确认状态": 10,
        "摘要": 48,
        "出现次数": 10,
        "一级科目": 14,
        "二级科目": 14,
        "三级科目": 14,
        "供应商名称": 22,
        "命中层": 12,
        "分类ID": 14,
        "备注": 18,
        "建议操作": 36,
    }
    for name, w in widths.items():
        if name in headers:
            ws.column_dimensions[get_column_letter(headers[name])].width = w

    last_data = max(ws.max_row or 2, 2)
    last = last_data + 200

    # 预留空行便于继续填
    if "确认状态" in headers:
        for r in range(last_data + 1, min(last_data + 50, last) + 1):
            if not ws.cell(r, headers["确认状态"]).value:
                ws.cell(r, headers["确认状态"], STATUS_PENDING)

    for name in ("确认状态", "摘要", "一级科目", "二级科目", "三级科目", "供应商名称", "备注"):
        if name not in headers:
            continue
        col = headers[name]
        for r in range(2, last_data + 1):
            if name == "确认状态":
                val = str(ws.cell(r, col).value or "")
                ws.cell(r, col).fill = (
                    ok_fill if _is_confirmed(val) else pending_fill
                )
            else:
                ws.cell(r, col).fill = core_fill

    # 确认状态下拉
    if "确认状态" in headers:
        c_st = headers["确认状态"]
        # 清掉旧的确认验证
        keep = []
        letter = get_column_letter(c_st)
        for dv in ws.data_validations.dataValidation:
            sq = str(dv.sqref or "")
            if letter in sq and (dv.promptTitle == "确认状态" or "确认" in str(dv.formula1 or "")):
                continue
            keep.append(dv)
        ws.data_validations.dataValidation = keep
        dv_st = DataValidation(
            type="list",
            formula1=f'"{STATUS_CONFIRMED},{STATUS_PENDING}"',
            allow_blank=False,
            showDropDown=False,
            showErrorMessage=True,
            errorTitle="确认状态",
            error="请选择：确认 或 待分类",
            showInputMessage=True,
            promptTitle="确认状态",
            prompt="改完一/二/三级后选「确认」，再运行 apply-pending-fixes",
        )
        dv_st.add(f"{letter}2:{letter}{last}")
        ws.add_data_validation(dv_st)

    # 一/二/三级联动
    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from payment_hub.category_dropdowns import (
        apply_category_dropdowns,
        apply_dropdowns_to_worksheet,
    )

    n = apply_category_dropdowns(wb, data_dir=data_dir, sheet_names=[sheet_name])
    apply_dropdowns_to_worksheet(ws, max_row=last)
    wb.save(report_path)
    return n


def merge_pending_fixes_into_master(
    pending_path: Path,
    master_path: Path,
    data_dir: Path,
    *,
    only_confirmed: bool = True,
    sheet_name: str = PENDING_SHEET,
) -> Dict[str, Any]:
    """
    把「仍待人工_高频」里已改好的行写回付款分类总表，并学习叶子变更。
    默认仅「确认状态=确认」的行。
    """
    pending_path = Path(pending_path)
    master_path = Path(master_path)
    data_dir = Path(data_dir)
    if not pending_path.is_file():
        raise FileNotFoundError(f"待人工表不存在: {pending_path}")
    if not master_path.is_file():
        raise FileNotFoundError(f"分类总表不存在: {master_path}")

    try:
        pend = pd.read_excel(pending_path, sheet_name=sheet_name)
    except Exception:
        pend = pd.read_excel(pending_path, sheet_name=0)

    try:
        master = pd.read_excel(master_path, sheet_name=MASTER_SHEET)
    except Exception:
        master = pd.read_excel(master_path, sheet_name=0)

    def S(v: Any) -> str:
        return _safe_str(v)

    pcols = {str(c).strip(): c for c in pend.columns}
    if "摘要" not in pcols:
        raise ValueError(f"待人工表需含「摘要」列，当前={list(pend.columns)}")

    status_col = pcols.get("确认状态") or pcols.get("确认")
    updates: Dict[str, Dict[str, str]] = {}
    for _, row in pend.iterrows():
        summ = S(row.get(pcols["摘要"]))
        if not summ:
            continue
        st = S(row.get(status_col)) if status_col else STATUS_CONFIRMED
        if only_confirmed and not _is_confirmed(st):
            continue
        l2 = S(row.get(pcols["二级科目"])) if "二级科目" in pcols else ""
        l3 = S(row.get(pcols["三级科目"])) if "三级科目" in pcols else ""
        l1 = S(row.get(pcols["一级科目"])) if "一级科目" in pcols else ""
        if not _is_learnable_l2(l2) and not l3:
            continue
        updates[summ] = {
            "一级科目": l1,
            "二级科目": l2,
            "三级科目": l3,
            "供应商名称": S(row.get(pcols["供应商名称"])) if "供应商名称" in pcols else "",
            "确认状态": STATUS_CONFIRMED,
            "备注": S(row.get(pcols["备注"])) if "备注" in pcols else "",
        }

    if not updates:
        return {"merged": 0, "message": "没有可合并的确认行（请改分类并把确认状态设为「确认」）"}

    tax = Taxonomy.load(data_dir / "taxonomy.json")
    mcols = {str(c).strip(): c for c in master.columns}
    if "摘要" not in mcols:
        raise ValueError("分类总表缺摘要列")

    merged = 0
    for i, row in master.iterrows():
        summ = S(row.get(mcols["摘要"]))
        if summ not in updates:
            continue
        u = updates[summ]
        # 按叶子解析对齐
        leaf = resolve_leaf_from_row(tax, u["一级科目"], u["二级科目"], u["三级科目"])
        if leaf:
            l1, l2, l3 = tax.display_triple(leaf)
            u = {
                **u,
                "一级科目": l1,
                "二级科目": l2,
                "三级科目": l3 or "",
                "分类ID": leaf,
            }
        for k, v in u.items():
            col = mcols.get(k, k)
            if col not in master.columns:
                master[col] = ""
            master.at[i, col] = v
        merged += 1
        del updates[summ]

    # 总表没有的摘要：追加
    appended = 0
    for summ, u in updates.items():
        leaf = resolve_leaf_from_row(tax, u["一级科目"], u["二级科目"], u["三级科目"])
        if leaf:
            l1, l2, l3 = tax.display_triple(leaf)
            u = {**u, "一级科目": l1, "二级科目": l2, "三级科目": l3 or "", "分类ID": leaf}
        rec = {c: "" for c in master.columns}
        rec["摘要"] = summ
        for k, v in u.items():
            if k in rec:
                rec[k] = v
            else:
                master[k] = ""
                rec[k] = v
        if "出现次数" in rec and not rec["出现次数"]:
            rec["出现次数"] = 1
        if "命中层" in rec:
            rec["命中层"] = "pending_fix"
        master = pd.concat([master, pd.DataFrame([rec])], ignore_index=True)
        appended += 1
        merged += 1

    master_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(master_path, engine="openpyxl") as w:
        master.to_excel(w, sheet_name=MASTER_SHEET, index=False)
    try:
        _apply_master_format(master_path, data_dir)
    except Exception as e:
        print(f"  ⚠️ 总表格式/下拉刷新失败: {e}")

    return {
        "merged": merged,
        "appended": appended,
        "master": str(master_path),
    }
