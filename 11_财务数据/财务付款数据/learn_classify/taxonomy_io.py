# -*- coding: utf-8 -*-
"""
一级/二级科目体系：Excel ⇄ taxonomy.json

方便在 Excel 中增删改分类名称、启用状态、顺序，再导入引擎。
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .taxonomy import Taxonomy

L1_SHEET = "一级科目"
L2_SHEET = "二级科目"
L3_SHEET = "三级科目"
README_SHEET = "使用说明"
DEFAULT_NAME = "科目体系.xlsx"


def _slug_id(prefix: str, name: str, existing: set) -> str:
    """从中文名生成稳定-ish id，冲突则加后缀。"""
    base = re.sub(r"[^\w\u4e00-\u9fff]+", "_", name, flags=re.UNICODE).strip("_")
    if not base:
        base = "item"
    # 尽量用 ascii 片段；全中文时用 hash
    ascii_part = re.sub(r"[^\w]+", "_", base.encode("ascii", "ignore").decode() or "")
    if len(ascii_part) < 2:
        ascii_part = f"x{abs(hash(name)) % 100000}"
    cand = f"{prefix}_{ascii_part}"[:40]
    if cand not in existing:
        existing.add(cand)
        return cand
    i = 2
    while f"{cand}_{i}" in existing:
        i += 1
    out = f"{cand}_{i}"
    existing.add(out)
    return out


def export_taxonomy_excel(data_dir: Path, out_path: Optional[Path] = None) -> Path:
    """把 taxonomy.json 导出为易编辑 Excel。"""
    data_dir = Path(data_dir)
    tax = Taxonomy.load(data_dir / "taxonomy.json")
    out_path = Path(out_path) if out_path else data_dir / DEFAULT_NAME

    l1_rows = []
    for x in sorted(tax.level1.values(), key=lambda z: (z.order, z.id)):
        l1_rows.append(
            {
                "ID": x.id,
                "名称": x.name,
                "排序": x.order,
                "启用": "是" if x.active else "否",
                "备注": x.notes or "",
            }
        )

    l2_rows = []
    for x in sorted(tax.level2.values(), key=lambda z: (z.order, z.id)):
        parent_name = tax.level1[x.parent].name if x.parent in tax.level1 else x.parent
        l2_rows.append(
            {
                "ID": x.id,
                "名称": x.name,
                "所属一级ID": x.parent,
                "所属一级名称": parent_name,
                "排序": x.order,
                "启用": "是" if x.active else "否",
                "别名": ";".join(x.aliases or []),
                "备注": x.notes or "",
            }
        )

    l3_rows = []
    for x in sorted(tax.level3.values(), key=lambda z: (z.order, z.id)):
        parent_name = tax.level2[x.parent].name if x.parent in tax.level2 else x.parent
        l3_rows.append(
            {
                "ID": x.id,
                "名称": x.name,
                "所属二级ID": x.parent,
                "所属二级名称": parent_name,
                "排序": x.order,
                "启用": "是" if x.active else "否",
                "别名": ";".join(x.aliases or []),
                "备注": x.notes or "",
            }
        )

    wb = Workbook()
    # 说明
    ws0 = wb.active
    ws0.title = README_SHEET
    tips = [
        "【科目体系 · 三级分类】",
        "",
        "结构：一级 → 二级 → 三级（可选）",
        "  例：6_管理费 → 办公费 → 网络通讯费 / 邮寄快递费 / 办公物品采购 / 其他办公",
        "",
        "粒度：业务大类，不要型号规格（C30/HRB400 等）。",
        "",
        "1. 改完保存后：python 财务付款一体化.py import-taxonomy",
        "2. 导出：python 财务付款一体化.py export-taxonomy",
        "3. 刷新总表下拉：python 财务付款一体化.py export-class",
        "",
        "【一级】ID / 名称 / 排序 / 启用 / 备注",
        "【二级】所属一级名称或ID；别名用分号分隔",
        "【三级】所属二级名称或ID；无三级的二级可不建三级",
        "",
        f"当前版本: {tax.version}",
        f"待复核二级ID: {tax.pending_l2_id}",
    ]
    for i, t in enumerate(tips, 1):
        ws0.cell(i, 1, t)
    ws0.column_dimensions["A"].width = 90

    def _write_sheet(title: str, rows: List[Dict], headers: List[str]):
        ws = wb.create_sheet(title)
        fill = PatternFill("solid", fgColor="1F4E79")
        font = Font(color="FFFFFF", bold=True)
        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center")
        for r, row in enumerate(rows, 2):
            for c, h in enumerate(headers, 1):
                ws.cell(r, c, row.get(h, ""))
        for c in range(1, len(headers) + 1):
            ws.column_dimensions[chr(64 + c) if c <= 26 else "A"].width = 16
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        # fix widths
        widths = {
            "ID": 18,
            "名称": 22,
            "排序": 8,
            "启用": 8,
            "备注": 24,
            "所属一级ID": 16,
            "所属一级名称": 20,
            "别名": 28,
        }
        from openpyxl.utils import get_column_letter

        for c, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 14)

    _write_sheet(
        L1_SHEET,
        l1_rows,
        ["ID", "名称", "排序", "启用", "备注"],
    )
    _write_sheet(
        L2_SHEET,
        l2_rows,
        ["ID", "名称", "所属一级ID", "所属一级名称", "排序", "启用", "别名", "备注"],
    )
    _write_sheet(
        L3_SHEET,
        l3_rows,
        ["ID", "名称", "所属二级ID", "所属二级名称", "排序", "启用", "别名", "备注"],
    )

    # 元数据
    meta = wb.create_sheet("_meta")
    meta["A1"] = "version"
    meta["B1"] = tax.version
    meta["A2"] = "pending_l2_id"
    meta["B2"] = tax.pending_l2_id
    meta["A3"] = "default_l2_id"
    meta["B3"] = tax.default_l2_id
    meta.sheet_state = "hidden"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def import_taxonomy_excel(
    data_dir: Path,
    excel_path: Optional[Path] = None,
    *,
    backup: bool = True,
) -> Dict[str, Any]:
    """
    从 Excel 导入并覆盖 taxonomy.json。
    返回统计信息。
    """
    data_dir = Path(data_dir)
    excel_path = Path(excel_path) if excel_path else data_dir / DEFAULT_NAME
    if not excel_path.is_file():
        raise FileNotFoundError(f"科目体系文件不存在: {excel_path}")

    tax_path = data_dir / "taxonomy.json"
    old = Taxonomy.load(tax_path) if tax_path.is_file() else None

    if backup and tax_path.is_file():
        bak = data_dir / f"taxonomy_backup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.json"
        bak.write_text(tax_path.read_text(encoding="utf-8"), encoding="utf-8")

    l1_df = pd.read_excel(excel_path, sheet_name=L1_SHEET)
    l2_df = pd.read_excel(excel_path, sheet_name=L2_SHEET)
    try:
        l3_df = pd.read_excel(excel_path, sheet_name=L3_SHEET)
    except Exception:
        l3_df = pd.DataFrame()

    def colmap(df):
        return {str(c).strip(): c for c in df.columns}

    def yes(v) -> bool:
        s = str(v).strip().lower() if v is not None and not (isinstance(v, float) and pd.isna(v)) else "是"
        return s not in {"否", "no", "0", "false", "n", "停用", "禁用"}

    def cell(row, cmap, *names, default=""):
        for n in names:
            if n in cmap:
                v = row.get(cmap[n])
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return default
                return v
        return default

    # version from meta if present
    version = old.version if old else "1.0.0"
    pending = old.pending_l2_id if old else ""
    default_l2 = old.default_l2_id if old else ""
    try:
        meta = pd.read_excel(excel_path, sheet_name="_meta", header=None)
        for _, r in meta.iterrows():
            k = str(r.iloc[0]).strip() if len(r) > 0 else ""
            v = str(r.iloc[1]).strip() if len(r) > 1 else ""
            if k == "version" and v:
                version = v
            if k == "pending_l2_id" and v:
                pending = v
            if k == "default_l2_id" and v:
                default_l2 = v
    except Exception:
        pass

    existing_ids: set = set()
    level1: List[Dict[str, Any]] = []
    cmap1 = colmap(l1_df)
    name_to_l1_id: Dict[str, str] = {}

    for _, row in l1_df.iterrows():
        name = str(cell(row, cmap1, "名称", "name", default="")).strip()
        if not name:
            continue
        rid = str(cell(row, cmap1, "ID", "id", default="")).strip()
        if not rid:
            rid = _slug_id("L1", name, existing_ids)
        else:
            existing_ids.add(rid)
        order = cell(row, cmap1, "排序", "order", default=0)
        try:
            order = int(float(order))
        except Exception:
            order = 0
        active = yes(cell(row, cmap1, "启用", "active", default="是"))
        notes = str(cell(row, cmap1, "备注", "notes", default="") or "")
        level1.append(
            {"id": rid, "name": name, "order": order, "active": active, "notes": notes}
        )
        name_to_l1_id[name] = rid
        name_to_l1_id[name.lower()] = rid

    level2: List[Dict[str, Any]] = []
    cmap2 = colmap(l2_df)
    for _, row in l2_df.iterrows():
        name = str(cell(row, cmap2, "名称", "name", default="")).strip()
        if not name:
            continue
        parent_id = str(cell(row, cmap2, "所属一级ID", "parent", default="")).strip()
        parent_name = str(cell(row, cmap2, "所属一级名称", default="")).strip()
        if not parent_id and parent_name:
            parent_id = name_to_l1_id.get(parent_name) or name_to_l1_id.get(parent_name.lower()) or ""
        if not parent_id:
            raise ValueError(f"二级「{name}」缺少所属一级（填 所属一级名称 或 所属一级ID）")
        rid = str(cell(row, cmap2, "ID", "id", default="")).strip()
        if not rid:
            rid = _slug_id("L2", name, existing_ids)
        else:
            existing_ids.add(rid)
        order = cell(row, cmap2, "排序", "order", default=0)
        try:
            order = int(float(order))
        except Exception:
            order = 0
        active = yes(cell(row, cmap2, "启用", "active", default="是"))
        aliases_raw = str(cell(row, cmap2, "别名", "aliases", default="") or "")
        aliases = [a.strip() for a in aliases_raw.replace("；", ";").split(";") if a.strip()]
        notes = str(cell(row, cmap2, "备注", "notes", default="") or "")
        level2.append(
            {
                "id": rid,
                "name": name,
                "parent": parent_id,
                "order": order,
                "active": active,
                "aliases": aliases,
                "notes": notes,
            }
        )

    if not level1 or not level2:
        raise ValueError("一级或二级为空，请检查 Excel")

    # 三级
    level3: List[Dict[str, Any]] = []
    cmap3 = colmap(l3_df) if not l3_df.empty else {}
    name_to_l2_id = {x["name"]: x["id"] for x in level2}
    name_to_l2_id.update({x["name"].lower(): x["id"] for x in level2})
    if not l3_df.empty:
        for _, row in l3_df.iterrows():
            name = str(cell(row, cmap3, "名称", "name", default="")).strip()
            if not name:
                continue
            parent_id = str(cell(row, cmap3, "所属二级ID", "parent", default="")).strip()
            parent_name = str(cell(row, cmap3, "所属二级名称", default="")).strip()
            if not parent_id and parent_name:
                parent_id = name_to_l2_id.get(parent_name) or name_to_l2_id.get(
                    parent_name.lower()
                ) or ""
            if not parent_id:
                raise ValueError(f"三级「{name}」缺少所属二级")
            rid = str(cell(row, cmap3, "ID", "id", default="")).strip()
            if not rid:
                rid = _slug_id("L3", name, existing_ids)
            else:
                existing_ids.add(rid)
            order = cell(row, cmap3, "排序", "order", default=0)
            try:
                order = int(float(order))
            except Exception:
                order = 0
            active = yes(cell(row, cmap3, "启用", "active", default="是"))
            aliases_raw = str(cell(row, cmap3, "别名", "aliases", default="") or "")
            aliases = [
                a.strip()
                for a in aliases_raw.replace("；", ";").split(";")
                if a.strip()
            ]
            notes = str(cell(row, cmap3, "备注", "notes", default="") or "")
            level3.append(
                {
                    "id": rid,
                    "name": name,
                    "parent": parent_id,
                    "order": order,
                    "active": active,
                    "aliases": aliases,
                    "notes": notes,
                }
            )

    # pending / default
    l2_ids = {x["id"] for x in level2}
    l2_names = {x["name"]: x["id"] for x in level2}
    if pending not in l2_ids:
        pending = l2_names.get("待复核") or next(iter(l2_ids))
    if default_l2 not in l2_ids:
        default_l2 = pending

    data = {
        "version": version,
        "description": "财务付款三级分类树；Excel 科目体系导入",
        "default_l2_id": default_l2,
        "pending_l2_id": pending,
        "level1": level1,
        "level2": level2,
        "level3": level3,
    }
    # validate
    Taxonomy(data)
    tax_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "path": str(tax_path),
        "level1": len(level1),
        "level2": len(level2),
        "level3": len(level3),
        "version": version,
        "pending_l2_id": pending,
    }
