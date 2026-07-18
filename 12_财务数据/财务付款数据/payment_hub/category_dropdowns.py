# -*- coding: utf-8 -*-
"""
为标准明细 Excel 的「一级科目 / 二级科目」列添加下拉选项。

- 一级：taxonomy 中全部一级显示名
- 二级：随同行一级联动（INDIRECT + 命名区域）
- 选项源写在隐藏表「_分类选项」中，随 taxonomy 更新
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

CODE_ROOT = Path(__file__).resolve().parent.parent
OPTIONS_SHEET = "_分类选项"
DETAIL_SHEETS = ("明细数据", "明细数据_万泰铢")


def _ensure_learn_path() -> None:
    root = str(CODE_ROOT)
    if root not in sys.path:
        sys.path.insert(0, root)


def load_taxonomy_options(data_dir: Optional[Path] = None):
    """返回 (l1_names, l1_id_to_l2_names, l1_name_to_id)。仅 active。"""
    _ensure_learn_path()
    from learn_classify.taxonomy import Taxonomy

    data_dir = Path(data_dir) if data_dir else CODE_ROOT / "learn_classify" / "data"
    tax = Taxonomy.load(data_dir / "taxonomy.json")

    l1_sorted = sorted(
        [x for x in tax.level1.values() if x.active],
        key=lambda z: (z.order, z.id),
    )
    l1_names = [x.name for x in l1_sorted]
    l1_name_to_id = {x.name: x.id for x in l1_sorted}

    l1_id_to_l2: Dict[str, List[str]] = {x.id: [] for x in l1_sorted}
    for l2 in sorted(tax.level2.values(), key=lambda z: (z.order, z.id)):
        if not l2.active:
            continue
        if l2.parent not in l1_id_to_l2:
            continue
        # 显示名去重保序
        if l2.name not in l1_id_to_l2[l2.parent]:
            l1_id_to_l2[l2.parent].append(l2.name)
        for al in l2.aliases:
            if al and al not in l1_id_to_l2[l2.parent]:
                l1_id_to_l2[l2.parent].append(al)

    return l1_names, l1_id_to_l2, l1_name_to_id, tax


def _clear_old_category_validations(ws: Worksheet, col_indices: Sequence[int]) -> None:
    """移除覆盖指定列的旧 list 校验（避免重复叠加）。"""
    if not col_indices:
        return
    target_cols = set(col_indices)
    keep = []
    for dv in ws.data_validations.dataValidation:
        # 粗略：若 sqref 涉及目标列则丢弃（将重建）
        sq = str(dv.sqref or "")
        drop = False
        for col in target_cols:
            letter = get_column_letter(col)
            # 匹配 C2:C100 / $C$2 等
            if re.search(rf"(^|\s|,)\$?{letter}\$?\d", sq, re.I):
                drop = True
                break
            if re.search(rf"(^|\s|,)\$?{letter}\$?\d*:\$?{letter}\$?\d*", sq, re.I):
                drop = True
                break
        if not drop:
            keep.append(dv)
    ws.data_validations.dataValidation = keep


def _remove_defined_names_with_prefix(wb: Workbook, prefix: str = "PAYL2_") -> None:
    """删除本模块创建的旧命名区域。"""
    # openpyxl 3.x: defined_names 可迭代
    to_del = []
    try:
        for name in list(wb.defined_names):
            # name may be DefinedName or str depending on version
            n = name if isinstance(name, str) else getattr(name, "name", str(name))
            if str(n).startswith(prefix) or str(n) in {"PAY_L1_LIST", "PAY_L1_MAP"}:
                to_del.append(str(n))
    except Exception:
        # fallback: defined_names.definedName list
        try:
            for dn in list(getattr(wb.defined_names, "definedName", [])):
                n = getattr(dn, "name", "")
                if n.startswith(prefix) or n in {"PAY_L1_LIST", "PAY_L1_MAP"}:
                    to_del.append(n)
        except Exception:
            return

    for n in to_del:
        try:
            del wb.defined_names[n]
        except Exception:
            try:
                wb.defined_names.delete(n)
            except Exception:
                pass


def _write_options_sheet(
    wb: Workbook,
    l1_names: List[str],
    l1_id_to_l2: Dict[str, List[str]],
    l1_name_to_id: Dict[str, str],
) -> None:
    """写入/重建隐藏选项表与命名区域。"""
    if OPTIONS_SHEET in wb.sheetnames:
        del wb[OPTIONS_SHEET]
    ws = wb.create_sheet(OPTIONS_SHEET)

    # A: 一级列表
    ws["A1"] = "一级科目"
    for i, name in enumerate(l1_names, start=2):
        ws.cell(i, 1, name)
    l1_end = 1 + len(l1_names)

    # B-C: 一级名 → 命名区域后缀(id) 映射（供 VLOOKUP）
    ws["B1"] = "一级科目"
    ws["C1"] = "命名键"
    for i, name in enumerate(l1_names, start=2):
        ws.cell(i, 2, name)
        ws.cell(i, 3, l1_name_to_id[name])  # e.g. L1_sub
    map_end = 1 + len(l1_names)

    # 从 E 列起：每个一级一列二级列表，表头为 L1_id
    # E1 = L1_sub, E2:E = l2 names
    col = 5  # E
    range_defs: List[Tuple[str, str]] = []  # (defined_name, attr_text)
    for name in l1_names:
        lid = l1_name_to_id[name]
        l2s = l1_id_to_l2.get(lid) or []
        ws.cell(1, col, lid)
        for r, l2 in enumerate(l2s, start=2):
            ws.cell(r, col, l2)
        end_row = max(2, 1 + len(l2s))
        if l2s:
            letter = get_column_letter(col)
            # 命名区域：PAYL2_L1_sub -> '_分类选项'!$E$2:$E$n
            def_name = f"PAYL2_{lid}"
            attr = f"'{OPTIONS_SHEET}'!${letter}$2:${letter}${end_row}"
            range_defs.append((def_name, attr))
        else:
            # 至少一个空选项避免 INDIRECT 失败
            ws.cell(2, col, "")
            letter = get_column_letter(col)
            def_name = f"PAYL2_{lid}"
            attr = f"'{OPTIONS_SHEET}'!${letter}$2:${letter}$2"
            range_defs.append((def_name, attr))
        col += 1

    # 对照表（便于人工查看，放在右侧较远列）
    ref_col = col + 1
    ws.cell(1, ref_col, "对照_一级")
    ws.cell(1, ref_col + 1, "对照_二级")
    rr = 2
    for name in l1_names:
        lid = l1_name_to_id[name]
        for l2 in l1_id_to_l2.get(lid) or []:
            ws.cell(rr, ref_col, name)
            ws.cell(rr, ref_col + 1, l2)
            rr += 1

    ws.sheet_state = "hidden"

    # 定义名
    _remove_defined_names_with_prefix(wb, "PAYL2_")
    for extra in ("PAY_L1_LIST", "PAY_L1_MAP"):
        try:
            del wb.defined_names[extra]
        except Exception:
            pass

    def _add_name(n: str, attr: str) -> None:
        try:
            # openpyxl 3.1
            wb.defined_names[n] = DefinedName(n, attr_text=attr)
        except Exception:
            try:
                wb.defined_names.add(DefinedName(name=n, attr_text=attr))
            except Exception:
                pass

    if l1_names:
        _add_name("PAY_L1_LIST", f"'{OPTIONS_SHEET}'!$A$2:$A${l1_end}")
        _add_name("PAY_L1_MAP", f"'{OPTIONS_SHEET}'!$B$2:$C${map_end}")
    for def_name, attr in range_defs:
        _add_name(def_name, attr)


def _header_map(ws: Worksheet) -> Dict[str, int]:
    m = {}
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(1, c).value
        if v is not None:
            m[str(v).strip()] = c
    return m


def apply_dropdowns_to_worksheet(
    ws: Worksheet,
    *,
    l1_col_name: str = "一级科目",
    l2_col_name: str = "二级科目",
    max_row: Optional[int] = None,
) -> bool:
    """
    给单张工作表加下拉。依赖工作簿已写好 _分类选项 与命名区域。
    返回是否成功加上。
    """
    headers = _header_map(ws)
    c1 = headers.get(l1_col_name)
    c2 = headers.get(l2_col_name)
    if not c1 and not c2:
        return False

    last = max_row or ws.max_row or 2
    if last < 2:
        last = 2  # 预留空行也可用下拉

    # 多留一些空行方便增补
    last = max(last, 2) + 50

    cols_to_clear = [c for c in (c1, c2) if c]
    _clear_old_category_validations(ws, cols_to_clear)

    if c1:
        letter = get_column_letter(c1)
        dv1 = DataValidation(
            type="list",
            formula1="=PAY_L1_LIST",
            allow_blank=True,
            showDropDown=False,  # False = 显示下拉箭头（openpyxl 语义反直觉）
            showErrorMessage=False,  # 允许手工输入未在列表中的值
            showInputMessage=True,
            promptTitle="一级科目",
            prompt="请从列表选择一级分类",
        )
        dv1.add(f"{letter}2:{letter}{last}")
        ws.add_data_validation(dv1)

    if c2 and c1:
        l1_letter = get_column_letter(c1)
        l2_letter = get_column_letter(c2)
        # 二级 = INDIRECT("PAYL2_" & VLOOKUP(一级, MAP, 2, 0))
        # 相对引用：首行写 $L1_letter2，Excel 会对区域内各行调整行号
        formula = (
            f'INDIRECT("PAYL2_"&IFERROR(VLOOKUP({l1_letter}2,PAY_L1_MAP,2,FALSE),""))'
        )
        dv2 = DataValidation(
            type="list",
            formula1=f"={formula}",
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=False,
            showInputMessage=True,
            promptTitle="二级科目",
            prompt="先选一级，再选对应二级（列表随一级联动）",
        )
        dv2.add(f"{l2_letter}2:{l2_letter}{last}")
        ws.add_data_validation(dv2)
    elif c2:
        # 无一级列时：全部二级扁平列表（写在选项表 D 列备用——此处用对照列）
        # 退化为一级表所有二级的并集命名区域；简化：不联动，提示用户
        pass

    return True


def apply_category_dropdowns(
    wb: Workbook,
    data_dir: Optional[Path] = None,
    sheet_names: Optional[Sequence[str]] = None,
) -> int:
    """
    给工作簿内目标 sheet 加一二级下拉。
    返回成功处理的 sheet 数。
    """
    l1_names, l1_id_to_l2, l1_name_to_id, _tax = load_taxonomy_options(data_dir)
    if not l1_names:
        return 0

    _write_options_sheet(wb, l1_names, l1_id_to_l2, l1_name_to_id)

    targets = list(sheet_names) if sheet_names else list(DETAIL_SHEETS)
    # 也处理其它含一级/二级表头的 sheet
    applied = 0
    seen = set()
    for name in list(targets) + list(wb.sheetnames):
        if name in seen or name == OPTIONS_SHEET:
            continue
        if name not in wb.sheetnames:
            continue
        seen.add(name)
        ws = wb[name]
        headers = _header_map(ws)
        if "一级科目" not in headers and "二级科目" not in headers:
            continue
        if apply_dropdowns_to_worksheet(ws):
            applied += 1
    return applied


def apply_category_dropdowns_to_file(
    path: Path,
    data_dir: Optional[Path] = None,
    sheet_names: Optional[Sequence[str]] = None,
) -> int:
    path = Path(path)
    wb = load_workbook(path)
    n = apply_category_dropdowns(wb, data_dir=data_dir, sheet_names=sheet_names)
    wb.save(path)
    return n


def apply_category_dropdowns_to_std_dir(
    std_dir: Path,
    data_dir: Optional[Path] = None,
) -> Dict[str, int]:
    """批量给 B_std 下所有 xlsx 加下拉。"""
    std_dir = Path(std_dir)
    stats = {"files": 0, "sheets": 0, "errors": 0}
    for f in sorted(std_dir.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        try:
            n = apply_category_dropdowns_to_file(f, data_dir=data_dir)
            stats["files"] += 1
            stats["sheets"] += n
            print(f"  ✅ 下拉 {f.name} ({n} sheets)")
        except Exception as e:
            stats["errors"] += 1
            print(f"  ⚠️ 失败 {f.name}: {e}")
    return stats
