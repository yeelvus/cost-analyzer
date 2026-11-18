#!/usr/bin/env python3
"""
步骤2：扫描 raw 各月台账，按项目对比上月，剔除未更新文件；
每个项目只保留 1 份最新有效文件 → EXTRACT_DIR
（B_采购部采购台账_extract）。

含税金额若是公式，先用 Excel 缓存，缓存没有则按本行公式算出数值。
拷贝到提取目录时把这些金额写成数字，避免只剩公式、汇总读成空。
含税总金额只加总每个项目选中的最新一份台账。
月度新增是相对更早台账第一次出现的采购行；每个项目的第一份台账只作基线，不计入新增。

判定优先级：
  文件哈希相同 / 内容哈希相同 → 复用上月（不体现）
  金额+行数一致 → 建议复核后复用
  仅金额一致 → 建议人工复核
  金额变化 → 保留

金额变化写出总金额增加量（本月含税合计减上月含税合计），
并与新增行的含税金额对照，标明是否一致。
各文件汇总列出每一个 Excel 的含税总金额。
"""
from __future__ import annotations

import ast
import hashlib
import re
import shutil
import stat
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula
import sys

_PKG_ROOT = Path(__file__).resolve().parents[1]
_CODE_ROOT = Path(__file__).resolve().parents[2]
for _p in (str(_PKG_ROOT), str(_CODE_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from _采购路径 import (
    EXTRACT_DIR,
    RAW_DIR,
    extract_month_iso,
    extract_project_name,
    is_excel,
    month_sort_key,
)

SKIP_DECISION = "复用上月，当前月不体现"
ANALYSIS_DIR_NAME = "去重分析输出"
HEADER_ALIASES = {
    "材料归类": ["材料归类"],
    "采购部门": ["采购部门"],
    "物品名称": ["物品名称", "采购物资名称", "物资名称"],
    "规格": ["规格", "规格型号", "型号", "规格型号/参数"],
    "单位": ["标准单位", "单位"],
    "数量": ["数量"],
    "单价": ["单价", "不含税单价", "含税单价"],
    "不含税金额": ["不含税金额", "金额"],
    "7%增值税": ["7%增值税", "税额", "增值税", "税金"],
    "含税金额": ["含税金额", "总金额", "价税合计", "含税总金额"],
    "货币种类": ["货币种类", "币种"],
    "汇率": ["汇率"],
    "签订日期": ["签订日期", "日期"],
    "供货商名称": ["供货商名称", "供货单位", "供应商", "供应商名称"],
    "备注": ["备注"],
}
DETAIL_COLUMNS = [
    "材料归类",
    "采购部门",
    "物品名称",
    "规格",
    "单位",
    "数量",
    "单价",
    "不含税金额",
    "7%增值税",
    "含税金额",
    "货币种类",
    "汇率",
    "签订日期",
    "供货商名称",
    "备注",
]
MONEY_COLUMNS = {"数量", "单价", "不含税金额", "7%增值税", "含税金额", "汇率"}
KEY_COLUMNS = ["物品名称", "规格", "单位", "数量", "单价", "货币种类", "签订日期", "供货商名称"]
TOTAL_NAMES = {"合计", "总计", "小计", "合 计"}
_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)%")
_ROW_RE = re.compile(r"^ROW\(\)-(\d+)$")
_SUM_RANGE_RE = re.compile(r"SUM\(([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)\)")


def _norm(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _as_formula(value: object) -> str | None:
    if isinstance(value, ArrayFormula):
        text = value.text or ""
        return text if text.startswith("=") else f"={text}"
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def _literal_number(value: object) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text or text.startswith("="):
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def _safe_arith(expr: str) -> float:
    tree = ast.parse(expr, mode="eval")

    def walk(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            number = walk(node.operand)
            return number if isinstance(node.op, ast.UAdd) else -number
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            left = walk(node.left)
            right = walk(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if right == 0:
                raise ZeroDivisionError
            return left / right
        raise ValueError("unsupported")

    return walk(tree)


def _has_column_total(formula_values: dict[int, object]) -> bool:
    """跨行 SUM 是合计行。同一行的 SUM(I84:J84) 是含税金额，要保留。"""
    for value in formula_values.values():
        formula = _as_formula(value)
        if not formula:
            continue
        expr = formula.upper().replace(" ", "")
        for match in _SUM_RANGE_RE.finditer(expr):
            if int(match.group(2)) != int(match.group(4)):
                return True
    return False


def _formula_refs_text(formula: str, row_number: int, values: dict[int, object]) -> bool:
    expr = formula[1:].replace(" ", "").upper()
    for match in _REF_RE.finditer(expr):
        if int(match.group(2)) != row_number:
            continue
        value = values.get(column_index_from_string(match.group(1)))
        if isinstance(value, str) and not value.startswith("="):
            return True
    return False


def _eval_formula(
    formula: str,
    row_number: int,
    numbers: dict[int, float],
    raw_values: dict[int, object] | None = None,
) -> float | None:
    expr = formula[1:].replace(" ", "").upper()
    row_match = _ROW_RE.fullmatch(expr)
    if row_match:
        return float(row_number - int(row_match.group(1)))

    def expand_sum(match: re.Match[str]) -> str:
        left = column_index_from_string(match.group(1))
        right = column_index_from_string(match.group(3))
        if int(match.group(2)) != row_number or int(match.group(4)) != row_number or left > right:
            raise ValueError("column sum")
        parts: list[str] = []
        for column in range(left, right + 1):
            if column in numbers:
                parts.append(f"({numbers[column]!r})")
                continue
            raw = None if raw_values is None else raw_values.get(column)
            if _as_formula(raw):
                raise ValueError("sum dep")
            parts.append("(0)")
        return "(" + "+".join(parts) + ")"

    def replace_ref(match: re.Match[str]) -> str:
        column = column_index_from_string(match.group(1))
        if int(match.group(2)) != row_number or column not in numbers:
            raise ValueError("missing ref")
        return f"({numbers[column]!r})"

    try:
        expr = _SUM_RANGE_RE.sub(expand_sum, expr)
        if "SUM(" in expr:
            return None
        expr = _REF_RE.sub(replace_ref, expr)
        expr = _PCT_RE.sub(r"(\1/100)", expr)
        return _safe_arith(expr)
    except (ValueError, SyntaxError, ZeroDivisionError, KeyError):
        return None


def _display(value: object, column: str) -> object:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    number = _literal_number(value)
    if number is not None and column in MONEY_COLUMNS:
        return round(number, 2 if column != "汇率" else 6)
    if number is not None:
        return round(number, 6)
    return _norm(value)


def _find_header(ws) -> tuple[int, dict[str, int]] | None:
    max_col = min(ws.max_column or 1, 40)
    for row_idx in range(1, min(ws.max_row or 1, 12) + 1):
        labels = []
        for col in range(1, max_col + 1):
            labels.append(_norm(ws.cell(row_idx, col).value).replace(" ", ""))
        if not any(label in {"序号", "编号", "NO."} for label in labels):
            continue
        found: dict[str, int] = {}
        for name, aliases in HEADER_ALIASES.items():
            for alias in aliases:
                if alias in labels:
                    found[name] = labels.index(alias) + 1
                    break
        if "物品名称" in found and ("含税金额" in found or "不含税金额" in found):
            return row_idx, found
    return None


def _resolve_row(
    formula_values: dict[int, object],
    cached_values: dict[int, object],
    row_number: int,
) -> tuple[dict[int, object], list[tuple[int, float]]]:
    """返回本行最终值，以及需要写回工作簿的 (列, 数值)。"""
    resolved: dict[int, object] = {}
    numbers: dict[int, float] = {}
    pending: dict[int, str] = {}
    for col, raw in formula_values.items():
        formula = _as_formula(raw)
        cached = _literal_number(cached_values.get(col))
        literal = _literal_number(raw)
        if formula is None:
            resolved[col] = raw
            if literal is not None:
                numbers[col] = literal
            continue
        if cached is not None:
            resolved[col] = cached
            numbers[col] = cached
            pending[col] = formula
            continue
        pending[col] = formula
        resolved[col] = raw

    for _ in range(6):
        progress = False
        for col, formula in list(pending.items()):
            if col in numbers:
                continue
            number = _eval_formula(formula, row_number, numbers, formula_values)
            if number is None:
                continue
            numbers[col] = number
            resolved[col] = number
            progress = True
        if not progress:
            break
    # 算出的数和已有缓存都写成数字，避免另存后缓存丢失。
    patches = [(col, numbers[col]) for col in pending if col in numbers]
    return resolved, patches


def read_ledger(path: Path) -> dict:
    """读出台账数值。含税金额优先用缓存，否则按本行公式计算。"""
    formula_wb = load_workbook(path, data_only=False)
    value_wb = load_workbook(path, data_only=True)
    amount_sum = 0.0
    unresolved = 0
    repaired = 0
    data_rows: list[dict] = []
    patches: list[tuple[str, int, int, float]] = []
    hash_acc = hashlib.sha256()
    try:
        for sheet in formula_wb.sheetnames:
            if sheet not in value_wb.sheetnames:
                continue
            fws = formula_wb[sheet]
            vws = value_wb[sheet]
            header = _find_header(fws)
            if header is None:
                continue
            header_row, columns = header
            max_col = max(columns.values())
            sheet_lines: list[str] = []
            for row_idx in range(header_row + 1, (fws.max_row or header_row) + 1):
                formula_values = {col: fws.cell(row_idx, col).value for col in range(1, max_col + 1)}
                cached_values = {col: vws.cell(row_idx, col).value for col in range(1, max_col + 1)}
                if _has_column_total(formula_values):
                    continue
                resolved, row_patches = _resolve_row(formula_values, cached_values, row_idx)
                tax_col = columns.get("含税金额")
                net_col = columns.get("不含税金额")
                vat_col = columns.get("7%增值税")
                if tax_col and _literal_number(resolved.get(tax_col)) is None:
                    tax_formula = _as_formula(formula_values.get(tax_col))
                    net = _literal_number(resolved.get(net_col)) if net_col else None
                    vat = _literal_number(resolved.get(vat_col)) if vat_col else None
                    # 公式把供货商之类的文字加进含税金额时，Excel 得到 #VALUE!。
                    # 不含税金额和增值税已经算出时，按这两项之和计入。
                    if (
                        tax_formula
                        and net is not None
                        and vat is not None
                        and _formula_refs_text(tax_formula, row_idx, formula_values)
                    ):
                        resolved[tax_col] = net + vat
                        row_patches = [item for item in row_patches if item[0] != tax_col]
                        row_patches.append((tax_col, net + vat))
                        repaired += 1
                record = {name: _display(resolved.get(col, ""), name) for name, col in columns.items()}
                item_name = _norm(record.get("物品名称"))
                if item_name in TOTAL_NAMES or "合计" in item_name:
                    continue
                if not any(_norm(record.get(name)) for name in DETAIL_COLUMNS if name in record):
                    continue
                tax = record.get("含税金额", "")
                if isinstance(tax, (int, float)):
                    amount_sum += float(tax)
                elif columns.get("含税金额") and _as_formula(formula_values.get(columns["含税金额"])):
                    unresolved += 1
                for col, number in row_patches:
                    patches.append((sheet, row_idx, col, round(number, 6)))
                data_rows.append(record)
                sheet_lines.append("\t".join(_norm(record.get(name, "")) for name in DETAIL_COLUMNS))
            if sheet_lines:
                payload = sheet + "\n" + "\n".join(sheet_lines)
                hash_acc.update(payload.encode())
    finally:
        formula_wb.close()
        value_wb.close()
    return {
        "含税金额合计": round(amount_sum, 2),
        "行数": len(data_rows),
        "内容哈希": hash_acc.hexdigest(),
        "未解析含税行": unresolved,
        "含税回补行": repaired,
        "明细": data_rows,
        "回写": patches,
    }


def _row_key(record: dict) -> tuple:
    parts = []
    for column in KEY_COLUMNS:
        value = record.get(column, "")
        if column in MONEY_COLUMNS and isinstance(value, (int, float)):
            parts.append(round(float(value), 2))
        else:
            parts.append(_norm(value))
    return tuple(parts)


def scan_raw(raw_root: Path) -> tuple[pd.DataFrame, dict[str, dict]]:
    table_rows = []
    ledgers: dict[str, dict] = {}
    files = [path for path in sorted(raw_root.rglob("*")) if is_excel(path)]
    for index, path in enumerate(files, start=1):
        month = extract_month_iso(str(path))
        if not month:
            continue
        ledger = read_ledger(path)
        ledgers[str(path)] = ledger
        table_rows.append(
            {
                "月份": month,
                "项目名称": extract_project_name(path.name),
                "文件名": path.name,
                "文件路径": str(path),
                "含税金额合计": ledger["含税金额合计"],
                "行数": ledger["行数"],
                "未解析含税行": ledger["未解析含税行"],
                "文件哈希": _sha256_file(path),
                "内容哈希": ledger["内容哈希"],
            }
        )
        if ledger["行数"] == 0:
            print(f"  未识别到采购行: {path}")
        if index % 40 == 0:
            print(f"  已读取 {index}/{len(files)}")
    return pd.DataFrame(table_rows), ledgers


def _tax_of(records: list[dict], indexes: list[int]) -> float:
    total = 0.0
    for index in indexes:
        amount = _money(records[index].get("含税金额"))
        if amount is not None:
            total += amount
    return round(total, 2)


def _new_and_removed(current: list[dict], previous: list[dict]) -> tuple[int, float, int, float]:
    """相对上一份台账，多出来的行和少掉的行。同一键按出现次数比较。"""
    current_groups: dict[tuple, list[int]] = defaultdict(list)
    previous_groups: dict[tuple, list[int]] = defaultdict(list)
    for index, record in enumerate(current):
        current_groups[_row_key(record)].append(index)
    for index, record in enumerate(previous):
        previous_groups[_row_key(record)].append(index)
    new_indexes: list[int] = []
    removed_indexes: list[int] = []
    for key, indexes in current_groups.items():
        extra = len(indexes) - len(previous_groups.get(key, []))
        if extra > 0:
            new_indexes.extend(indexes[-extra:])
    for key, indexes in previous_groups.items():
        extra = len(indexes) - len(current_groups.get(key, []))
        if extra > 0:
            removed_indexes.extend(indexes[-extra:])
    return (
        len(new_indexes),
        _tax_of(current, new_indexes),
        len(removed_indexes),
        _tax_of(previous, removed_indexes),
    )


def _increase_fields(row: pd.Series, prev: pd.Series | None, ledgers: dict[str, dict]) -> dict:
    """总金额增加量 = 本月含税合计 - 上月含税合计。新增行金额与它分开对照。"""
    current_total = round(float(row["含税金额合计"]), 2)
    current_rows = ledgers.get(str(row["文件路径"]), {}).get("明细", [])
    if prev is None:
        new_amount = _tax_of(current_rows, list(range(len(current_rows))))
        gap = round(current_total - new_amount, 2)
        return {
            "上月含税总金额": None,
            "总金额增加量": current_total,
            "新增行数": len(current_rows),
            "新增行含税金额": new_amount,
            "减少行数": 0,
            "减少行含税金额": 0.0,
            "相差": gap,
            "金额增加是否一致": "是" if abs(gap) <= 0.02 else "否",
        }
    previous_total = round(float(prev["含税金额合计"]), 2)
    delta = round(current_total - previous_total, 2)
    previous_rows = ledgers.get(str(prev["文件路径"]), {}).get("明细", [])
    new_count, new_amount, removed_count, removed_amount = _new_and_removed(current_rows, previous_rows)
    gap = round(delta - new_amount, 2)
    return {
        "上月含税总金额": previous_total,
        "总金额增加量": delta,
        "新增行数": new_count,
        "新增行含税金额": new_amount,
        "减少行数": removed_count,
        "减少行含税金额": removed_amount,
        "相差": gap,
        "金额增加是否一致": "是" if abs(gap) <= 0.02 else "否",
    }


def _conclusion_amount(fields: dict) -> str:
    text = (
        f"总金额增加 {fields['总金额增加量']:,.2f}；"
        f"新增{fields['新增行数']}行金额 {fields['新增行含税金额']:,.2f}；"
        f"{'一致' if fields['金额增加是否一致'] == '是' else '不一致'}"
    )
    if fields["金额增加是否一致"] != "是":
        text += f"，相差 {fields['相差']:,.2f}"
        if fields["减少行数"]:
            text += f"；减少{fields['减少行数']}行金额 {fields['减少行含税金额']:,.2f}"
    return text


def decide(df: pd.DataFrame, ledgers: dict[str, dict] | None = None) -> pd.DataFrame:
    ledgers = ledgers or {}
    extra = [
        "上月含税总金额", "总金额增加量", "新增行数", "新增行含税金额",
        "减少行数", "减少行含税金额", "相差", "金额增加是否一致",
        "对比结论", "处理建议", "是否在最新口径体现",
    ]
    if df.empty:
        return pd.DataFrame(columns=list(df.columns) + extra)

    out = []
    for _, g in df.groupby("项目名称", sort=False):
        g = g.sort_values("月份", key=lambda s: s.map(month_sort_key))
        prev = None
        for _, row in g.iterrows():
            fields = _increase_fields(row, prev, ledgers)
            reason, decision = "首月保留", "保留"
            if prev is not None:
                amt_diff = abs(fields["总金额增加量"])
                row_diff = int(row["行数"]) - int(prev["行数"])
                if row["文件哈希"] == prev["文件哈希"]:
                    reason, decision = "与上月文件字节完全一致", SKIP_DECISION
                elif row["内容哈希"] == prev["内容哈希"]:
                    reason, decision = "与上月内容值一致（忽略文件元数据）", SKIP_DECISION
                elif amt_diff <= 0.01 and row_diff == 0:
                    reason, decision = "含税金额合计一致且行数一致（疑似未更新）", "建议复核后复用上月"
                elif amt_diff <= 0.01:
                    reason, decision = "含税金额合计一致但行数变化", "建议人工复核"
                else:
                    reason, decision = _conclusion_amount(fields), "保留"
            out.append(
                {
                    **row.to_dict(),
                    **fields,
                    "对比结论": reason,
                    "处理建议": decision,
                    "是否在最新口径体现": "否" if decision == SKIP_DECISION else "是",
                }
            )
            prev = row
    return pd.DataFrame(out)


def _select_latest(decision_df: pd.DataFrame) -> pd.DataFrame:
    if decision_df.empty:
        return decision_df
    cand = decision_df[decision_df["处理建议"] != SKIP_DECISION].copy()
    if cand.empty:
        return cand
    priority = {"保留": 3, "建议复核后复用上月": 2, "建议人工复核": 1}
    cand["_mk"] = cand["月份"].map(month_sort_key)
    cand["_pri"] = cand["处理建议"].map(priority).fillna(0)
    cand = cand.sort_values(["_mk", "_pri"], ascending=[False, False])
    selected = cand.groupby("项目名称", sort=False, as_index=False).head(1)
    return selected.drop(columns=["_mk", "_pri"])


def _money(value: object) -> float | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def build_increment(decision_df: pd.DataFrame, ledgers: dict[str, dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """同一项目按月份先后比较。第一份台账只记入已见键，不记新增。

    之后某一行的业务键在更早台账里没出现过，才记为本月新增。金额用已经算出的数值。
    """
    details: list[dict] = []
    if decision_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    for _, group in decision_df.groupby("项目名称", sort=False):
        group = group.sort_values("月份", key=lambda s: s.map(month_sort_key))
        seen: Counter = Counter()
        seeded = False
        for _, row in group.iterrows():
            ledger = ledgers.get(str(row["文件路径"]), {})
            records = ledger.get("明细", [])
            grouped: dict[tuple, list[int]] = defaultdict(list)
            for index, record in enumerate(records):
                grouped[_row_key(record)].append(index)
            current_counts: Counter = Counter({key: len(indexes) for key, indexes in grouped.items()})
            if not seeded:
                seen.update(current_counts)
                seeded = True
                continue
            chosen: list[int] = []
            for key, indexes in grouped.items():
                extra = len(indexes) - seen[key]
                if extra > 0:
                    chosen.extend(indexes[-extra:])
            for index in sorted(chosen):
                record = records[index]
                item = {
                    "台账月份": row["月份"],
                    "项目名称": row["项目名称"],
                    "来源文件": row["文件名"],
                }
                for column in DETAIL_COLUMNS:
                    value = record.get(column, "")
                    item[column] = "" if isinstance(value, str) and value.startswith("=") else value
                details.append(item)
            for key, count in current_counts.items():
                if count > seen[key]:
                    seen[key] = count
    detail_df = pd.DataFrame(details)
    if detail_df.empty:
        return detail_df, pd.DataFrame(columns=["台账月份", "项目名称", "新增行数", "新增含税金额"])
    work = detail_df.copy()
    work["_金额"] = work["含税金额"].map(_money)
    summary = (
        work.groupby(["台账月份", "项目名称"], as_index=False)
        .agg(新增行数=("物品名称", "size"), 新增含税金额=("_金额", "sum"))
        .sort_values(["台账月份", "项目名称"])
    )
    summary["新增含税金额"] = summary["新增含税金额"].round(2)
    total = pd.DataFrame(
        [
            {
                "台账月份": "合计",
                "项目名称": "",
                "新增行数": int(summary["新增行数"].sum()),
                "新增含税金额": round(float(summary["新增含税金额"].sum()), 2),
            }
        ]
    )
    return detail_df, pd.concat([summary, total], ignore_index=True)


def project_amount_summary(selected: pd.DataFrame) -> pd.DataFrame:
    if selected.empty:
        return pd.DataFrame(columns=["项目名称", "选中月份", "行数", "含税总金额", "未解析含税行", "文件名"])
    table = selected[["项目名称", "月份", "行数", "含税金额合计", "未解析含税行", "文件名"]].copy()
    table = table.rename(columns={"月份": "选中月份", "含税金额合计": "含税总金额"})
    table["含税总金额"] = table["含税总金额"].round(2)
    total = pd.DataFrame(
        [
            {
                "项目名称": "合计",
                "选中月份": "",
                "行数": int(table["行数"].sum()),
                "含税总金额": round(float(table["含税总金额"].sum()), 2),
                "未解析含税行": int(table["未解析含税行"].sum()),
                "文件名": "",
            }
        ]
    )
    return pd.concat([table, total], ignore_index=True)


def _write_resolved_amounts(src: Path, dst: Path, patches: list[tuple[str, int, int, float]]) -> None:
    """把公式金额写成数字。只改提取目录里的副本，不动原始台账。"""
    shutil.copy2(src, dst)
    dst.chmod(dst.stat().st_mode | stat.S_IWUSR)
    if not patches:
        return
    wb = load_workbook(dst)
    try:
        for sheet, row_idx, col, number in patches:
            if sheet not in wb.sheetnames:
                continue
            wb[sheet].cell(row_idx, col).value = number
        wb.save(dst)
    finally:
        wb.close()


def export_latest_per_project(
    decision_df: pd.DataFrame,
    ledgers: dict[str, dict],
    out_dir: Path,
) -> tuple[int, int]:
    """每项目保留最新非明确重复文件；公式金额写成数字。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    analysis = out_dir / ANALYSIS_DIR_NAME
    analysis.mkdir(exist_ok=True)

    for item in out_dir.iterdir():
        if item.name == ANALYSIS_DIR_NAME:
            continue
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()

    selected = _select_latest(decision_df)
    if selected.empty:
        return 0, 0

    copied, conflicts = [], []
    used: set[str] = set()
    for _, row in selected.iterrows():
        src = Path(str(row["文件路径"]))
        if not src.is_file():
            continue
        dst = out_dir / src.name
        if src.name in used or dst.exists():
            conflicts.append({"项目名称": row["项目名称"], "文件名": src.name, "路径": str(src)})
            continue
        patches = ledgers.get(str(src), {}).get("回写", [])
        _write_resolved_amounts(src, dst, patches)
        used.add(src.name)
        copied.append(
            {
                "项目名称": row["项目名称"],
                "选中月份": row["月份"],
                "处理建议": row["处理建议"],
                "对比结论": row["对比结论"],
                "含税总金额": row["含税金额合计"],
                "未解析含税行": row["未解析含税行"],
                "文件名": src.name,
            }
        )

    pd.DataFrame(copied).to_excel(analysis / "A_拷贝清单.xlsx", index=False)
    if conflicts:
        pd.DataFrame(conflicts).to_excel(analysis / "同名冲突清单.xlsx", index=False)
    return len(copied), len(conflicts)


def run(raw_root: Path = RAW_DIR, extract_dir: Path = EXTRACT_DIR) -> None:
    print(f"源目录: {raw_root}")
    print(f"输出目录: {extract_dir}")

    df, ledgers = scan_raw(raw_root)
    if df.empty:
        print("未扫描到 Excel，结束")
        decision = pd.DataFrame()
        summary = pd.DataFrame()
        increment_detail = pd.DataFrame()
        increment_summary = pd.DataFrame()
        amount_summary = pd.DataFrame()
        file_amount = pd.DataFrame()
        repaired = 0
    else:
        df = df.sort_values(["项目名称", "月份"]).reset_index(drop=True)
        decision = decide(df, ledgers)
        summary = (
            decision.groupby(["月份", "处理建议"], as_index=False)
            .agg(
                文件数=("文件名", "count"),
                项目数=("项目名称", "nunique"),
                文件含税合计=("含税金额合计", "sum"),
            )
            .sort_values(["月份", "处理建议"])
        )
        summary["文件含税合计"] = summary["文件含税合计"].round(2)
        file_amount = decision[
            [
                "月份", "项目名称", "文件名", "行数", "含税金额合计",
                "上月含税总金额", "总金额增加量", "新增行数", "新增行含税金额",
                "减少行数", "减少行含税金额", "相差", "金额增加是否一致",
                "对比结论", "处理建议",
            ]
        ].rename(columns={"含税金额合计": "含税总金额"})
        increment_detail, increment_summary = build_increment(decision, ledgers)
        selected = _select_latest(decision)
        amount_summary = project_amount_summary(selected)
        repaired = 0
        if not selected.empty:
            repaired = sum(int(ledgers.get(str(path), {}).get("含税回补行", 0)) for path in selected["文件路径"])

    analysis = extract_dir / ANALYSIS_DIR_NAME
    analysis.mkdir(parents=True, exist_ok=True)
    report = analysis / "采购台账_月度去重分析.xlsx"
    with pd.ExcelWriter(report, engine="openpyxl") as writer:
        decision.to_excel(writer, index=False, sheet_name="逐文件判定")
        file_amount.to_excel(writer, index=False, sheet_name="各文件汇总")
        summary.to_excel(writer, index=False, sheet_name="月份汇总")
        amount_summary.to_excel(writer, index=False, sheet_name="项目含税总金额")
        increment_summary.to_excel(writer, index=False, sheet_name="月度新增汇总")
        increment_detail.to_excel(writer, index=False, sheet_name="月度新增明细")

    n_copy, n_conflict = export_latest_per_project(decision, ledgers, extract_dir)
    total = 0.0
    if not amount_summary.empty:
        total = float(amount_summary.loc[amount_summary["项目名称"] == "合计", "含税总金额"].iloc[0])
    unresolved = 0
    if not amount_summary.empty and "未解析含税行" in amount_summary.columns:
        unresolved = int(amount_summary.loc[amount_summary["项目名称"] == "合计", "未解析含税行"].iloc[0])
    print(f"分析报告: {report}")
    print(f"判定行数: {len(decision)} | 拷贝: {n_copy} | 冲突: {n_conflict}")
    print(f"最新台账含税总金额: {total:,.2f}")
    print(f"选中台账里未能算出含税金额的行: {unresolved}")
    if repaired:
        print(f"含税公式引用了文字、已按不含税金额+增值税计入的行: {repaired}")
    print(f"提取目录: {extract_dir}")


def main() -> None:
    run()


if __name__ == "__main__":
    main()
