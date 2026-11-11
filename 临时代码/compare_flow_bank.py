# -*- coding: utf-8 -*-
"""把「最新银行账」和「流水」按项目、按日期对照，并合成一套完整付款数据。

同一天里先配实付金额相同的单行。对不上的，再把若干行配成一组：
一笔银行账对多笔流水、多笔银行账对一笔流水，或两边各若干笔但合计相同。
金额用「分」比较，允许 1 分误差。银行账取实付金额，流水也取实付金额
（预扣税在流水的预扣税额、付款金额里，不拿来改配对）。

左右对照用「匹配组」和「对应关系」标出哪些行是同一组。
完整台账以流水为底（保留预扣税、增值税、不含税、付款金额），银行账只补实付；
只有银行账、没有流水的行追加在后面。

未识别编码的子公司银行账和 ZGS 子公司流水是同一套账、分在两个项目里。
并入时只认同一天、同一实付，并且摘要去掉「子公司」「支付」等前缀后对得上。
"""

from __future__ import annotations

import sys
from collections import defaultdict
from itertools import combinations
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BANK_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/B_财务付款_std/最新银行账"
)
FLOW_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/B_财务付款_std/流水"
)
OUT_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/02_Std/12_财务付款/流水与银行账按日对照"
)
SUMMARY_NAME = "00_项目汇总.xlsx"
COMPLETE_NAME = "00_完整台账.xlsx"

BANK_FIELDS = [
    ("项目编码", "银行账_项目编码"),
    ("摘要", "银行账_摘要"),
    ("实付金额", "银行账_实付金额"),
    ("日期", "银行账_日期"),
]
FLOW_FIELDS = [
    ("摘要", "流水_摘要"),
    ("应付金额", "流水_应付金额"),
    ("预扣税额", "流水_预扣税额"),
    ("实付金额", "流水_实付金额"),
    ("增值税额", "流水_增值税额"),
    ("不含税额", "流水_不含税额"),
    ("付款金额", "流水_付款金额"),
    ("日期", "流水_日期"),
]
TAX_FIELDS = ["应付金额", "预扣税额", "增值税额", "不含税额", "付款金额"]
AMOUNT_HEADERS = {
    "银行账_实付金额",
    "实付差值",
    "流水_应付金额",
    "流水_预扣税额",
    "流水_实付金额",
    "流水_增值税额",
    "流水_不含税额",
    "流水_付款金额",
    "应付金额",
    "预扣税额",
    "实付金额",
    "增值税额",
    "不含税额",
    "付款金额",
    "银行账实付金额",
    "流水实付金额",
    "本组银行实付",
    "本组流水实付",
    "银行账实付合计",
    "流水实付合计",
    "实付差值合计",
    "完整实付合计",
}
DATE_HEADERS = {"银行账_日期", "流水_日期", "对照日期", "日期"}
ONE_TO_ONE = "一对一"
BANK_TO_MANY = "一笔银行对多笔流水"
MANY_TO_BANK = "多笔银行对一笔流水"
COMBO = "组合对应"
BANK_ONLY = "仅银行账"
FLOW_ONLY = "仅流水"
MATCHED_KINDS = {ONE_TO_ONE, BANK_TO_MANY, MANY_TO_BANK, COMBO}
CENT = 1


def project_key(path: Path, folder_name: str) -> str:
    stem = path.stem
    suffix = f"_{folder_name}"
    if stem.endswith(suffix):
        return stem[: -len(suffix)]
    return stem


def list_books(folder: Path) -> dict[str, Path]:
    books: dict[str, Path] = {}
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$") or path.name.startswith("."):
            continue
        books[project_key(path, folder.name)] = path
    return books


def read_detail(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    sheet = "明细数据" if "明细数据" in xl.sheet_names else xl.sheet_names[0]
    frame = pd.read_excel(path, sheet_name=sheet)
    if "日期" in frame.columns:
        frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce")
    return frame


def paid_amount(row: pd.Series | None) -> float:
    if row is None:
        return 0.0
    for column in ("实付金额", "付款金额"):
        if column not in row.index:
            continue
        value = row[column]
        if pd.notna(value):
            return float(value)
    return 0.0


def cents_of(row: pd.Series | None) -> int:
    return int(round(paid_amount(row) * 100))


def day_key(value) -> str:
    if value is None or pd.isna(value):
        return ""
    stamp = pd.Timestamp(value)
    if pd.isna(stamp):
        return ""
    return stamp.strftime("%Y-%m-%d")


def _text(row: pd.Series | None, *columns: str) -> str:
    if row is None:
        return ""
    for column in columns:
        if column not in row.index:
            continue
        value = row[column]
        if value is None or pd.isna(value):
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def summary_score(left: pd.Series, right: pd.Series) -> float:
    """摘要越像，一对一里相同金额越优先配在一起。"""
    a = "".join(_text(left, "摘要", "业务摘要").split())
    b = "".join(_text(right, "摘要", "业务摘要").split())
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 1.0
    left_chars, right_chars = set(a), set(b)
    return len(left_chars & right_chars) / len(left_chars | right_chars)


def _max_subset_size(count: int) -> int:
    if count <= 12:
        return 6
    if count <= 18:
        return 5
    if count <= 28:
        return 4
    return 3


def _allow_sparse(count: int, size: int) -> bool:
    if size <= 2:
        return count <= 80
    if size == 3:
        return count <= 36
    if size == 4:
        return count <= 22
    return count <= 14


def find_smallest_subsets(indexed: list[tuple[int, int]], target: int, limit: int = 24) -> list[tuple[int, ...]]:
    """在 indexed=(原下标, 金额分) 里找合计等于 target 的最小组合。先看连着的行。"""
    count = len(indexed)
    if count < 2:
        return []
    max_size = _max_subset_size(count)
    for size in range(2, min(max_size, count) + 1):
        found: list[tuple[int, ...]] = []
        for start in range(0, count - size + 1):
            window = indexed[start : start + size]
            if abs(sum(value for _idx, value in window) - target) <= CENT:
                found.append(tuple(idx for idx, _value in window))
                if len(found) >= limit:
                    return found
        if _allow_sparse(count, size):
            for comb in combinations(range(count), size):
                if abs(sum(indexed[pos][1] for pos in comb) - target) <= CENT:
                    found.append(tuple(indexed[pos][0] for pos in comb))
                    if len(found) >= limit:
                        return found
        if found:
            return found
    if abs(sum(value for _idx, value in indexed) - target) <= CENT:
        return [tuple(idx for idx, _value in indexed)]
    return []


def _pick_subset(subsets: list[tuple[int, ...]], anchor: pd.Series, pool: list[pd.Series]) -> tuple[int, ...]:
    def sort_key(choice: tuple[int, ...]) -> tuple:
        score = sum(summary_score(anchor, pool[idx]) for idx in choice)
        ordered = sorted(choice)
        span = ordered[-1] - ordered[0] if ordered else 0
        return (-score, span, ordered)

    return min(subsets, key=sort_key)


SUBSIDIARY_FLOW_KEY = "ZGS_子公司"
_CORE_PREFIXES = ("子公司-", "子公司", "支付", "报销", "发放", "缴纳", "补缴", "代付", "代垫")


def core_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    text = "".join(str(value).split())
    if not text or text.lower() == "nan":
        return ""
    changed = True
    while changed and text:
        changed = False
        for prefix in _CORE_PREFIXES:
            if text.startswith(prefix):
                text = text[len(prefix) :].lstrip("-—_")
                changed = True
                break
    return text


def row_cores(row: pd.Series) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for column in ("业务摘要", "摘要"):
        if column not in row.index:
            continue
        text = core_text(row[column])
        if text and text not in seen:
            seen.add(text)
            found.append(text)
    return found


def text_rank(bank_row: pd.Series, flow_row: pd.Series) -> int:
    """3=摘要相同，2=一方摘要包含另一方。对不上返回 0。"""
    left = row_cores(bank_row)
    right = row_cores(flow_row)
    if not left or not right:
        return 0
    if set(left) & set(right):
        return 3
    for a in left:
        for b in right:
            short, longer = (a, b) if len(a) <= len(b) else (b, a)
            if len(short) >= 4 and short in longer:
                return 2
    return 0


def _lcs_len(left: str, right: str) -> int:
    if not left or not right:
        return 0
    prev = [0] * (len(right) + 1)
    best = 0
    for char in left:
        current = [0]
        for index, other in enumerate(right, start=1):
            if char == other:
                current.append(prev[index - 1] + 1)
                if current[-1] > best:
                    best = current[-1]
            else:
                current.append(0)
        prev = current
    return best


def attach_subsidiary_bank(
    bank_frames: dict[str, pd.DataFrame],
    flow_frames: dict[str, pd.DataFrame],
) -> int:
    """未识别的子公司银行账并入 ZGS 流水项目，便于按同一天同一实付配对。"""
    flow = flow_frames.get(SUBSIDIARY_FLOW_KEY)
    if flow is None or flow.empty:
        return 0
    bank_keys = [key for key in bank_frames if key.startswith("未识别编码_子公司")]
    if not bank_keys:
        return 0

    flow_rows = [row for _, row in flow.iterrows()]
    flow_index: dict[tuple[str, int], list[int]] = defaultdict(list)
    for pos, row in enumerate(flow_rows):
        flow_index[(day_key(row.get("日期")) or "", cents_of(row))].append(pos)

    bank_rows: list[tuple[str, object, pd.Series, str, int]] = []
    bank_index: dict[tuple[str, int], list[int]] = defaultdict(list)
    for key in bank_keys:
        frame = bank_frames[key]
        if frame is None or frame.empty:
            continue
        for idx, row in frame.iterrows():
            day = day_key(row.get("日期")) or ""
            amount = cents_of(row)
            bank_index[(day, amount)].append(len(bank_rows))
            bank_rows.append((key, idx, row, day, amount))

    candidates: list[tuple[int, int, int]] = []
    for slot, (_key, _idx, row, day, amount) in enumerate(bank_rows):
        flow_hits = flow_index.get((day, amount), [])
        bank_hits = bank_index.get((day, amount), [])
        ranked: list[tuple[int, int]] = []
        for flow_pos in flow_hits:
            rank = text_rank(row, flow_rows[flow_pos])
            if rank == 0 and len(flow_hits) == 1 and len(bank_hits) == 1:
                if _lcs_len("".join(row_cores(row)), "".join(row_cores(flow_rows[flow_pos]))) >= 3:
                    rank = 1
            if rank:
                ranked.append((rank, flow_pos))
        if not ranked:
            continue
        best = max(rank for rank, _pos in ranked)
        chosen = [pos for rank, pos in ranked if rank == best]
        if len(chosen) == 1:
            candidates.append((best, slot, chosen[0]))

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    used_bank: set[int] = set()
    used_flow: set[int] = set()
    taken: dict[str, list] = defaultdict(list)
    for _rank, slot, flow_pos in candidates:
        if slot in used_bank or flow_pos in used_flow:
            continue
        used_bank.add(slot)
        used_flow.add(flow_pos)
        key, idx, _row, _day, _amount = bank_rows[slot]
        taken[key].append(idx)
    if not taken:
        return 0

    parts: list[pd.DataFrame] = []
    existing = bank_frames.get(SUBSIDIARY_FLOW_KEY)
    if existing is not None and not existing.empty:
        parts.append(existing)
    moved = 0
    for key, indexes in taken.items():
        frame = bank_frames[key]
        parts.append(frame.loc[indexes])
        bank_frames[key] = frame.drop(index=indexes)
        moved += len(indexes)
    bank_frames[SUBSIDIARY_FLOW_KEY] = pd.concat(parts, ignore_index=True)
    return moved


def match_day(bank_rows: list, flow_rows: list) -> list[tuple[str, list[int], list[int]]]:
    """返回 (对应关系, 银行下标, 流水下标)。下标是当天列表里的位置。"""
    bank_open = set(range(len(bank_rows)))
    flow_open = set(range(len(flow_rows)))
    groups: list[tuple[str, list[int], list[int]]] = []

    def take_pairs() -> None:
        flow_by_amount: dict[int, list[int]] = {}
        for idx in flow_open:
            flow_by_amount.setdefault(cents_of(flow_rows[idx]), []).append(idx)
        bank_by_amount: dict[int, list[int]] = {}
        for idx in bank_open:
            bank_by_amount.setdefault(cents_of(bank_rows[idx]), []).append(idx)

        used_flow: set[int] = set()
        used_bank: set[int] = set()
        for amount, banks in bank_by_amount.items():
            flows = flow_by_amount.get(amount, [])
            if len(banks) == 1 and len(flows) == 1:
                groups.append((ONE_TO_ONE, [banks[0]], [flows[0]]))
                used_bank.add(banks[0])
                used_flow.add(flows[0])
        bank_open.difference_update(used_bank)
        flow_open.difference_update(used_flow)

        for bank_idx in sorted(bank_open):
            amount = cents_of(bank_rows[bank_idx])
            choices = [idx for idx in flow_by_amount.get(amount, []) if idx in flow_open]
            if not choices:
                continue
            flow_idx = max(
                choices,
                key=lambda idx: (summary_score(bank_rows[bank_idx], flow_rows[idx]), -idx),
            )
            groups.append((ONE_TO_ONE, [bank_idx], [flow_idx]))
            flow_open.remove(flow_idx)
            used_bank.add(bank_idx)
        bank_open.difference_update(used_bank)

    def take_splits() -> None:
        while bank_open and flow_open and (len(bank_open) >= 2 or len(flow_open) >= 2):
            candidates: list[tuple] = []
            if len(flow_open) >= 2:
                flow_indexed = [(idx, cents_of(flow_rows[idx])) for idx in sorted(flow_open)]
                for bank_idx in sorted(bank_open):
                    subsets = find_smallest_subsets(flow_indexed, cents_of(bank_rows[bank_idx]))
                    if not subsets:
                        continue
                    chosen = subsets[0] if len(subsets) == 1 else _pick_subset(subsets, bank_rows[bank_idx], flow_rows)
                    candidates.append((len(chosen), len(subsets), -cents_of(bank_rows[bank_idx]), BANK_TO_MANY, bank_idx, chosen))
            if len(bank_open) >= 2:
                bank_indexed = [(idx, cents_of(bank_rows[idx])) for idx in sorted(bank_open)]
                for flow_idx in sorted(flow_open):
                    subsets = find_smallest_subsets(bank_indexed, cents_of(flow_rows[flow_idx]))
                    if not subsets:
                        continue
                    chosen = subsets[0] if len(subsets) == 1 else _pick_subset(subsets, flow_rows[flow_idx], bank_rows)
                    candidates.append((len(chosen), len(subsets), -cents_of(flow_rows[flow_idx]), MANY_TO_BANK, flow_idx, chosen))
            if not candidates:
                return
            uniques = [item for item in candidates if item[1] == 1]
            pool = uniques or [min(candidates)]
            pool.sort()
            moved = False
            claimed_bank: set[int] = set()
            claimed_flow: set[int] = set()
            for _size, _alts, _amount, kind, anchor, subset in pool:
                if kind == BANK_TO_MANY:
                    if anchor in claimed_bank or anchor not in bank_open:
                        continue
                    if any(idx in claimed_flow or idx not in flow_open for idx in subset):
                        continue
                    groups.append((kind, [anchor], list(subset)))
                    claimed_bank.add(anchor)
                    claimed_flow.update(subset)
                else:
                    if anchor in claimed_flow or anchor not in flow_open:
                        continue
                    if any(idx in claimed_bank or idx not in bank_open for idx in subset):
                        continue
                    groups.append((kind, list(subset), [anchor]))
                    claimed_flow.add(anchor)
                    claimed_bank.update(subset)
                moved = True
                if not uniques:
                    break
            if not moved:
                return
            bank_open.difference_update(claimed_bank)
            flow_open.difference_update(claimed_flow)

    def take_combo() -> None:
        while len(bank_open) >= 2 and len(flow_open) >= 2 and len(bank_open) <= 10 and len(flow_open) <= 10:
            bank_indexed = [(idx, cents_of(bank_rows[idx])) for idx in sorted(bank_open)]
            flow_indexed = [(idx, cents_of(flow_rows[idx])) for idx in sorted(flow_open)]
            sums: dict[int, list[tuple[int, ...]]] = {}
            for size in range(2, min(4, len(bank_indexed)) + 1):
                for comb in combinations(bank_indexed, size):
                    total = sum(value for _idx, value in comb)
                    sums.setdefault(total, []).append(tuple(idx for idx, _value in comb))
            best = None
            for size in range(2, min(4, len(flow_indexed)) + 1):
                for comb in combinations(flow_indexed, size):
                    total = sum(value for _idx, value in comb)
                    bank_choices = sums.get(total)
                    if not bank_choices:
                        continue
                    flow_choice = tuple(idx for idx, _value in comb)
                    # 整组两边正好用完的，留给最后一次合计判断，避免拆碎。
                    if len(bank_choices[0]) == len(bank_open) and len(flow_choice) == len(flow_open):
                        continue
                    cand = (len(bank_choices[0]) + len(flow_choice), bank_choices[0], flow_choice)
                    if best is None or cand < best:
                        best = cand
            if best is None:
                return
            _total, bank_choice, flow_choice = best
            groups.append((COMBO, list(bank_choice), list(flow_choice)))
            bank_open.difference_update(bank_choice)
            flow_open.difference_update(flow_choice)

    take_pairs()
    take_splits()
    take_combo()
    if bank_open and flow_open:
        bank_sum = sum(cents_of(bank_rows[idx]) for idx in bank_open)
        flow_sum = sum(cents_of(flow_rows[idx]) for idx in flow_open)
        if abs(bank_sum - flow_sum) <= CENT:
            kind = COMBO
            if len(bank_open) == 1 and len(flow_open) > 1:
                kind = BANK_TO_MANY
            elif len(flow_open) == 1 and len(bank_open) > 1:
                kind = MANY_TO_BANK
            elif len(bank_open) == 1 and len(flow_open) == 1:
                kind = ONE_TO_ONE
            groups.append((kind, sorted(bank_open), sorted(flow_open)))
            bank_open.clear()
            flow_open.clear()

    for idx in sorted(bank_open):
        groups.append((BANK_ONLY, [idx], []))
    for idx in sorted(flow_open):
        groups.append((FLOW_ONLY, [], [idx]))

    def sort_key(group: tuple[str, list[int], list[int]]) -> tuple:
        _kind, banks, flows = group
        if banks:
            return (0, banks[0], flows[0] if flows else 0)
        return (1, flows[0] if flows else 0, 0)

    groups.sort(key=sort_key)
    return groups


def side_values(row: pd.Series | None, fields: list[tuple[str, str]]) -> dict:
    out = {header: None for _source, header in fields}
    if row is None:
        return out
    for source, header in fields:
        if source not in row.index:
            continue
        value = row[source]
        if pd.isna(value):
            continue
        out[header] = value
    return out


def _number(row: pd.Series | None, column: str):
    if row is None or column not in row.index:
        return None
    value = row[column]
    if value is None or pd.isna(value):
        return None
    return value


def comparison_rows(kind: str, banks: list, flows: list, group_no, code: str, name: str, day: str) -> list[dict]:
    bank_total = round(sum(paid_amount(row) for row in banks), 2) if banks else None
    flow_total = round(sum(paid_amount(row) for row in flows), 2) if flows else None
    base = {
        "项目编码": code,
        "项目名称": name,
        "对照日期": None if day == "9999-99-99" else day,
        "匹配组": group_no,
        "对应关系": kind,
        "本组银行实付": bank_total,
        "本组流水实付": flow_total,
    }

    def one(bank_row, flow_row) -> dict:
        record = dict(base)
        record.update(side_values(bank_row, BANK_FIELDS))
        record.update(side_values(flow_row, FLOW_FIELDS))
        record["实付差值"] = round(paid_amount(bank_row) - paid_amount(flow_row), 2)
        return record

    if kind == ONE_TO_ONE:
        return [one(banks[0], flows[0])]
    rows = [one(row, None) for row in banks]
    rows.extend(one(None, row) for row in flows)
    return rows


def complete_rows(kind: str, banks: list, flows: list, group_no, code: str, name: str, day: str) -> list[dict]:
    """流水行保留税额；银行账实付记在银行那一行上。实付金额是可加总的并集。"""
    bank_total = round(sum(paid_amount(row) for row in banks), 2) if banks else None
    flow_total = round(sum(paid_amount(row) for row in flows), 2) if flows else None
    day_value = None if day == "9999-99-99" else day

    def pack(bank_row, flow_row, include_paid: bool) -> dict:
        source = flow_row if flow_row is not None else bank_row
        paid = None
        if include_paid:
            paid = paid_amount(flow_row) if flow_row is not None else paid_amount(bank_row)
            paid = round(paid, 2)
        tax_source = flow_row if flow_row is not None else (bank_row if kind == BANK_ONLY else None)
        record = {
            "项目编码": code,
            "项目名称": name,
            "日期": day_value,
            "匹配组": group_no,
            "对应关系": kind,
            "摘要": _text(flow_row, "摘要", "业务摘要") or _text(bank_row, "摘要", "业务摘要"),
            "应付金额": _number(tax_source, "应付金额"),
            "预扣税额": _number(tax_source, "预扣税额"),
            "实付金额": paid,
            "增值税额": _number(tax_source, "增值税额"),
            "不含税额": _number(tax_source, "不含税额"),
            "付款金额": _number(tax_source, "付款金额"),
            "银行账摘要": _text(bank_row, "摘要") or None,
            "银行账实付金额": None if bank_row is None else round(paid_amount(bank_row), 2),
            "流水摘要": _text(flow_row, "摘要") or None,
            "流水实付金额": None if flow_row is None else round(paid_amount(flow_row), 2),
            "本组银行实付": bank_total,
            "本组流水实付": flow_total,
        }
        if kind == BANK_ONLY and record["付款金额"] is None and bank_row is not None:
            record["付款金额"] = _number(bank_row, "付款金额")
        return record

    if kind == ONE_TO_ONE:
        return [pack(banks[0], flows[0], True)]
    rows = [pack(row, None, kind == BANK_ONLY) for row in banks]
    rows.extend(pack(None, row, True) for row in flows)
    return rows


def pair_project(bank: pd.DataFrame | None, flow: pd.DataFrame | None, key: str) -> tuple[list[dict], list[dict]]:
    bank_rows = [] if bank is None else [row for _, row in bank.iterrows()]
    flow_rows = [] if flow is None else [row for _, row in flow.iterrows()]
    bank_rows.sort(key=lambda row: (day_key(row.get("日期")) or "9999-99-99",))
    flow_rows.sort(key=lambda row: (day_key(row.get("日期")) or "9999-99-99",))

    days: list[str] = []
    seen: set[str] = set()
    for row in bank_rows + flow_rows:
        day = day_key(row.get("日期")) or "9999-99-99"
        if day not in seen:
            seen.add(day)
            days.append(day)
    days.sort()

    compared: list[dict] = []
    complete: list[dict] = []
    code, name = split_project_key(key)
    group_no = 0
    for day in days:
        left = [row for row in bank_rows if (day_key(row.get("日期")) or "9999-99-99") == day]
        right = [row for row in flow_rows if (day_key(row.get("日期")) or "9999-99-99") == day]
        for kind, bank_idx, flow_idx in match_day(left, right):
            banks = [left[idx] for idx in bank_idx]
            flows = [right[idx] for idx in flow_idx]
            number = None
            if kind in MATCHED_KINDS:
                group_no += 1
                number = group_no
            compared.extend(comparison_rows(kind, banks, flows, number, code, name, day))
            complete.extend(complete_rows(kind, banks, flows, number, code, name, day))
    return compared, complete


def split_project_key(key: str) -> tuple[str, str]:
    """项目键是「项目编码_项目名称」，名称里可以再带下划线或连字符。"""
    code, sep, name = key.partition("_")
    if not sep:
        return "", key
    return code, name


def detail_columns() -> list[str]:
    return (
        ["项目编码", "项目名称", "对照日期", "匹配组", "对应关系", "本组银行实付", "本组流水实付"]
        + [header for _source, header in BANK_FIELDS]
        + ["实付差值"]
        + [header for _source, header in FLOW_FIELDS]
    )


def complete_columns() -> list[str]:
    return [
        "项目编码",
        "项目名称",
        "日期",
        "匹配组",
        "对应关系",
        "摘要",
        "应付金额",
        "预扣税额",
        "实付金额",
        "增值税额",
        "不含税额",
        "付款金额",
        "银行账摘要",
        "银行账实付金额",
        "流水摘要",
        "流水实付金额",
        "本组银行实付",
        "本组流水实付",
    ]


def safe_workbook_name(key: str) -> str:
    text = key.strip() or "未命名项目"
    for char in '\\/:*?"<>|':
        text = text.replace(char, "_")
    return text[:150] + ".xlsx"


def one_project_summary(key: str, detail: pd.DataFrame, complete: pd.DataFrame) -> dict:
    bank_paid = pd.to_numeric(detail["银行账_实付金额"], errors="coerce").sum(min_count=1)
    flow_paid = pd.to_numeric(detail["流水_实付金额"], errors="coerce").sum(min_count=1)
    diff = pd.to_numeric(detail["实付差值"], errors="coerce").sum()
    included = pd.to_numeric(complete["实付金额"], errors="coerce").sum(min_count=1)
    code, name = split_project_key(key)
    matched = detail[detail["对应关系"].isin(MATCHED_KINDS)]
    combo_groups = matched.loc[matched["对应关系"] != ONE_TO_ONE, "匹配组"].nunique()
    return {
        "项目编码": code,
        "项目名称": name,
        "文件名": safe_workbook_name(key),
        "银行账行数": int(detail["银行账_实付金额"].notna().sum()) if not detail.empty else 0,
        "流水行数": int(detail["流水_实付金额"].notna().sum()) if not detail.empty else 0,
        "一对一行数": int((detail["对应关系"] == ONE_TO_ONE).sum()) if not detail.empty else 0,
        "组合组数": int(combo_groups),
        "仅银行行数": int((detail["对应关系"] == BANK_ONLY).sum()) if not detail.empty else 0,
        "仅流水行数": int((detail["对应关系"] == FLOW_ONLY).sum()) if not detail.empty else 0,
        "银行账实付合计": None if pd.isna(bank_paid) else round(float(bank_paid), 2),
        "流水实付合计": None if pd.isna(flow_paid) else round(float(flow_paid), 2),
        "实付差值合计": round(float(diff), 2) if not detail.empty else 0.0,
        "完整实付合计": None if pd.isna(included) else round(float(included), 2),
    }


def _excel_writer(path: Path):
    # openpyxl 引擎不会把 date_format 传给父类，默认格式 YYYY-MM-DD 会让 Excel 要求修复。
    writer = pd.ExcelWriter(path, engine="openpyxl")
    writer._date_format = "yyyy-mm-dd"
    writer._datetime_format = "yyyy-mm-dd"
    return writer


def write_project_book(detail: pd.DataFrame, complete: pd.DataFrame, path: Path) -> None:
    one_sided = detail[detail["对应关系"].isin([BANK_ONLY, FLOW_ONLY])].copy()
    path.parent.mkdir(parents=True, exist_ok=True)
    with _excel_writer(path) as writer:
        detail.to_excel(writer, sheet_name="按日对照", index=False)
        one_sided.to_excel(writer, sheet_name="仅差额", index=False)
        complete.to_excel(writer, sheet_name="完整台账", index=False)
        for sheet_name in ("按日对照", "仅差额", "完整台账"):
            style_sheet(writer.book[sheet_name])


def write_summary_book(rows: list[dict], path: Path) -> None:
    frame = pd.DataFrame(rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    with _excel_writer(path) as writer:
        frame.to_excel(writer, sheet_name="项目汇总", index=False)
        style_sheet(writer.book["项目汇总"], highlight_diff=False)


def write_complete_book(frame: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with _excel_writer(path) as writer:
        frame.to_excel(writer, sheet_name="完整台账", index=False)
        style_sheet(writer.book["完整台账"])


def style_sheet(ws, highlight_diff: bool = True) -> None:
    header_fill = {
        "银行": PatternFill("solid", fgColor="D6EAF8"),
        "流水": PatternFill("solid", fgColor="D5F5E3"),
        "mid": PatternFill("solid", fgColor="FDEBD0"),
    }
    red_fill = PatternFill("solid", fgColor="F5B7B1")
    green_fill = PatternFill("solid", fgColor="A9DFBF")
    pair_fill = PatternFill("solid", fgColor="D5F5E3")
    band_fills = [
        PatternFill("solid", fgColor="EAF2F8"),
        PatternFill("solid", fgColor="FEF9E7"),
    ]
    header_font = Font(bold=True)
    center = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30
    headers = [cell.value for cell in ws[1]]
    for cell in ws[1]:
        text = str(cell.value or "")
        if text.startswith("银行账"):
            cell.fill = header_fill["银行"]
        elif text.startswith("流水"):
            cell.fill = header_fill["流水"]
        else:
            cell.fill = header_fill["mid"]
        cell.font = header_font
        cell.alignment = center
    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        if header in AMOUNT_HEADERS:
            ws.column_dimensions[letter].width = 16
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col_idx).number_format = '#,##0.00'
        elif header in DATE_HEADERS:
            ws.column_dimensions[letter].width = 14
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col_idx).number_format = "yyyy-mm-dd"
        elif header and "摘要" in str(header):
            ws.column_dimensions[letter].width = 42
        elif header == "项目名称":
            ws.column_dimensions[letter].width = 28
        elif header == "对应关系":
            ws.column_dimensions[letter].width = 22
        else:
            ws.column_dimensions[letter].width = 18
    if not highlight_diff or ws.max_row < 2 or "对应关系" not in headers:
        return
    kind_col = headers.index("对应关系") + 1
    group_col = headers.index("匹配组") + 1 if "匹配组" in headers else None
    paint_cols = []
    for name in ("对应关系", "实付差值", "银行账_实付金额", "流水_实付金额", "实付金额", "银行账实付金额", "流水实付金额"):
        if name in headers:
            paint_cols.append(headers.index(name) + 1)
    for row in range(2, ws.max_row + 1):
        kind = str(ws.cell(row, kind_col).value or "")
        if kind == ONE_TO_ONE:
            for col in paint_cols:
                ws.cell(row, col).fill = pair_fill
            if "实付差值" in headers:
                ws.cell(row, headers.index("实付差值") + 1).fill = green_fill
        elif kind in {BANK_ONLY, FLOW_ONLY}:
            for col in paint_cols:
                ws.cell(row, col).fill = red_fill
        elif kind in MATCHED_KINDS:
            group_value = ws.cell(row, group_col).value if group_col else 0
            try:
                band = band_fills[int(group_value) % 2]
            except (TypeError, ValueError):
                band = band_fills[0]
            for col in range(1, len(headers) + 1):
                ws.cell(row, col).fill = band
            if "实付差值" in headers:
                ws.cell(row, headers.index("实付差值") + 1).fill = green_fill


def prepare_output_dir() -> None:
    legacy = OUT_DIR.with_suffix(".xlsx")
    if legacy.is_file():
        legacy.unlink()
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def main() -> int:
    if not BANK_DIR.is_dir() or not FLOW_DIR.is_dir():
        print("找不到银行账或流水目录", file=sys.stderr)
        return 1
    print(f"银行账: {BANK_DIR}")
    print(f"流水:   {FLOW_DIR}")
    bank_books = list_books(BANK_DIR)
    flow_books = list_books(FLOW_DIR)
    bank_frames = {key: read_detail(path) for key, path in bank_books.items()}
    flow_frames = {key: read_detail(path) for key, path in flow_books.items()}
    moved = attach_subsidiary_bank(bank_frames, flow_frames)
    print(f"子公司银行账并入 {SUBSIDIARY_FLOW_KEY}: {moved} 行")
    keys = sorted(set(bank_books) | set(flow_books) | set(bank_frames) | set(flow_frames))
    prepare_output_dir()
    columns = detail_columns()
    full_columns = complete_columns()
    summaries: list[dict] = []
    complete_frames: list[pd.DataFrame] = []
    written: set[str] = {SUMMARY_NAME, COMPLETE_NAME}
    total_rows = 0
    total_one_sided = 0
    for key in keys:
        bank = bank_frames.get(key)
        flow = flow_frames.get(key)
        if bank is not None and bank.empty:
            bank = None
        if flow is not None and flow.empty:
            flow = None
        if bank is None and flow is None:
            continue
        compared, complete = pair_project(bank, flow, key)
        detail = pd.DataFrame(compared, columns=columns)
        full = pd.DataFrame(complete, columns=full_columns)
        for frame in (detail, full):
            if "匹配组" in frame.columns:
                frame["匹配组"] = frame["匹配组"].astype("Int64")
        name = safe_workbook_name(key)
        write_project_book(detail, full, OUT_DIR / name)
        summaries.append(one_project_summary(key, detail, full))
        if not full.empty:
            complete_frames.append(full)
        written.add(name)
        one_sided = int(detail["对应关系"].isin([BANK_ONLY, FLOW_ONLY]).sum()) if not detail.empty else 0
        total_rows += len(detail)
        total_one_sided += one_sided
        print(f"  {name}: 对照 {len(detail)} 行，未配上 {one_sided} 行")
    write_summary_book(summaries, OUT_DIR / SUMMARY_NAME)
    merged = pd.concat(complete_frames, ignore_index=True) if complete_frames else pd.DataFrame(columns=full_columns)
    write_complete_book(merged, OUT_DIR / COMPLETE_NAME)
    for old in OUT_DIR.glob("*.xlsx"):
        if old.name not in written and not old.name.startswith("~$"):
            old.unlink()
    print(f"写出 {len(keys)} 个项目文件 → {OUT_DIR}")
    print(f"对照 {total_rows} 行，其中仅一侧 {total_one_sided} 行；完整台账 {len(merged)} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
