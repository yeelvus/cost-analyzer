# -*- coding: utf-8 -*-
"""把两份 2024-08-12 以后的原始流水、银行账按发生日期左右对照。

银行账在左，流水在右。配对规则调用 compare_flow_bank.match_day：
同一天里先配实付相同的单行，再配一笔对多笔、多笔对一笔和组合，
对不上的留在仅银行账或仅流水。金额按分比较，允许 1 分误差。

银行账的金额列是「付款金额THB」，这里当作实付。流水用「实付金额」，
不用付款金额（付款金额含预扣税）。付款金额THB 为空的是收款、结息，
不参与配对，原样记为仅银行账。

银行账项目名带「子公司-2026年2月以前」这类后缀，流水是「子公司」。
先把以「子公司」开头的名称收成同一项目，以「INSEE」开头的收成 INSEE，
再按这个项目和发生日期配对。

供应商再收一次：同一项目、同一天里没配上的流水，供应商名称相同的
整组实付正好等于剩下的一笔银行账，就收成一组。银行摘要不必写报销
或差旅。不拆这个供应商里面的子集，供应商空着的行不参与。别的供应商
对不上的行留在原处。金额默认是泰铢。整组或单笔乘当月人民币、美元汇率
后能对上银行账时，应付、预扣、实付、增值税、不含税保留原币并标上
CNY 或 USD，付款金额写成换算后的泰铢。分位四舍五入的尾差补在该组
最后一笔，使付款金额合计等于银行实付。行上已有增值税额的视为泰铢税票，
不按外币换算。单笔外币只要求这个供应商当天只剩这一笔，而且只有一种
货币、一笔银行账对得上。同一天里先配泰铢整组，再配外币整组。
配上的流水把该组银行账摘要写入备注。

项目内收完之后，同一天里同一个供应商的剩余流水如果散在两个以上项目，
再把这些行当成一整组，对任意项目里剩下的一笔银行账。先按泰铢合计，
再按无增值税的外币整组。对上的对照项目用这些项目名加号连起来。
对不上的行仍留在原来的项目。

上面两步都做过之后，同一天里两边还剩下的行，如果实付合计相同，
就再合成一组。这样多笔银行账可以对应多笔流水，项目名不一致的
同一笔也会留在这一组里。只在整天剩下的合计对得上时才合并。

完整台账整表按日期从早到晚排列，同一天里再按对照项目放在一起。
没有配上的行排在该项目当天已配对行后面：先是仅银行账，再是仅流水，
两边各自按实付金额从大到小，不交错。
仅银行账和仅流水的实付相差在 500 以内时，整行标成浅黄。

不改 compare_flow_bank.py，也不写按项目拆开的标准对照目录。
"""

from __future__ import annotations

import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from openpyxl.styles import PatternFill

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from compare_flow_bank import (  # noqa: E402
    BANK_ONLY,
    BANK_TO_MANY,
    COMBO,
    FLOW_ONLY,
    MANY_TO_BANK,
    MATCHED_KINDS,
    ONE_TO_ONE,
    _excel_writer,
    _number,
    _text,
    day_key,
    match_day,
    style_sheet,
)

FLOW_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/流水 copy"
    "/2024年8月12以后_除去确认项目外_流水.xlsx"
)
BANK_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款/最新银行账"
    "/账簿汇总_2024-08-12以后 copy.xlsx"
)
OUT_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/01_Raw/12_财务付款"
    "/2024-08-12以后_流水与银行账按日对照.xlsx"
)
FX_PATH = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/汇率三表.xlsx"
)
NEAR_GAP = 500
NEAR_FILL = PatternFill("solid", fgColor="F9E79F")
_YM_RE = re.compile(r"(20\d{2})\D*(\d{1,2})")

DETAIL_COLUMNS = [
    "对照项目",
    "对照日期",
    "匹配组",
    "对应关系",
    "本组银行实付",
    "本组流水实付",
    "银行账_全序号",
    "银行账_项目编码",
    "银行账_项目名称",
    "银行账_摘要",
    "银行账_实付金额",
    "银行账_日期",
    "实付差值",
    "流水_项目编码",
    "流水_项目名称",
    "流水_摘要",
    "流水_应付金额",
    "流水_预扣税额",
    "流水_实付金额",
    "流水_增值税额",
    "流水_不含税额",
    "流水_付款金额",
    "流水_供应商名称",
    "备注",
    "流水_日期",
]
COMPLETE_COLUMNS = [
    "对照项目",
    "项目编码",
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
    "供应商名称",
    "备注",
    "银行账_全序号",
    "银行账项目名称",
    "银行账摘要",
    "银行账实付金额",
    "流水项目名称",
    "流水摘要",
    "流水实付金额",
    "本组银行实付",
    "本组流水实付",
]


def family_name(value) -> str:
    """银行账的期间后缀和流水的项目名收成同一个对照项目。"""
    if value is None or pd.isna(value):
        return "未命名项目"
    text = str(value).strip()
    if not text:
        return "未命名项目"
    if text.startswith("子公司"):
        return "子公司"
    if text.startswith("INSEE"):
        return "INSEE"
    return text


def load_flow() -> pd.DataFrame:
    frame = pd.read_excel(FLOW_PATH, sheet_name="明细")
    frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce")
    frame["实付金额"] = pd.to_numeric(frame["实付金额"], errors="coerce")
    frame["_family"] = frame["项目名称"].map(family_name)
    return frame


def load_bank() -> pd.DataFrame:
    frame = pd.read_excel(BANK_PATH, sheet_name="汇总")
    frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce")
    # 银行账没有单独的实付列。付款金额THB 就是实付；空值是收款、结息。
    frame["实付金额"] = pd.to_numeric(frame["付款金额THB"], errors="coerce")
    frame["_family"] = frame["项目名称"].map(family_name)
    return frame


def cell_value(row: pd.Series | None, column: str):
    if row is None or column not in row.index:
        return None
    value = row[column]
    try:
        missing = value is None or pd.isna(value)
    except TypeError:
        missing = False
    if missing:
        return None
    return value


def paid_or_none(row: pd.Series | None):
    value = cell_value(row, "实付金额")
    if value is None:
        return None
    return round(float(value), 2)


def as_date(value):
    if value is None or value == "" or value == "9999-99-99":
        return None
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    stamp = pd.Timestamp(value)
    if pd.isna(stamp):
        return None
    return stamp.date()


def seq_value(row: pd.Series | None):
    value = cell_value(row, "全序号")
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if number.is_integer():
        return int(number)
    return value


def line_diff(bank_row, flow_row):
    bank_paid = matched_paid(bank_row)
    flow_paid = matched_paid(flow_row)
    if bank_paid is None and flow_paid is None:
        return None
    return round((bank_paid or 0.0) - (flow_paid or 0.0), 2)


def total_of(rows: list) -> float | None:
    amounts = [matched_paid(row) for row in rows]
    amounts = [value for value in amounts if value is not None]
    if not amounts:
        return None
    return round(sum(amounts), 2)


def comparison_record(kind, bank_row, flow_row, group_no, family, day, bank_total, flow_total, remark) -> dict:
    return {
        "对照项目": family,
        "对照日期": as_date(day),
        "匹配组": group_no,
        "对应关系": kind,
        "本组银行实付": bank_total,
        "本组流水实付": flow_total,
        "银行账_全序号": seq_value(bank_row),
        "银行账_项目编码": cell_value(bank_row, "项目编码"),
        "银行账_项目名称": cell_value(bank_row, "项目名称"),
        "银行账_摘要": _text(bank_row, "摘要") or None,
        "银行账_实付金额": paid_or_none(bank_row),
        "银行账_日期": as_date(cell_value(bank_row, "日期")),
        "实付差值": line_diff(bank_row, flow_row),
        "流水_项目编码": cell_value(flow_row, "项目编码"),
        "流水_项目名称": cell_value(flow_row, "项目名称"),
        "流水_摘要": _text(flow_row, "摘要") or None,
        "流水_应付金额": shown_amount(flow_row, "应付金额"),
        "流水_预扣税额": shown_amount(flow_row, "预扣税额"),
        "流水_实付金额": shown_amount(flow_row, "实付金额"),
        "流水_增值税额": shown_amount(flow_row, "增值税额"),
        "流水_不含税额": shown_amount(flow_row, "不含税额"),
        "流水_付款金额": shown_payment(flow_row),
        "流水_供应商名称": cell_value(flow_row, "供应商名称"),
        "备注": remark,
        "流水_日期": as_date(cell_value(flow_row, "日期")),
    }


def comparison_rows(kind, banks, flows, group_no, family, day) -> list[dict]:
    bank_total = total_of(banks)
    flow_total = total_of(flows)
    remark = bank_remark(banks) if kind in MATCHED_KINDS else None
    if kind == ONE_TO_ONE:
        return [comparison_record(kind, banks[0], flows[0], group_no, family, day, bank_total, flow_total, remark)]
    rows = [
        comparison_record(kind, row, None, group_no, family, day, bank_total, flow_total, None)
        for row in banks
    ]
    rows.extend(
        comparison_record(kind, None, row, group_no, family, day, bank_total, flow_total, remark)
        for row in flows
    )
    return rows


def complete_record(kind, bank_row, flow_row, group_no, family, day, bank_total, flow_total, include_paid, remark) -> dict:
    tax_source = flow_row if flow_row is not None else None
    code = cell_value(flow_row, "项目编码") if flow_row is not None else None
    if code is None:
        code = cell_value(bank_row, "项目编码")
    summary = _text(flow_row, "摘要") or _text(bank_row, "摘要") or None
    if include_paid and flow_row is not None and fx_currency(flow_row):
        paid = shown_amount(flow_row, "实付金额")
    elif include_paid:
        source = flow_row if flow_row is not None else bank_row
        paid = paid_or_none(source)
    else:
        paid = None
    return {
        "对照项目": family,
        "项目编码": code,
        "日期": as_date(day),
        "匹配组": group_no,
        "对应关系": kind,
        "摘要": summary,
        "应付金额": shown_amount(tax_source, "应付金额"),
        "预扣税额": shown_amount(tax_source, "预扣税额"),
        "实付金额": paid,
        "增值税额": shown_amount(tax_source, "增值税额"),
        "不含税额": shown_amount(tax_source, "不含税额"),
        "付款金额": shown_payment(tax_source),
        "供应商名称": cell_value(flow_row, "供应商名称") if flow_row is not None else None,
        "备注": remark,
        "银行账_全序号": seq_value(bank_row),
        "银行账项目名称": cell_value(bank_row, "项目名称"),
        "银行账摘要": _text(bank_row, "摘要") or None,
        "银行账实付金额": paid_or_none(bank_row),
        "流水项目名称": cell_value(flow_row, "项目名称"),
        "流水摘要": _text(flow_row, "摘要") or None,
        "流水实付金额": shown_amount(flow_row, "实付金额"),
        "本组银行实付": bank_total,
        "本组流水实付": flow_total,
    }


def complete_rows(kind, banks, flows, group_no, family, day) -> list[dict]:
    bank_total = total_of(banks)
    flow_total = total_of(flows)
    remark = bank_remark(banks) if kind in MATCHED_KINDS else None
    if kind == ONE_TO_ONE:
        return [complete_record(kind, banks[0], flows[0], group_no, family, day, bank_total, flow_total, True, remark)]
    rows = [
        complete_record(kind, row, None, group_no, family, day, bank_total, flow_total, kind == BANK_ONLY, None)
        for row in banks
    ]
    rows.extend(
        complete_record(kind, None, row, group_no, family, day, bank_total, flow_total, True, remark)
        for row in flows
    )
    return rows


def rows_of(frame: pd.DataFrame, family: str) -> list[pd.Series]:
    part = frame[frame["_family"] == family]
    return [row for _, row in part.iterrows()]


def bucket_by_day(rows: list[pd.Series]) -> dict[str, list[pd.Series]]:
    buckets: dict[str, list[pd.Series]] = defaultdict(list)
    for row in rows:
        buckets[day_key(row.get("日期")) or "9999-99-99"].append(row)
    return buckets


def kind_of(bank_count: int, flow_count: int) -> str:
    if bank_count == 1 and flow_count == 1:
        return ONE_TO_ONE
    if bank_count == 1 and flow_count > 1:
        return BANK_TO_MANY
    if flow_count == 1 and bank_count > 1:
        return MANY_TO_BANK
    return COMBO


def cents(amount: float) -> int:
    return int(round(amount * 100))


def matched_paid(row: pd.Series | None):
    """配对用的泰铢。外币行用换算后的付款金额，避免把人民币数当成泰铢。"""
    if row is not None and "_pay_thb" in row.index:
        value = row["_pay_thb"]
        try:
            missing = value is None or pd.isna(value)
        except TypeError:
            missing = False
        if not missing:
            return round(float(value), 2)
    return paid_or_none(row)


def fx_currency(row: pd.Series | None) -> str:
    if row is None or "_fx_currency" not in row.index:
        return ""
    value = row["_fx_currency"]
    try:
        if value is None or pd.isna(value):
            return ""
    except TypeError:
        return ""
    return str(value).strip()


def money_label(value, currency: str):
    cents_value = int(round(float(value) * 100))
    sign = "-" if cents_value < 0 else ""
    cents_value = abs(cents_value)
    if cents_value % 100 == 0:
        body = str(cents_value // 100)
    else:
        body = f"{cents_value // 100}.{cents_value % 100:02d}"
    return f"{sign}{body}{currency}"


def shown_amount(row: pd.Series | None, column: str):
    value = _number(row, column)
    if value is None:
        return None
    currency = fx_currency(row)
    if not currency:
        return value
    return money_label(value, currency)


def shown_payment(row: pd.Series | None):
    if row is None:
        return None
    if fx_currency(row):
        return matched_paid(row)
    return _number(row, "付款金额")


def bank_remark(banks: list) -> str | None:
    seen: list[str] = []
    for row in banks:
        text = _text(row, "摘要")
        if text and text not in seen:
            seen.append(text)
    return "；".join(seen) or None


def supplier_key(row: pd.Series) -> str:
    value = cell_value(row, "供应商名称")
    if value is None:
        return ""
    return " ".join(str(value).split())


def has_vat(row: pd.Series) -> bool:
    value = cell_value(row, "增值税额")
    if value is None:
        return False
    try:
        return float(value) != 0
    except (TypeError, ValueError):
        return True


def amounts_of(rows: list) -> list[float] | None:
    values = []
    for row in rows:
        amount = paid_or_none(row)
        if amount is None:
            return None
        values.append(amount)
    return values


def load_month_rates(path: Path) -> dict[str, tuple[float | None, float | None]]:
    """年月 → (1 CNY 兑泰铢, 1 USD 兑泰铢)。表里没有的月份不换算。"""
    if not path.is_file():
        print(f"找不到汇率表 {path}，外币不换算", flush=True)
        return {}
    table = pd.read_excel(path, sheet_name="汇率表")
    folded = {str(col).replace("\n", "").replace(" ", ""): col for col in table.columns}
    ym_col = next((folded[name] for name in folded if "年月" in name), None)
    cny_col = next((folded[name] for name in folded if "人民币对泰铢" in name), None)
    usd_col = next((folded[name] for name in folded if "美元对泰铢" in name), None)
    if ym_col is None or cny_col is None or usd_col is None:
        print(f"汇率表缺列: {list(table.columns)}", flush=True)
        return {}
    rates: dict[str, tuple[float | None, float | None]] = {}
    for _, row in table.iterrows():
        found = _YM_RE.search(str(row[ym_col]))
        if not found:
            continue
        month = int(found.group(2))
        if not 1 <= month <= 12:
            continue
        key = f"{found.group(1)}-{month:02d}"
        cny = pd.to_numeric(row[cny_col], errors="coerce")
        usd = pd.to_numeric(row[usd_col], errors="coerce")
        cny_rate = float(cny) if pd.notna(cny) and float(cny) > 0 else None
        usd_rate = float(usd) if pd.notna(usd) and float(usd) > 0 else None
        rates[key] = (cny_rate, usd_rate)
    return rates


def rates_for(day: str, table: dict) -> list[tuple[str, float]]:
    if not day or len(day) < 7:
        return []
    cny_rate, usd_rate = table.get(day[:7], (None, None))
    found = []
    if cny_rate:
        found.append(("CNY", cny_rate))
    if usd_rate:
        found.append(("USD", usd_rate))
    return found


def converted_cents(amounts: list[float], rate: float) -> int:
    return int(round(sum(amounts) * rate * 100))


def allocate_thb(amounts: list[float], rate: float, bank_cents: int) -> list[float] | None:
    """各笔先按汇率保留到分，尾差补在最后一笔，使合计等于银行实付。"""
    line_cents = [int(round(amount * rate * 100)) for amount in amounts]
    drift = bank_cents - sum(line_cents)
    if abs(drift) > max(1, len(amounts)):
        return None
    line_cents[-1] += drift
    return [value / 100 for value in line_cents]


def with_fx(row: pd.Series, currency: str, pay_thb: float) -> pd.Series:
    marked = row.copy()
    marked["_fx_currency"] = currency
    marked["_pay_thb"] = float(pay_thb)
    return marked


def _near_banks(banks: list, target: int) -> list:
    found = []
    for row in banks:
        amount = paid_or_none(row)
        if amount is None:
            continue
        if abs(cents(amount) - target) <= 1:
            found.append(row)
    return found


def _unique_options(options: list[tuple]) -> list[tuple]:
    """同一供应商或同一笔银行账出现两次就不猜。"""
    supplier_count = Counter(item[0] for item in options)
    bank_count = Counter(id(item[2]) for item in options)
    return [
        item
        for item in options
        if supplier_count[item[0]] == 1 and bank_count[id(item[2])] == 1
    ]


def pair_supplier_groups(banks: list, flows: list, day: str, family: str, rate_table: dict):
    """整组供应商对上一笔银行账。返回新组、剩下的银行账、剩下的流水、外币原币分、外币泰铢分。"""
    made: list[tuple] = []
    used_bank: set[int] = set()
    used_flow: set[int] = set()
    fx_original_cents = 0
    fx_thb_cents = 0
    by_supplier: dict[str, list] = defaultdict(list)
    for row in flows:
        key = supplier_key(row)
        if key:
            by_supplier[key].append(row)

    thb_options = []
    for key, rows in by_supplier.items():
        if len(rows) < 2:
            continue
        amounts = amounts_of(rows)
        if not amounts:
            continue
        near = _near_banks(banks, sum(cents(amount) for amount in amounts))
        if len(near) == 1:
            thb_options.append((key, rows, near[0], amounts))
    for key, rows, bank, amounts in _unique_options(thb_options):
        kind = BANK_TO_MANY if len(rows) > 1 else ONE_TO_ONE
        made.append((kind, [bank], list(rows), day, family))
        used_bank.add(id(bank))
        used_flow.update(id(row) for row in rows)
        print(
            f"供应商对应 {day} {family} {key}: {len(rows)} 笔 {sum(amounts):.2f}"
            f" → {_text(bank, '摘要')} {paid_or_none(bank):.2f} {kind}",
            flush=True,
        )

    remain_banks = [row for row in banks if id(row) not in used_bank]
    currency_rates = rates_for(day, rate_table)
    fx_options = []
    if currency_rates:
        for key, rows in by_supplier.items():
            if len(rows) < 2 or any(id(row) in used_flow for row in rows):
                continue
            if any(has_vat(row) for row in rows):
                continue
            amounts = amounts_of(rows)
            if not amounts:
                continue
            for currency, rate in currency_rates:
                near = _near_banks(remain_banks, converted_cents(amounts, rate))
                if len(near) != 1:
                    continue
                bank = near[0]
                pays = allocate_thb(amounts, rate, cents(paid_or_none(bank)))
                if pays is None:
                    continue
                fx_options.append((key, rows, bank, currency, rate, amounts, pays))
        for key, rows, bank, currency, rate, amounts, pays in _unique_options(fx_options):
            marked = [with_fx(row, currency, pay) for row, pay in zip(rows, pays)]
            made.append((BANK_TO_MANY, [bank], marked, day, family))
            used_bank.add(id(bank))
            used_flow.update(id(row) for row in rows)
            fx_original_cents += sum(cents(amount) for amount in amounts)
            fx_thb_cents += sum(cents(pay) for pay in pays)
            print(
                f"外币对应 {day} {family} {key} {currency}×{rate}: {len(rows)} 笔 {sum(amounts):.2f}"
                f" → {sum(pays):.2f} {_text(bank, '摘要')}",
                flush=True,
            )

        remain_banks = [row for row in banks if id(row) not in used_bank]
        single_options = []
        for key, rows in by_supplier.items():
            if len(rows) != 1 or any(id(row) in used_flow for row in rows):
                continue
            row = rows[0]
            if has_vat(row):
                continue
            amount = paid_or_none(row)
            if amount is None:
                continue
            hits = []
            for bank in remain_banks:
                bank_cents = cents(paid_or_none(bank))
                for currency, rate in currency_rates:
                    if abs(converted_cents([amount], rate) - bank_cents) > 1:
                        continue
                    pays = allocate_thb([amount], rate, bank_cents)
                    if pays is None:
                        continue
                    hits.append((bank, currency, rate, pays[0]))
            if len(hits) != 1:
                continue
            bank, currency, rate, pay = hits[0]
            single_options.append((key, [row], bank, currency, rate, [amount], [pay]))
        for key, rows, bank, currency, rate, amounts, pays in _unique_options(single_options):
            marked = [with_fx(rows[0], currency, pays[0])]
            made.append((ONE_TO_ONE, [bank], marked, day, family))
            used_bank.add(id(bank))
            used_flow.add(id(rows[0]))
            fx_original_cents += cents(amounts[0])
            fx_thb_cents += cents(pays[0])
            print(
                f"外币对应 {day} {family} {key} {currency}×{rate}: 1 笔 {amounts[0]:.2f}"
                f" → {pays[0]:.2f} {_text(bank, '摘要')}",
                flush=True,
            )

    left_banks = [row for row in banks if id(row) not in used_bank]
    left_flows = [row for row in flows if id(row) not in used_flow]
    return made, left_banks, left_flows, fx_original_cents, fx_thb_cents


def bind_reimbursements(groups: list[tuple], day: str, family: str, rate_table: dict):
    """项目内配对之后，再用供应商整组和汇率收当天剩下的行。"""
    kept: list[tuple] = []
    loose_bank = []
    loose_flow = []
    for group in groups:
        kind, banks, flows, _day, _family = group
        if kind == BANK_ONLY and len(banks) == 1 and paid_or_none(banks[0]) is not None:
            loose_bank.append(banks[0])
        elif kind == FLOW_ONLY and len(flows) == 1 and paid_or_none(flows[0]) is not None:
            loose_flow.append(flows[0])
        else:
            kept.append(group)
    fx_original_cents = 0
    fx_thb_cents = 0
    if loose_bank and loose_flow:
        made, loose_bank, loose_flow, fx_original_cents, fx_thb_cents = pair_supplier_groups(
            loose_bank, loose_flow, day, family, rate_table
        )
        kept.extend(made)
    for row in loose_bank:
        kept.append((BANK_ONLY, [row], [], day, family))
    for row in loose_flow:
        kept.append((FLOW_ONLY, [], [row], day, family))
    return kept, fx_original_cents, fx_thb_cents


def _family_label(rows: list) -> str:
    families = sorted({family_name(row.get("项目名称")) for row in rows})
    return "+".join(families)


def bind_cross_family_suppliers(groups: list[tuple], rate_table: dict) -> tuple[list[tuple], int, int]:
    """同一天、同一供应商、至少两个项目的剩余流水，整组对上一笔剩余银行账。

    只收整组，不拆子集。供应商或银行账对上两笔就不猜。对不上的行留在原项目。
    """
    kept: list[tuple] = []
    loose_bank: list[tuple] = []
    loose_flow: list[tuple] = []
    for group in groups:
        kind, banks, flows, _day, _family = group
        if kind == BANK_ONLY and len(banks) == 1 and paid_or_none(banks[0]) is not None:
            loose_bank.append(group)
        elif kind == FLOW_ONLY and len(flows) == 1 and paid_or_none(flows[0]) is not None:
            loose_flow.append(group)
        else:
            kept.append(group)

    by_supplier: dict[tuple[str, str], list] = defaultdict(list)
    for group in loose_flow:
        row = group[2][0]
        key = supplier_key(row)
        if key:
            by_supplier[(group[3], key)].append(row)

    banks_by_day: dict[str, list] = defaultdict(list)
    for group in loose_bank:
        banks_by_day[group[3]].append(group[1][0])

    candidates = []
    for (day, key), rows in by_supplier.items():
        families = {family_name(row.get("项目名称")) for row in rows}
        if len(families) < 2:
            continue
        amounts = amounts_of(rows)
        if not amounts:
            continue
        candidates.append((day, key, rows, amounts))

    used_bank: set[int] = set()
    used_flow: set[int] = set()
    made: list[tuple] = []
    fx_original_cents = 0
    fx_thb_cents = 0

    thb_options = []
    for day, key, rows, amounts in candidates:
        near = _near_banks(banks_by_day.get(day, []), sum(cents(amount) for amount in amounts))
        if len(near) == 1:
            thb_options.append(((day, key), rows, near[0], amounts))
    for (day, key), rows, bank, amounts in _unique_options(thb_options):
        label = _family_label([bank, *rows])
        kind = kind_of(1, len(rows))
        made.append((kind, [bank], list(rows), day, label))
        used_bank.add(id(bank))
        used_flow.update(id(row) for row in rows)
        print(
            f"跨项目供应商 {day} {label} {key}: {len(rows)} 笔 {sum(amounts):.2f}"
            f" → {_text(bank, '摘要')} {paid_or_none(bank):.2f} {kind}",
            flush=True,
        )

    fx_options = []
    for day, key, rows, amounts in candidates:
        if any(id(row) in used_flow for row in rows):
            continue
        if any(has_vat(row) for row in rows):
            continue
        remain = [row for row in banks_by_day.get(day, []) if id(row) not in used_bank]
        for currency, rate in rates_for(day, rate_table):
            near = _near_banks(remain, converted_cents(amounts, rate))
            if len(near) != 1:
                continue
            bank = near[0]
            pays = allocate_thb(amounts, rate, cents(paid_or_none(bank)))
            if pays is None:
                continue
            fx_options.append(((day, key), rows, bank, currency, rate, amounts, pays))
    for (day, key), rows, bank, currency, rate, amounts, pays in _unique_options(fx_options):
        marked = [with_fx(row, currency, pay) for row, pay in zip(rows, pays)]
        label = _family_label([bank, *marked])
        made.append((BANK_TO_MANY, [bank], marked, day, label))
        used_bank.add(id(bank))
        used_flow.update(id(row) for row in rows)
        fx_original_cents += sum(cents(amount) for amount in amounts)
        fx_thb_cents += sum(cents(pay) for pay in pays)
        print(
            f"跨项目外币 {day} {label} {key} {currency}×{rate}: {len(rows)} 笔 {sum(amounts):.2f}"
            f" → {sum(pays):.2f} {_text(bank, '摘要')}",
            flush=True,
        )

    for group in loose_bank:
        if id(group[1][0]) not in used_bank:
            kept.append(group)
    for group in loose_flow:
        if id(group[2][0]) not in used_flow:
            kept.append(group)
    kept.extend(made)
    return kept, fx_original_cents, fx_thb_cents


def bind_same_day(groups: list[tuple]) -> list[tuple]:
    """同一天项目内没配上的行，两边实付合计相同就合成一组。

    groups 里每一项是 (kind, banks, flows, day, family)。
    无金额的收款、结息留在仅银行账，不拿来凑合计。
    """
    kept: list[tuple] = []
    loose_bank: dict[str, list[tuple]] = defaultdict(list)
    loose_flow: dict[str, list[tuple]] = defaultdict(list)
    for group in groups:
        kind, banks, flows, day, _family = group
        if kind == BANK_ONLY and len(banks) == 1 and paid_or_none(banks[0]) is not None:
            loose_bank[day].append(group)
        elif kind == FLOW_ONLY and len(flows) == 1 and paid_or_none(flows[0]) is not None:
            loose_flow[day].append(group)
        else:
            kept.append(group)
    for day in sorted(set(loose_bank) | set(loose_flow)):
        bank_groups = loose_bank.get(day, [])
        flow_groups = loose_flow.get(day, [])
        if not bank_groups or not flow_groups:
            kept.extend(bank_groups)
            kept.extend(flow_groups)
            continue
        bank_rows = [group[1][0] for group in bank_groups]
        flow_rows = [group[2][0] for group in flow_groups]
        bank_sum = sum(cents(paid_or_none(row)) for row in bank_rows)
        flow_sum = sum(cents(paid_or_none(row)) for row in flow_rows)
        if abs(bank_sum - flow_sum) > 1:
            kept.extend(bank_groups)
            kept.extend(flow_groups)
            continue
        families = sorted({family_name(row.get("项目名称")) for row in bank_rows + flow_rows})
        label = "+".join(families)
        kind = kind_of(len(bank_rows), len(flow_rows))
        print(
            f"同一日合并 {day} {label}: 银行 {len(bank_rows)} 笔 {bank_sum/100:.2f}，"
            f"流水 {len(flow_rows)} 笔 {flow_sum/100:.2f}，{kind}",
            flush=True,
        )
        kept.append((kind, bank_rows, flow_rows, day, label))
    kept.sort(key=lambda group: (group[4], group[3]))
    return kept


def pair_frames(bank: pd.DataFrame, flow: pd.DataFrame, rate_table: dict) -> tuple[list[dict], list[dict], int, int]:
    payable = bank[bank["实付金额"].notna()]
    receipts = bank[bank["实付金额"].isna()]
    families = sorted(set(bank["_family"]) | set(flow["_family"]))
    groups: list[tuple] = []
    fx_original_cents = 0
    fx_thb_cents = 0
    for family in families:
        bank_days = bucket_by_day(rows_of(payable, family))
        flow_days = bucket_by_day(rows_of(flow, family))
        receipt_days = bucket_by_day(rows_of(receipts, family))
        days = sorted(set(bank_days) | set(flow_days) | set(receipt_days))
        print(
            f"{family}: 银行 {sum(len(v) for v in bank_days.values())} 行, "
            f"流水 {sum(len(v) for v in flow_days.values())} 行, "
            f"无金额 {sum(len(v) for v in receipt_days.values())} 行, {len(days)} 天",
            flush=True,
        )
        for day in days:
            left = bank_days.get(day, [])
            right = flow_days.get(day, [])
            day_groups = []
            for kind, bank_idx, flow_idx in match_day(left, right):
                banks = [left[idx] for idx in bank_idx]
                flows = [right[idx] for idx in flow_idx]
                day_groups.append((kind, banks, flows, day, family))
            day_groups, original_cents, thb_cents = bind_reimbursements(day_groups, day, family, rate_table)
            fx_original_cents += original_cents
            fx_thb_cents += thb_cents
            groups.extend(day_groups)
            for row in receipt_days.get(day, []):
                groups.append((BANK_ONLY, [row], [], day, family))
    groups, cross_original, cross_thb = bind_cross_family_suppliers(groups, rate_table)
    fx_original_cents += cross_original
    fx_thb_cents += cross_thb
    groups = bind_same_day(groups)
    compared: list[dict] = []
    complete: list[dict] = []
    group_no = 0
    for kind, banks, flows, day, family in groups:
        number = None
        if kind in MATCHED_KINDS:
            group_no += 1
            number = group_no
        compared.extend(comparison_rows(kind, banks, flows, number, family, day))
        complete.extend(complete_rows(kind, banks, flows, number, family, day))
    return compared, complete, fx_original_cents, fx_thb_cents


def frame_of(rows: list[dict], columns: list[str]) -> pd.DataFrame:
    frame = pd.DataFrame(rows, columns=columns)
    if not frame.empty:
        frame["匹配组"] = frame["匹配组"].astype("Int64")
    return frame


def summary_rows(detail: pd.DataFrame, complete: pd.DataFrame) -> list[dict]:
    rows = []
    families = list(dict.fromkeys(detail["对照项目"].tolist()))
    for family in families:
        part = detail[detail["对照项目"] == family]
        full = complete[complete["对照项目"] == family]
        matched = part[part["对应关系"].isin(MATCHED_KINDS)]
        combo_groups = matched.loc[matched["对应关系"] != ONE_TO_ONE, "匹配组"].nunique()
        bank_paid = pd.to_numeric(part["银行账_实付金额"], errors="coerce").sum(min_count=1)
        flow_paid = numeric_with_fallback(part, "流水_实付金额", "流水_付款金额").sum(min_count=1)
        diff = pd.to_numeric(part["实付差值"], errors="coerce").sum(min_count=1)
        included = numeric_with_fallback(full, "实付金额", "付款金额").sum(min_count=1)
        rows.append(
            {
                "对照项目": family,
                "银行账行数": int(part["银行账_日期"].notna().sum()),
                "流水行数": int(part["流水_日期"].notna().sum()),
                "一对一行数": int((part["对应关系"] == ONE_TO_ONE).sum()),
                "组合组数": int(combo_groups),
                "仅银行行数": int((part["对应关系"] == BANK_ONLY).sum()),
                "仅流水行数": int((part["对应关系"] == FLOW_ONLY).sum()),
                "银行账无金额行数": int(part["银行账_日期"].notna().sum() - part["银行账_实付金额"].notna().sum()),
                "银行账实付合计": None if pd.isna(bank_paid) else round(float(bank_paid), 2),
                "流水实付合计": None if pd.isna(flow_paid) else round(float(flow_paid), 2),
                "实付差值合计": None if pd.isna(diff) else round(float(diff), 2),
                "完整实付合计": None if pd.isna(included) else round(float(included), 2),
            }
        )
    return rows


def numeric_with_fallback(frame: pd.DataFrame, primary: str, fallback: str) -> pd.Series:
    """实付被标成 3102CNY 时，合计改用同行的泰铢付款金额。"""
    paid = pd.to_numeric(frame[primary], errors="coerce") if primary in frame.columns else pd.Series(dtype=float)
    if fallback not in frame.columns:
        return paid
    other = pd.to_numeric(frame[fallback], errors="coerce")
    return paid.where(paid.notna(), other)


def _paid_cell(row: pd.Series):
    value = row.get("实付金额")
    try:
        missing = value is None or pd.isna(value)
    except TypeError:
        missing = False
    if missing:
        return None
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def _sort_by_amount(block: pd.DataFrame) -> pd.DataFrame:
    if block.empty:
        return block
    ordered = block.copy()
    ordered["_amt"] = ordered.apply(_paid_cell, axis=1)
    ordered = ordered.sort_values("_amt", ascending=False, na_position="last")
    return ordered.drop(columns="_amt")


def arrange_complete(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[int]]:
    """整表按日期升序。未配对的银行账、流水各自按实付从大到小排，两边不交错。"""
    if frame.empty:
        return frame, []
    frame = frame.copy()
    frame["_day"] = pd.to_datetime(frame["日期"], errors="coerce")
    frame = frame.sort_values(
        ["_day", "对照项目"],
        ascending=[True, True],
        kind="mergesort",
        na_position="last",
    ).drop(columns="_day")
    frame = frame.reset_index(drop=True)
    pieces: list[pd.DataFrame] = []
    near_rows: list[int] = []
    placed = 0
    grouped = frame.groupby(["对照项目", "日期"], sort=False, dropna=False)
    for _, part in grouped:
        matched = part[part["对应关系"].isin(MATCHED_KINDS)]
        open_rows = part[~part["对应关系"].isin(MATCHED_KINDS)].copy()
        banks_only = _sort_by_amount(open_rows[open_rows["对应关系"] == BANK_ONLY])
        flows_only = _sort_by_amount(open_rows[open_rows["对应关系"] == FLOW_ONLY])
        open_rows = pd.concat([banks_only, flows_only], axis=0)
        open_rows["_amt"] = open_rows.apply(_paid_cell, axis=1)
        banks: list[tuple[int, int]] = []
        flows: list[tuple[int, int]] = []
        for offset, (_, row) in enumerate(open_rows.iterrows()):
            amount = row["_amt"]
            if amount is None or pd.isna(amount):
                continue
            cents = int(round(float(amount) * 100))
            if row["对应关系"] == BANK_ONLY:
                banks.append((offset, cents))
            elif row["对应关系"] == FLOW_ONLY:
                flows.append((offset, cents))
        near_offsets: set[int] = set()
        for bank_at, bank_cents in banks:
            for flow_at, flow_cents in flows:
                if abs(bank_cents - flow_cents) <= NEAR_GAP * 100:
                    near_offsets.add(bank_at)
                    near_offsets.add(flow_at)
        first_open = placed + len(matched)
        near_rows.extend(first_open + offset + 2 for offset in sorted(near_offsets))
        placed += len(matched) + len(open_rows)
        open_rows = open_rows.drop(columns=["_amt"])
        pieces.append(pd.concat([matched, open_rows], axis=0))
    ordered = pd.concat(pieces, axis=0).reset_index(drop=True)
    return ordered, near_rows


def mark_near_rows(ws, excel_rows: list[int]) -> None:
    if not excel_rows:
        return
    last_col = ws.max_column
    for row in excel_rows:
        for col in range(1, last_col + 1):
            ws.cell(row, col).fill = NEAR_FILL


def style_book(ws, highlight: bool) -> None:
    style_sheet(ws, highlight_diff=highlight)
    for cell in ws[1]:
        text = str(cell.value or "")
        if text == "对照项目":
            ws.column_dimensions[cell.column_letter].width = 36
        elif "项目名称" in text:
            ws.column_dimensions[cell.column_letter].width = 28
        elif "供应商" in text:
            ws.column_dimensions[cell.column_letter].width = 32
        elif text == "备注":
            ws.column_dimensions[cell.column_letter].width = 42
        elif "全序号" in text:
            ws.column_dimensions[cell.column_letter].width = 12


def check_totals(
    bank: pd.DataFrame,
    flow: pd.DataFrame,
    detail: pd.DataFrame,
    complete: pd.DataFrame,
    fx_original_cents: int,
    fx_thb_cents: int,
) -> None:
    bank_rows = int(bank["日期"].notna().sum())
    flow_rows = int(flow["日期"].notna().sum())
    got_bank = int(detail["银行账_日期"].notna().sum())
    got_flow = int(detail["流水_日期"].notna().sum())
    if got_bank != bank_rows or got_flow != flow_rows:
        raise SystemExit(f"行数对不上: 银行 {got_bank}/{bank_rows}, 流水 {got_flow}/{flow_rows}")
    flow_paid = round(float(pd.to_numeric(flow["实付金额"], errors="coerce").sum()), 2)
    bank_only = complete[complete["对应关系"] == BANK_ONLY]
    bank_only_paid = pd.to_numeric(bank_only["实付金额"], errors="coerce").sum()
    bank_only_paid = 0.0 if pd.isna(bank_only_paid) else round(float(bank_only_paid), 2)
    complete_paid = numeric_with_fallback(complete, "实付金额", "付款金额").sum()
    complete_paid = 0.0 if pd.isna(complete_paid) else round(float(complete_paid), 2)
    fx_original = fx_original_cents / 100
    fx_thb = fx_thb_cents / 100
    expect = round(flow_paid - fx_original + fx_thb + bank_only_paid, 2)
    if abs(complete_paid - expect) > 0.05:
        raise SystemExit(
            f"完整实付 {complete_paid} 不等于流水 {flow_paid} - 外币原币 {fx_original} + 外币泰铢 {fx_thb} + 仅银行账 {bank_only_paid}"
        )
    blank_bank = int((detail["对应关系"] == BANK_ONLY).sum() - detail.loc[detail["对应关系"] == BANK_ONLY, "银行账_实付金额"].notna().sum())
    print(
        f"核对: 银行 {got_bank} 行, 流水 {got_flow} 行, 完整实付 {complete_paid}, "
        f"其中仅银行账实付 {bank_only_paid}, 外币原币 {fx_original:.2f} 换成泰铢 {fx_thb:.2f}, "
        f"无金额银行行 {blank_bank}",
        flush=True,
    )


def print_samples(detail: pd.DataFrame) -> None:
    day = pd.Timestamp("2024-08-12").date()
    sample = detail[(detail["对照项目"] == "子公司") & (detail["对照日期"] == day)]
    keep = ["对应关系", "匹配组", "银行账_摘要", "银行账_实付金额", "流水_摘要", "流水_实付金额", "实付差值"]
    print("--- 2024-08-12 子公司 ---", flush=True)
    print(sample[keep].to_string(index=False), flush=True)
    wages = detail[
        (detail["对照项目"] == "子公司")
        & (detail["对照日期"] == pd.Timestamp("2026-07-25").date())
        & (
            detail["银行账_摘要"].fillna("").str.contains("劳务工资|手续费")
            | detail["流水_摘要"].fillna("").str.contains("劳务工资|手续费")
        )
    ]
    print("--- 2026-07-25 子公司 劳务工资/手续费 ---", flush=True)
    print(wages[keep].to_string(index=False), flush=True)
    show = ["对应关系", "匹配组", "银行账_摘要", "银行账_实付金额", "流水_摘要", "流水_实付金额", "流水_付款金额", "流水_供应商名称", "备注"]
    for text in ("2025-05-19", "2025-05-23"):
        block = detail[(detail["对照项目"] == "子公司") & (detail["对照日期"] == pd.Timestamp(text).date())]
        print(f"--- {text} 子公司 ---", flush=True)
        print(block[show].to_string(index=False), flush=True)
    day = pd.Timestamp("2025-05-06").date()
    block = detail[detail["对照日期"] == day].copy()
    bank_text = block["银行账_摘要"].fillna("")
    flow_text = block["流水_摘要"].fillna("")
    mask = bank_text.str.contains("培训、水电|墨盒|餐具") | flow_text.str.contains(
        "投标保证金|墨粉|餐具|电费|加油费|过路费|网费|开工仪式|安全讲座"
    )
    picked = block.loc[mask, ["对照项目", *show]]
    print("--- 2025-05-06 培训水电/墨盒/餐具/保证金 ---", flush=True)
    print(picked.to_string(index=False), flush=True)


def main() -> int:
    if not FLOW_PATH.is_file() or not BANK_PATH.is_file():
        print("找不到流水或银行账文件", file=sys.stderr)
        return 1
    print(f"流水: {FLOW_PATH}", flush=True)
    print(f"银行账: {BANK_PATH}", flush=True)
    flow = load_flow()
    bank = load_bank()
    rate_table = load_month_rates(FX_PATH)
    print(f"汇率: {FX_PATH}，{len(rate_table)} 个月", flush=True)
    compared, complete, fx_original_cents, fx_thb_cents = pair_frames(bank, flow, rate_table)
    detail = frame_of(compared, DETAIL_COLUMNS)
    full = frame_of(complete, COMPLETE_COLUMNS)
    check_totals(bank, flow, detail, full, fx_original_cents, fx_thb_cents)
    print_samples(detail)
    one_sided = detail[detail["对应关系"].isin([BANK_ONLY, FLOW_ONLY])].copy()
    summary = pd.DataFrame(summary_rows(detail, full))
    full, near_rows = arrange_complete(full)
    print(f"完整台账接近行（相差 {NEAR_GAP} 以内）: {len(near_rows)}", flush=True)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _excel_writer(OUT_PATH) as writer:
        detail.to_excel(writer, sheet_name="按日对照", index=False)
        one_sided.to_excel(writer, sheet_name="仅差额", index=False)
        full.to_excel(writer, sheet_name="完整台账", index=False)
        summary.to_excel(writer, sheet_name="项目汇总", index=False)
        for name in ("按日对照", "仅差额", "完整台账"):
            style_book(writer.book[name], True)
        mark_near_rows(writer.book["完整台账"], near_rows)
        style_book(writer.book["项目汇总"], False)
    print(f"写出 {OUT_PATH}", flush=True)
    print(
        f"按日对照 {len(detail)} 行, 仅差额 {len(one_sided)} 行, 完整台账 {len(full)} 行",
        flush=True,
    )
    print(summary.to_string(index=False), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
