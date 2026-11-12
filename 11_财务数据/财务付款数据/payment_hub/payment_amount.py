# -*- coding: utf-8 -*-
"""付款金额 = 预扣税额(泰铢) + 实付金额(泰铢)。

五列金额格子支持 RMB / CNY / 人民币 / USD / 美元 / THB 前缀或后缀。
五列写出去掉标注后的原币数字，币种写在对应「*币种」列。
源表若已拆成「应付金额/预扣税额/实付金额/增值税额/不含税额」和
「应付币种/预扣币种/实付币种/增值税币种/不含税币种」，数字按原币读，币种看分列。
「币别」或「币种」为 CNY/USD，且格子和分列都没有标注时，该行数字按该币种处理。
预扣、实付带外币时换成泰铢再相加。汇率优先用摘要里的 1CNY=，
其次用本行「汇率」列，再其次用当月汇率表。应付带 RMB、实付已是泰铢的，
只标注应付，实付不再乘汇率。再缺则付款金额留空、汇率缺失=1，不把原币写进泰铢合计。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd

PAYMENT_AMOUNT_COL = "付款金额"
WITHHOLD_PAYMENT_COL = "预扣税付款金额"
WITHHOLD_GROSS_COL = "含预扣税合计"
FX_RATE_COL = "汇率"
YEAR_MONTH_COL = "年月"
BILIE_COL = "币别"
RATE_MISSING_COL = "汇率缺失"
# A1 保留的来源金额列：存原始数字（去掉币种标注，不乘汇率）。
LEGACY_AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]
CURRENCY_COLUMNS = {
    "应付金额": "应付币种",
    "预扣税额": "预扣币种",
    "实付金额": "实付币种",
    "增值税额": "增值税币种",
    "不含税额": "不含税币种",
}
# 相加的两列。带人民币、美元标注（或币别）的换成泰铢再加。
PAYMENT_PART_COLUMNS = ["预扣税额", "实付金额"]

_CNY_RE = re.compile(r"(?i)rmb|cny|人民币|￥|¥")
_USD_RE = re.compile(r"(?i)(?<![A-Za-z])usd(?![A-Za-z])|美元|us\$")
_THB_RE = re.compile(r"(?i)(?<![A-Za-z])thb(?![A-Za-z])|泰铢|฿")
_NUMBER_RE = re.compile(r"[-+]?\d+(?:,\d{3})*(?:\.\d+)?")
_YM_RE = re.compile(r"(20\d{2})\D*(\d{1,2})")
_INLINE_RATE_RE = re.compile(
    r"1\s*(CNY|RMB|USD)\s*=\s*(\d+(?:\.\d+)?)\s*(THB)?",
    re.IGNORECASE,
)
_LOOSE_RATE_RE = re.compile(r"汇率\s*[:：]\s*(\d+(?:\.\d+)?)")
_CURRENCY_LABELS = {
    "CNY": "CNY",
    "RMB": "CNY",
    "人民币": "CNY",
    "USD": "USD",
    "美元": "USD",
    "THB": "THB",
    "泰铢": "THB",
}


@dataclass
class FxRates:
    """年月 → 汇率。cny_thb：1 CNY = THB；usd_thb：1 USD = THB。"""

    cny_thb: Dict[str, float] = field(default_factory=dict)
    usd_thb: Dict[str, float] = field(default_factory=dict)


@dataclass
class FxConvertStats:
    """外币标注与换算计数。"""

    marked_cny: int = 0
    converted_cny: int = 0
    marked_usd: int = 0
    converted_usd: int = 0
    marked_withhold: int = 0
    converted_withhold: int = 0
    stamped_bilie: int = 0
    stamped_split: int = 0
    inline_rate: int = 0
    row_rate: int = 0
    missing_months: Dict[str, int] = field(default_factory=dict)

    def add_missing(self, key: str) -> None:
        self.missing_months[key] = self.missing_months.get(key, 0) + 1


def _currency_code(text: str) -> str:
    """'CNY'、'USD'、'THB'，或空字符串。两种外币同时出现时不猜。"""
    folded = "".join(str(text).split())
    if not folded:
        return ""
    mapped = _CURRENCY_LABELS.get(folded) or _CURRENCY_LABELS.get(folded.upper())
    if mapped:
        return mapped
    has_cny = _CNY_RE.search(folded) is not None
    has_usd = _USD_RE.search(folded) is not None
    has_thb = _THB_RE.search(folded) is not None
    found = [code for flag, code in ((has_cny, "CNY"), (has_usd, "USD"), (has_thb, "THB")) if flag]
    if len(found) == 1:
        return found[0]
    return ""


def parse_money_cell(value) -> tuple[float, str]:
    """(金额, 币种)。币种为 'CNY' / 'USD' / ''。空值 → (0, '')。"""
    if value is None or isinstance(value, bool):
        return 0.0, ""
    try:
        if pd.isna(value):
            return 0.0, ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float)):
        return float(value), ""
    try:
        import numpy as np

        if isinstance(value, np.number):
            return float(value), ""
    except Exception:
        pass

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat", "-"}:
        return 0.0, ""
    currency = _currency_code(text)
    cleaned = (
        text.replace(",", "")
        .replace("，", "")
        .replace(" ", "")
        .replace("\u00a0", "")
        .replace("\u3000", "")
    )
    negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negative = True
        cleaned = cleaned[1:-1]
    found = _NUMBER_RE.search(cleaned)
    if not found:
        return 0.0, currency
    number = float(found.group(0).replace(",", ""))
    if negative and number > 0:
        number = -number
    return number, currency


def money_number(value) -> Optional[float]:
    """格子里的原币数字。空单元格 → None。含 RMB 文本也能读出。"""
    amount, _ccy = _source_amount(value)
    return amount


def inline_fx_rate(remark, currency: str) -> Optional[float]:
    """摘要里的 1CNY=4.959THB。"""
    if currency not in {"CNY", "USD"} or remark is None:
        return None
    try:
        if pd.isna(remark):
            return None
    except (TypeError, ValueError):
        pass
    text = str(remark)
    found = _INLINE_RATE_RE.search(text)
    if found:
        code = found.group(1).upper()
        if code == "RMB":
            code = "CNY"
        if code == currency:
            rate = float(found.group(2))
            if rate > 0:
                return rate
    loose = _LOOSE_RATE_RE.search(text)
    if not loose:
        return None
    rate = float(loose.group(1))
    if rate <= 0:
        return None
    if currency == "CNY" and 3.0 <= rate <= 8.0:
        return rate
    if currency == "USD" and 20.0 <= rate <= 45.0:
        return rate
    return None


def year_month_key(year, month) -> Optional[str]:
    """年、月 → 'YYYY-MM'。不合法则 None。"""
    try:
        if pd.isna(year) or pd.isna(month):
            return None
    except (TypeError, ValueError):
        pass
    try:
        y = int(year)
        m = int(month)
    except (TypeError, ValueError):
        return None
    if y < 1900 or not 1 <= m <= 12:
        return None
    return f"{y:04d}-{m:02d}"


def _normalize_ym_key(value) -> Optional[str]:
    if value is None or isinstance(value, bool):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "year") and hasattr(value, "month"):
        try:
            y, m = int(value.year), int(value.month)
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        except (TypeError, ValueError):
            pass
    found = _YM_RE.search(str(value).strip())
    if not found:
        return None
    month = int(found.group(2))
    if not 1 <= month <= 12:
        return None
    return f"{found.group(1)}-{month:02d}"


def fx_workbook_path(paths=None) -> Path:
    """汇率三表.xlsx，与项目基本信息表同一目录。"""
    if paths is None:
        from .config import Paths

        paths = Paths()
    return Path(paths.project_info_xlsx).parent / "汇率三表.xlsx"


def _rate_table(table: pd.DataFrame, ym_col, rate_col) -> Dict[str, float]:
    rates: Dict[str, float] = {}
    if rate_col is None:
        return rates
    for _, row in table.iterrows():
        key = _normalize_ym_key(row[ym_col])
        if not key:
            continue
        rate = pd.to_numeric(row[rate_col], errors="coerce")
        if pd.isna(rate) or float(rate) <= 0:
            continue
        rates[key] = float(rate)
    return rates


def load_fx_rates(path: Path) -> FxRates:
    """汇率表：人民币对泰铢、美元对泰铢。"""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)
    book = pd.ExcelFile(path)
    sheet = "汇率表" if "汇率表" in book.sheet_names else book.sheet_names[0]
    table = pd.read_excel(book, sheet_name=sheet)
    book.close()

    ym_col = None
    cny_col = None
    usd_col = None
    for col in table.columns:
        name = str(col).replace("\n", "").replace(" ", "")
        if ym_col is None and "年月" in name:
            ym_col = col
        if cny_col is None and "人民币对泰铢" in name:
            cny_col = col
        if usd_col is None and "美元对泰铢" in name:
            usd_col = col
    if ym_col is None or cny_col is None or usd_col is None:
        raise ValueError(f"汇率表缺少年月、人民币对泰铢或美元对泰铢列: {list(table.columns)}")
    return FxRates(
        cny_thb=_rate_table(table, ym_col, cny_col),
        usd_thb=_rate_table(table, ym_col, usd_col),
    )


def load_cny_thb_rates(path: Path) -> Dict[str, float]:
    """汇率表：年月 → 人民币对泰铢（1 CNY = 该数 THB）。"""
    return load_fx_rates(path).cny_thb


def format_year_month(value) -> Optional[str]:
    """归一成 '2026-09'。空值或无法识别则 None。"""
    return _normalize_ym_key(value)


def _is_blank_money(value) -> bool:
    if value is None or isinstance(value, bool):
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        return False
    if isinstance(value, (int, float)):
        return False
    try:
        import numpy as np

        if isinstance(value, np.number):
            return False
    except Exception:
        pass
    text = str(value).strip()
    return (not text) or text.lower() in {"nan", "none", "nat", "-"}


def _source_amount(value) -> Tuple[Optional[float], str]:
    """原始金额（不乘汇率）。空单元格 → (None, '')。"""
    if _is_blank_money(value):
        return None, ""
    amount, currency = parse_money_cell(value)
    return amount, currency


def _optional_positive(value) -> Optional[float]:
    if _is_blank_money(value):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = pd.to_numeric(value, errors="coerce")
    try:
        if pd.isna(number) or float(number) <= 0:
            return None
    except (TypeError, ValueError):
        return None
    return float(number)


def _rate_fits(currency: str, rate: Optional[float]) -> Optional[float]:
    """行内「汇率」只在量级对得上该币种时才用，避免把人民币汇率套到美元。"""
    if rate is None or rate <= 0:
        return None
    if currency == "CNY" and 3.0 <= rate <= 8.0:
        return rate
    if currency == "USD" and 20.0 <= rate <= 45.0:
        return rate
    return None


def _lookup_rate(
    currency: str,
    ym: Optional[str],
    remark,
    fx: FxRates,
    stats: FxConvertStats,
    preferred: Optional[float] = None,
) -> Optional[float]:
    inline = inline_fx_rate(remark, currency)
    if inline:
        stats.inline_rate += 1
        return inline
    fitted = _rate_fits(currency, preferred)
    if fitted:
        stats.row_rate += 1
        return fitted
    table = fx.cny_thb if currency == "CNY" else fx.usd_thb
    rate = table.get(ym) if ym else None
    if rate is not None and rate > 0:
        return rate
    return None


def _convert_part(
    amount: float,
    currency: str,
    ym: Optional[str],
    remark,
    fx: FxRates,
    stats: FxConvertStats,
    *,
    kind: str,
    preferred: Optional[float] = None,
) -> Tuple[Optional[float], Optional[float], bool]:
    """原币 → 泰铢。返回 (泰铢, 汇率, 是否缺汇率)。"""
    table_label = "人民币" if currency == "CNY" else "美元"
    if kind == "actual":
        if currency == "CNY":
            stats.marked_cny += 1
        else:
            stats.marked_usd += 1
    elif kind == "withhold":
        stats.marked_withhold += 1
    rate = _lookup_rate(currency, ym, remark, fx, stats, preferred=preferred)
    if rate is None or rate <= 0:
        stats.add_missing(f"{ym or '无年月'} {table_label}")
        return None, None, True
    if kind == "actual":
        if currency == "CNY":
            stats.converted_cny += 1
        else:
            stats.converted_usd += 1
    elif kind == "withhold":
        stats.converted_withhold += 1
    return amount * rate, rate, False


def column_thb_amounts(
    df: pd.DataFrame,
    column: str,
    *,
    kind: str,
    fx: Optional[FxRates] = None,
) -> pd.Series:
    """一列原币金额换成泰铢。未标币种或泰铢按原数。外币汇率与付款金额相同。"""
    fx = fx or FxRates()
    stats = FxConvertStats()
    if df is None or df.empty:
        return pd.Series(dtype="float64")
    n = len(df)
    parsed = [
        _source_amount(v)
        for v in (df[column].tolist() if column in df.columns else [None] * n)
    ]
    header = CURRENCY_COLUMNS.get(column)
    split_codes = [None] * n
    if header and header in df.columns:
        split_codes = df[header].tolist()
    stamps = [
        _currency_code(v) if not _is_blank_money(v) else ""
        for v in (df[BILIE_COL].tolist() if BILIE_COL in df.columns else [None] * n)
    ]
    remarks = df["摘要"].tolist() if "摘要" in df.columns else [None] * n
    years = df["年"].tolist() if "年" in df.columns else [None] * n
    months = df["月"].tolist() if "月" in df.columns else [None] * n
    row_rates = df[FX_RATE_COL].tolist() if FX_RATE_COL in df.columns else [None] * n
    amounts: list[float] = []
    for index, (amount, embedded) in enumerate(parsed):
        currency = embedded or _currency_code(split_codes[index])
        if not currency and stamps[index] in {"CNY", "USD"}:
            currency = stamps[index]
        if amount is None:
            amounts.append(0.0)
            continue
        if currency in {"CNY", "USD"}:
            thb, _rate, missing = _convert_part(
                amount,
                currency,
                year_month_key(years[index], months[index]),
                remarks[index],
                fx,
                stats,
                kind=kind,
                preferred=_optional_positive(row_rates[index]),
            )
            amounts.append(0.0 if missing or thb is None else round(float(thb), 2))
        else:
            amounts.append(round(float(amount), 2))
    return pd.Series(amounts, index=df.index, dtype="float64")


def collapse_to_payment_amount(
    df: pd.DataFrame,
    fx: FxRates,
) -> tuple[pd.DataFrame, FxConvertStats]:
    """保留五列原始金额，并写上币种、付款金额、汇率、年月。

    五列是去掉币种标注后的原数字。付款金额是预扣+实付换成泰铢后的合计。
    汇率填在发生了换算的行。年月为 '2026-09'。
    """
    stats = FxConvertStats()
    if df is None:
        return df, stats
    out = df.copy()
    if out.empty:
        out[PAYMENT_AMOUNT_COL] = pd.Series(dtype="float64")
        out[FX_RATE_COL] = pd.Series(dtype="float64")
        out[YEAR_MONTH_COL] = pd.Series(dtype="object")
        out[RATE_MISSING_COL] = pd.Series(dtype="int64")
        for header in CURRENCY_COLUMNS.values():
            out[header] = pd.Series(dtype="object")
        return out, stats

    n = len(out)
    parsed = {
        col: [
            _source_amount(v)
            for v in (out[col].tolist() if col in out.columns else [None] * n)
        ]
        for col in LEGACY_AMOUNT_COLUMNS
    }
    stamps = [
        _currency_code(v) if not _is_blank_money(v) else ""
        for v in (out[BILIE_COL].tolist() if BILIE_COL in out.columns else [None] * n)
    ]
    remarks = out["摘要"].tolist() if "摘要" in out.columns else [None] * n
    years = out["年"].tolist() if "年" in out.columns else [None] * n
    months = out["月"].tolist() if "月" in out.columns else [None] * n

    currencies = {col: [pair[1] for pair in parsed[col]] for col in LEGACY_AMOUNT_COLUMNS}
    # 已拆开的「应付币种 / 实付币种 …」。格子里没有 RMB 字样时以这列为准。
    split_used_rows = set()
    for col in LEGACY_AMOUNT_COLUMNS:
        header = CURRENCY_COLUMNS[col]
        if header not in out.columns:
            continue
        for index, raw_ccy in enumerate(out[header].tolist()):
            if currencies[col][index]:
                continue
            code = _currency_code(raw_ccy)
            if not code:
                continue
            currencies[col][index] = code
            if code in {"CNY", "USD"}:
                split_used_rows.add(index)
    stats.stamped_split = len(split_used_rows)
    row_rate_values = (
        out[FX_RATE_COL].tolist() if FX_RATE_COL in out.columns else [None] * n
    )
    for index, stamp in enumerate(stamps):
        if stamp not in {"CNY", "USD"}:
            continue
        used = False
        for col in LEGACY_AMOUNT_COLUMNS:
            amount, ccy = parsed[col][index]
            if amount is None or ccy:
                continue
            currencies[col][index] = stamp
            used = True
        if used:
            stats.stamped_bilie += 1

    source_pay = (
        [money_number(v) for v in out[PAYMENT_AMOUNT_COL].tolist()]
        if PAYMENT_AMOUNT_COL in out.columns
        else [None] * n
    )
    payments = []
    rate_cells = []
    year_months = []
    missing_flags = []
    for index, (year, month) in enumerate(zip(years, months)):
        ym = year_month_key(year, month)
        year_months.append(ym)
        remark = remarks[index]
        used_rate = None
        missing = False
        parts_thb = []
        for col, kind in (("预扣税额", "withhold"), ("实付金额", "actual")):
            amount = parsed[col][index][0]
            ccy = currencies[col][index]
            if amount is None:
                parts_thb.append(0.0)
                continue
            if ccy in {"CNY", "USD"}:
                thb, rate, miss = _convert_part(
                    amount,
                    ccy,
                    ym,
                    remark,
                    fx,
                    stats,
                    kind=kind,
                    preferred=_optional_positive(row_rate_values[index]),
                )
                if miss:
                    missing = True
                    parts_thb.append(None)
                else:
                    parts_thb.append(thb)
                    if rate is not None:
                        used_rate = rate
            else:
                parts_thb.append(amount)
        if missing:
            payments.append(None)
        else:
            total = round(sum(parts_thb), 2)
            src = source_pay[index]
            # 实付和预扣都空时，用表上已经写好的付款金额。
            if abs(total) <= 1e-9 and src is not None and abs(src) > 1e-9:
                total = round(float(src), 2)
            payments.append(total)
        rate_cells.append(used_rate)
        missing_flags.append(1 if missing else 0)

    for col in LEGACY_AMOUNT_COLUMNS:
        out[col] = [pair[0] for pair in parsed[col]]
        out[CURRENCY_COLUMNS[col]] = [
            ccy if ccy else None for ccy in currencies[col]
        ]
    out[PAYMENT_AMOUNT_COL] = payments
    out[FX_RATE_COL] = rate_cells
    out[YEAR_MONTH_COL] = year_months
    out[RATE_MISSING_COL] = missing_flags
    return out, stats


def describe_fx_stats(stats: FxConvertStats, fx: FxRates) -> str:
    lines = [
        "  付款金额 = 预扣税额(泰铢) + 实付金额(泰铢)",
        (
            f"  实付金额人民币 {stats.marked_cny} 笔，已换算 {stats.converted_cny} 笔；"
            f"美元 {stats.marked_usd} 笔，已换算 {stats.converted_usd} 笔"
        ),
        (
            f"  汇率表 人民币 {len(fx.cny_thb)} 个月，美元 {len(fx.usd_thb)} 个月"
            "（1 单位外币 = 当月泰铢）"
        ),
    ]
    if stats.stamped_split:
        lines.append(f"  按分列币种（应付/预扣/实付/增值税/不含税）换算 {stats.stamped_split} 行")
    if stats.row_rate:
        lines.append(f"  用表内「汇率」列 {stats.row_rate} 笔")
    if stats.stamped_bilie:
        lines.append(f"  按「币别」列补标 CNY/USD {stats.stamped_bilie} 行")
    if stats.inline_rate:
        lines.append(f"  用摘要里的 1CNY=/1USD= 汇率 {stats.inline_rate} 笔")
    if stats.marked_withhold:
        lines.append(
            f"  预扣税额外币 {stats.marked_withhold} 笔，已换算 {stats.converted_withhold} 笔"
        )
    if stats.missing_months:
        detail = ", ".join(
            f"{key}×{count}" for key, count in sorted(stats.missing_months.items())
        )
        lines.append(f"  ⚠️ 缺汇率，这些外币行付款金额留空、不按原币计入泰铢: {detail}")
    return "\n".join(lines)


def same_currency_identity_mask(df: pd.DataFrame) -> pd.Series:
    """应付与预扣+实付可直接比原币的行。跨币种（应付人民币、实付泰铢）不参与。"""
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    pay_ccy = (
        df[CURRENCY_COLUMNS["应付金额"]].fillna("").astype(str).str.upper()
        if CURRENCY_COLUMNS["应付金额"] in df.columns
        else pd.Series("", index=df.index)
    )
    act_ccy = (
        df[CURRENCY_COLUMNS["实付金额"]].fillna("").astype(str).str.upper()
        if CURRENCY_COLUMNS["实付金额"] in df.columns
        else pd.Series("", index=df.index)
    )
    w_ccy = (
        df[CURRENCY_COLUMNS["预扣税额"]].fillna("").astype(str).str.upper()
        if CURRENCY_COLUMNS["预扣税额"] in df.columns
        else pd.Series("", index=df.index)
    )
    pay_ccy = pay_ccy.replace({"NAN": "", "NONE": "", "THB": ""})
    act_ccy = act_ccy.replace({"NAN": "", "NONE": "", "THB": ""})
    w_ccy = w_ccy.replace({"NAN": "", "NONE": "", "THB": ""})
    fx = {"CNY", "USD"}
    mixed = pay_ccy.isin(fx) & ~act_ccy.isin(fx)
    mixed = mixed | (act_ccy.isin(fx) & ~pay_ccy.isin(fx) & (pay_ccy != ""))
    mixed = mixed | (w_ccy.isin(fx) & ~act_ccy.isin(fx) & ~pay_ccy.isin(fx))
    return ~mixed


def verify_collapsed_amounts(df: pd.DataFrame, stats: FxConvertStats) -> None:
    """核对：原币五列还在，付款金额 = 已换算预扣 + 已换算实付。"""
    if df is None or df.empty:
        return
    if PAYMENT_AMOUNT_COL not in df.columns:
        raise ValueError("缺少付款金额列")
    for col in LEGACY_AMOUNT_COLUMNS:
        if col not in df.columns:
            raise ValueError(f"缺少金额列 {col}")
    pay = pd.to_numeric(df[PAYMENT_AMOUNT_COL], errors="coerce")
    missing = int(df[RATE_MISSING_COL].fillna(0).sum()) if RATE_MISSING_COL in df.columns else 0
    converted = stats.converted_cny + stats.converted_usd + stats.converted_withhold
    nonempty_pay = int(pay.notna().sum())
    print(
        f"  完整性: 明细 {len(df)} 行，付款金额有值 {nonempty_pay} 行，"
        f"外币已换算 {converted} 笔，缺汇率 {missing} 行",
        flush=True,
    )


def component_nonzero_mask(df: pd.DataFrame) -> pd.Series:
    """预扣或实付能读出非零数字（含 RMB 文本）的行。"""
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    mask = pd.Series(False, index=df.index)
    for col in PAYMENT_PART_COLUMNS:
        if col not in df.columns:
            continue
        mask = mask | df[col].map(lambda v: abs(parse_money_cell(v)[0]) > 1e-9)
    return mask


def ensure_payment_amount(df: pd.DataFrame) -> pd.DataFrame:
    """读表用：已有付款金额则只转成数字。

    旧表没有该列时，付款金额 = 数字（预扣税额）+ 数字（实付金额），不再做汇率换算。
    """
    if df is None:
        return df
    out = df.copy()
    if PAYMENT_AMOUNT_COL in out.columns:
        out[PAYMENT_AMOUNT_COL] = pd.to_numeric(
            out[PAYMENT_AMOUNT_COL], errors="coerce"
        ).fillna(0.0)
        return out
    withhold = (
        out["预扣税额"].map(lambda v: money_number(v) or 0.0)
        if "预扣税额" in out.columns
        else 0.0
    )
    actual = (
        out["实付金额"].map(lambda v: money_number(v) or 0.0)
        if "实付金额" in out.columns
        else 0.0
    )
    out[PAYMENT_AMOUNT_COL] = withhold + actual
    return out


def payment_total(df: pd.DataFrame) -> float:
    if df is None or df.empty:
        return 0.0
    work = df if PAYMENT_AMOUNT_COL in df.columns else ensure_payment_amount(df)
    if work is None or PAYMENT_AMOUNT_COL not in work.columns:
        return 0.0
    return float(pd.to_numeric(work[PAYMENT_AMOUNT_COL], errors="coerce").fillna(0).sum())


def self_test() -> None:
    samples = [
        ("11095RMB", 11095.0, "CNY"),
        ("RMB2532.39", 2532.39, "CNY"),
        ("938rmb", 938.0, "CNY"),
        ("25美元", 25.0, "USD"),
        ("350CNY", 350.0, "CNY"),
        ("1000THB", 1000.0, "THB"),
        (54365.5, 54365.5, ""),
    ]
    for raw, expect_amount, expect_ccy in samples:
        amount, ccy = parse_money_cell(raw)
        if abs(amount - expect_amount) > 1e-9 or ccy != expect_ccy:
            raise SystemExit(f"解析失败 {raw!r} → {(amount, ccy)}")
    fx = FxRates(cny_thb={"2026-08": 4.892, "2024-11": 4.7138}, usd_thb={"2026-08": 33.2})
    frame = pd.DataFrame(
        {
            "年": [2026, 2026, 2024, 2026],
            "月": [8, 8, 11, 10],
            "摘要": ["配电箱", "伙食费", "钢结构", "内部资金调拨 汇率：1CNY=4.959THB"],
            "应付金额": ["132500RMB", "11095RMB", 3000000, 2238751.73],
            "预扣税额": [None, None, None, None],
            "实付金额": ["132500RMB", 54365.5, 3000000, 2238751.73],
            "增值税额": [None, None, None, None],
            "不含税额": ["132500RMB", 54365.5, 3000000, 2238751.73],
            "币别": [None, "THB", "CNY", "CNY"],
        }
    )
    out, stats = collapse_to_payment_amount(frame, fx)
    if abs(out.loc[0, PAYMENT_AMOUNT_COL] - 132500 * 4.892) > 0.05:
        raise SystemExit(f"整行 RMB 未换成泰铢: {out.loc[0, PAYMENT_AMOUNT_COL]}")
    if out.loc[0, "应付币种"] != "CNY" or out.loc[0, "应付金额"] != 132500:
        raise SystemExit("原币应付未保留")
    if out.loc[1, PAYMENT_AMOUNT_COL] != 54365.5 or out.loc[1, "应付币种"] != "CNY":
        raise SystemExit("应付 RMB、实付泰铢 处理错了")
    if abs(out.loc[2, PAYMENT_AMOUNT_COL] - 3000000 * 4.7138) > 0.05:
        raise SystemExit(f"币别 CNY 未换算: {out.loc[2, PAYMENT_AMOUNT_COL]}")
    if abs(out.loc[3, PAYMENT_AMOUNT_COL] - round(2238751.73 * 4.959, 2)) > 0.05:
        raise SystemExit(f"摘要汇率未用上: {out.loc[3, PAYMENT_AMOUNT_COL]}")
    comparable = same_currency_identity_mask(out)
    if bool(comparable.iloc[1]):
        raise SystemExit("应付人民币、实付泰铢 不应进入原币恒等核对")
    if stats.converted_cny < 1:
        raise SystemExit("换算计数为空")
    # 金额、币种已经拆开的表：数字是原币，币种在「实付币种」，泰铢合计看行内汇率。
    split = pd.DataFrame(
        {
            "年": [2024, 2026],
            "月": [11, 8],
            "摘要": ["钢结构", "泰铢劳务"],
            "应付金额": [3000000, 1000],
            "预扣税额": [None, 30],
            "实付金额": [3000000, 970],
            "增值税额": [None, None],
            "不含税额": [3000000, 970],
            "应付币种": ["CNY", None],
            "预扣币种": [None, None],
            "实付币种": ["CNY", None],
            "增值税币种": [None, None],
            "不含税币种": ["CNY", None],
            "币别": [None, "THB"],
            "汇率": [4.7847, None],
        }
    )
    split_out, split_stats = collapse_to_payment_amount(
        split,
        FxRates(cny_thb={"2024-11": 4.7138}, usd_thb={}),
    )
    if abs(split_out.loc[0, PAYMENT_AMOUNT_COL] - round(3000000 * 4.7847, 2)) > 0.05:
        raise SystemExit(f"分列币种未按行内汇率换算: {split_out.loc[0, PAYMENT_AMOUNT_COL]}")
    if split_out.loc[0, "实付币种"] != "CNY" or split_out.loc[0, "实付金额"] != 3000000:
        raise SystemExit("分列原币金额被改掉了")
    if abs(split_out.loc[1, PAYMENT_AMOUNT_COL] - 1000) > 0.05:
        raise SystemExit(f"币别 THB 被误换算: {split_out.loc[1, PAYMENT_AMOUNT_COL]}")
    if split_stats.stamped_split != 1:
        raise SystemExit(f"分列币种计数不对: {split_stats.stamped_split}")
    print("payment_amount 自检通过", flush=True)


if __name__ == "__main__":
    self_test()
