# -*- coding: utf-8 -*-
"""付款金额 = 预扣税额 + 实付金额。

实付金额格子上带 RMB / CNY / 人民币 时，按该行年月的「人民币对泰铢」换成泰铢。
带 USD / 美元 时，按该行年月的「美元对泰铢」换成泰铢。
没有标注的数字原样相加。缺汇率时不编造，原数字保留并计数。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd

PAYMENT_AMOUNT_COL = "付款金额"
FX_RATE_COL = "汇率"
YEAR_MONTH_COL = "年月"
# A1 保留的来源金额列：存原始数字（去掉币种标注，不乘汇率）。
LEGACY_AMOUNT_COLUMNS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]
# 相加的两列。只有实付金额上的人民币、美元标注才乘对应汇率。
PAYMENT_PART_COLUMNS = ["预扣税额", "实付金额"]

_CNY_RE = re.compile(r"(?i)rmb|cny|人民币|￥|¥")
_USD_RE = re.compile(r"(?i)(?<![A-Za-z])usd(?![A-Za-z])|美元|us\$")
_NUMBER_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")
_YM_RE = re.compile(r"(20\d{2})\D*(\d{1,2})")


@dataclass
class FxRates:
    """年月 → 汇率。cny_thb：1 CNY = THB；usd_thb：1 USD = THB。"""

    cny_thb: Dict[str, float] = field(default_factory=dict)
    usd_thb: Dict[str, float] = field(default_factory=dict)


@dataclass
class FxConvertStats:
    """实付金额外币标注的换算计数。"""

    marked_cny: int = 0
    converted_cny: int = 0
    marked_usd: int = 0
    converted_usd: int = 0
    marked_withhold: int = 0
    missing_months: Dict[str, int] = field(default_factory=dict)

    def add_missing(self, key: str) -> None:
        self.missing_months[key] = self.missing_months.get(key, 0) + 1


def _currency_code(text: str) -> str:
    """'CNY'、'USD'，或空字符串。两种标注同时出现时不猜。"""
    has_cny = _CNY_RE.search(text) is not None
    has_usd = _USD_RE.search(text) is not None
    if has_cny and has_usd:
        return ""
    if has_usd:
        return "USD"
    if has_cny:
        return "CNY"
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
    number = float(found.group(0))
    if negative and number > 0:
        number = -number
    return number, currency


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


def _convert_actual(
    amount: float,
    currency: str,
    ym: Optional[str],
    fx: FxRates,
    stats: FxConvertStats,
) -> Tuple[float, Optional[float]]:
    """实付金额按币种换成泰铢。返回 (泰铢金额, 使用的汇率)。"""
    table = fx.cny_thb if currency == "CNY" else fx.usd_thb
    label = "人民币" if currency == "CNY" else "美元"
    if currency == "CNY":
        stats.marked_cny += 1
    else:
        stats.marked_usd += 1
    rate = table.get(ym) if ym else None
    if rate is None or rate <= 0:
        stats.add_missing(f"{ym or '无年月'} {label}")
        return amount, None
    if currency == "CNY":
        stats.converted_cny += 1
    else:
        stats.converted_usd += 1
    return amount * rate, rate


def collapse_to_payment_amount(
    df: pd.DataFrame,
    fx: FxRates,
) -> tuple[pd.DataFrame, FxConvertStats]:
    """保留五列原始金额，并写上付款金额、汇率、年月。

    五列是去掉币种标注后的原数字。汇率只填在实付金额发生了换算的行。
    年月为 '2026-09'。
    """
    stats = FxConvertStats()
    if df is None:
        return df, stats
    out = df.copy()
    if out.empty:
        out[PAYMENT_AMOUNT_COL] = pd.Series(dtype="float64")
        out[FX_RATE_COL] = pd.Series(dtype="float64")
        out[YEAR_MONTH_COL] = pd.Series(dtype="object")
        return out, stats

    n = len(out)
    parsed = {
        col: [
            _source_amount(v)
            for v in (out[col].tolist() if col in out.columns else [None] * n)
        ]
        for col in LEGACY_AMOUNT_COLUMNS
    }
    years = out["年"].tolist() if "年" in out.columns else [None] * n
    months = out["月"].tolist() if "月" in out.columns else [None] * n

    payments = []
    rate_cells = []
    year_months = []
    for index, (year, month) in enumerate(zip(years, months)):
        w_amt, w_ccy = parsed["预扣税额"][index]
        a_amt, a_ccy = parsed["实付金额"][index]
        ym = year_month_key(year, month)
        year_months.append(ym)
        w_num = 0.0 if w_amt is None else w_amt
        a_num = 0.0 if a_amt is None else a_amt
        used_rate = None
        if w_ccy:
            stats.marked_withhold += 1
        if a_ccy in {"CNY", "USD"}:
            a_num, used_rate = _convert_actual(a_num, a_ccy, ym, fx, stats)
        payments.append(w_num + a_num)
        rate_cells.append(used_rate)

    for col in LEGACY_AMOUNT_COLUMNS:
        out[col] = [pair[0] for pair in parsed[col]]
    out[PAYMENT_AMOUNT_COL] = payments
    out[FX_RATE_COL] = rate_cells
    out[YEAR_MONTH_COL] = year_months
    return out, stats


def describe_fx_stats(stats: FxConvertStats, fx: FxRates) -> str:
    lines = [
        "  付款金额 = 预扣税额 + 实付金额",
        (
            f"  实付金额带人民币标注 {stats.marked_cny} 笔，"
            f"已按当月人民币对泰铢换算 {stats.converted_cny} 笔"
        ),
        (
            f"  实付金额带美元标注 {stats.marked_usd} 笔，"
            f"已按当月美元对泰铢换算 {stats.converted_usd} 笔"
        ),
        (
            f"  汇率表 人民币 {len(fx.cny_thb)} 个月，美元 {len(fx.usd_thb)} 个月"
            "（1 单位外币 = 当月泰铢）"
        ),
    ]
    if stats.marked_withhold:
        lines.append(
            f"  预扣税额带人民币或美元标注 {stats.marked_withhold} 笔，"
            "按原数字加入，不另乘汇率"
        )
    if stats.missing_months:
        detail = ", ".join(
            f"{key}×{count}" for key, count in sorted(stats.missing_months.items())
        )
        lines.append(f"  ⚠️ 缺汇率，这些实付标注未换算、仍按原数字计入: {detail}")
    return "\n".join(lines)


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
        pd.to_numeric(out["预扣税额"], errors="coerce").fillna(0.0)
        if "预扣税额" in out.columns
        else 0.0
    )
    actual = (
        pd.to_numeric(out["实付金额"], errors="coerce").fillna(0.0)
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
