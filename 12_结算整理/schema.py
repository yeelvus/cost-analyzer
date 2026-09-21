# -*- coding: utf-8 -*-
from collections import OrderedDict
from typing import Any, Mapping

from config import DEFAULT_RATES

JSON_FIELDS = [
    "项目编码",
    "项目名称",
    "结算金额",
    "结算金额单位",
    "结算时间",
    "合同金额",
    "申报金额",
    "已付款金额",
    "余款",
    "质保金",
    "币种",
    "汇率",
    "备注",
]

MONEY_FIELDS = (
    "结算金额",
    "合同金额",
    "申报金额",
    "已付款金额",
    "余款",
    "质保金",
)


def empty_record(code: str = "", name: str = "") -> dict[str, Any]:
    return {
        "项目编码": code,
        "项目名称": name,
        "结算金额": "",
        "结算金额单位": "THB",
        "结算时间": "",
        "合同金额": "",
        "申报金额": "",
        "已付款金额": "",
        "余款": "",
        "质保金": "",
        "币种": "THB",
        "汇率": "",
        "备注": "",
    }


def parse_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def effective_rate(currency: Any, rate: Any) -> float:
    """填写了汇率用填写值；否则泰铢=1，CNY=5，USD=32。"""
    filled = parse_number(rate)
    if filled is not None:
        return filled
    cur = str(currency or "").strip()
    if not cur:
        return 1.0
    if cur in DEFAULT_RATES:
        return float(DEFAULT_RATES[cur])
    return float(DEFAULT_RATES.get(cur.upper(), 1.0))


def to_thb(amount: Any, currency: Any, rate: Any) -> float | None:
    num = parse_number(amount)
    if num is None:
        return None
    return num * effective_rate(currency, rate)


def ordered_record(data: Mapping[str, Any]) -> OrderedDict:
    out: OrderedDict[str, Any] = OrderedDict()
    blank = empty_record()
    for key in JSON_FIELDS:
        out[key] = data.get(key, blank.get(key, ""))
    for key, value in data.items():
        if key not in out:
            out[key] = value
    return out
