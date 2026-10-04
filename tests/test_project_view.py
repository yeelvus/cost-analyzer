# -*- coding: utf-8 -*-
"""项目查阅的分组规则。不读业务 Excel。"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PKG = ROOT / "13_项目查阅"
for path in (str(PKG), str(ROOT)):
    if path not in sys.path:
        sys.path.insert(0, path)

import project_view as view  # noqa: E402


class _Matcher:
    def __init__(self, table):
        self.table = table

    def match(self, name, fuzzy=True):
        return self.table.get(name)


def test_total_labels_are_not_projects():
    assert view.is_total_label("合计")
    assert view.is_total_label("【总计】")
    assert view.is_total_label("【全部项目合计】")
    assert view.is_total_label("小计")
    assert not view.is_total_label("晶赛科技")
    resolve = view.make_resolver(_Matcher({}))
    assert view.row_key("", "【合计】", resolve) == ""
    assert view.row_key("THCC-2026-005", "合计", resolve) == ""
    print("total label OK")


def test_formal_code():
    assert view.formal_code(" THCC-2026-005 ") == "THCC-2026-005"
    assert view.formal_code("lsbm-01") == "LSBM-01"
    assert view.formal_code("xcy-02") == "XCY-02"
    assert view.formal_code("需手动核对") == ""
    assert view.formal_code("未编码") == ""
    assert view.formal_code("") == ""
    print("formal_code OK")


def test_filename_code():
    assert view.code_from_filename("THCC-2026-005_晶赛科技_财务付款数据.xlsx") == "THCC-2026-005"
    assert view.code_from_filename("未识别编码_INSEE_财务付款数据.xlsx") == ""
    assert view.finance_loose_name("未识别编码_INSEE项目_财务付款数据.xlsx") == "INSEE项目"
    assert view.finance_loose_name("需手动核对_欣强项目_财务付款数据.xlsx") == "欣强项目"
    assert view.budget_loose_name("标段人材机汇总-伊莱特_processed.xlsx") == "伊莱特"
    print("filename OK")


def test_row_key_keeps_formal_code():
    resolve = view.make_resolver(
        _Matcher({"别名": {"project_name": "晶赛科技", "project_code": "THCC-2026-005", "source": "alias"}})
    )
    assert view.row_key("THCC-2024-002", "北新项目", resolve) == "THCC-2024-002"
    assert view.row_key("", "别名", resolve) == "THCC-2026-005"
    assert view.row_key("", "没有的项目", resolve) == "未编码:" + view.norm_name("没有的项目")
    assert view.row_key("需手动核对", "INSEE项目", resolve) == "未编码:" + view.norm_name("INSEE项目")
    print("row_key OK")


def test_latest_month_and_metrics():
    assert view.month_key("2026年09月") == (2026, 9)
    assert view.month_key("2024-06") == (2024, 6)
    frame = pd.DataFrame(
        [
            {"项目编码": "THCC-2026-005", "项目名称": "晶赛科技", "月份": "2026年08月", "预计总成本": 10},
            {"项目编码": "THCC-2026-005", "项目名称": "晶赛科技", "月份": "2026年09月", "预计总成本": 12},
            {"项目编码": "", "项目名称": "别名", "月份": "2026年01月", "预计总成本": 3},
        ]
    )
    resolve = view.make_resolver(
        _Matcher({"别名": {"project_name": "晶赛科技", "project_code": "THCC-2026-005", "source": "alias"}})
    )
    keyed = view.assign_keys(frame, resolve)
    result = view.aggregate_metrics(keyed, (("成本_预计总成本", "预计总成本", "last"), ("成本_统计月份", "月份", "last")))
    row = result.loc[result["_key"].eq("THCC-2026-005")].iloc[0]
    assert row["成本_预计总成本"] == 12
    assert row["成本_统计月份"] == "2026年09月"
    print("metrics OK")


def test_arrears_keeps_currency_apart():
    frame = pd.DataFrame(
        [
            {"_key": "THCC-2026-005", "货币": "泰铢", "欠款金额（泰铢）": 10},
            {"_key": "THCC-2026-005", "货币": "人民币", "欠款金额（泰铢）": 2},
        ]
    )
    result = view.aggregate_arrears(frame).iloc[0]
    assert result["欠款_泰铢"] == 10
    assert "人民币" in result["欠款_其他货币"]
    print("arrears OK")


def test_sheet_and_file_names():
    used = set()
    assert view.safe_sheet_name("财务_流水明细", used) == "财务_流水明细"
    assert view.safe_sheet_name("财务_流水明细", used).startswith("财务_流水明细_")
    assert "/" not in view.safe_filename("THCC-2023-012", "振威特（土建/装修）")
    assert view.safe_filename("", "INSEE项目").startswith("未编码_")
    print("names OK")


def test_keep_latest_subcontract_month():
    frame = pd.DataFrame(
        [
            {"_key": "THCC-2026-005", "_月份序": 202607, "合同额": 1},
            {"_key": "THCC-2026-005", "_月份序": 202608, "合同额": 2},
            {"_key": "THCC-2024-002", "_月份序": 202607, "合同额": 9},
        ]
    )
    kept = view.keep_latest_month(frame)
    assert set(kept["合同额"]) == {2, 9}
    print("latest month OK")


if __name__ == "__main__":
    test_total_labels_are_not_projects()
    test_formal_code()
    test_filename_code()
    test_row_key_keeps_formal_code()
    test_latest_month_and_metrics()
    test_arrears_keeps_currency_apart()
    test_sheet_and_file_names()
    test_keep_latest_subcontract_month()
    print("all OK")
