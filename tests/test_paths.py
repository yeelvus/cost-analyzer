# -*- coding: utf-8 -*-
"""0_Raw / 0_Std / 0_html / 0_Sum 路径解析：序号前缀不影响匹配。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.paths import (  # noqa: E402
    dir_key,
    find_month_dirs,
    find_named_dir,
    iter_excel_files,
    month_label_from_dirname,
    resolve_under,
    strip_order_prefix,
)


def test_strip_prefix():
    assert strip_order_prefix("0_Raw") == "Raw"
    assert strip_order_prefix("1_2026年08月") == "2026年08月"
    assert strip_order_prefix("3.1_责任成本") == "责任成本"
    assert strip_order_prefix("10_机具费") == "机具费"
    assert strip_order_prefix("12-财务数据") == "财务数据"
    assert strip_order_prefix("Raw") == "Raw"
    print("strip_prefix OK")


def test_dir_key():
    assert dir_key("0_Raw") == "raw"
    assert dir_key("9_采购台账📦") == "采购台账"
    assert dir_key("10_机具费🚜") == dir_key("机具费")
    print("dir_key OK")


def test_find_named_and_resolve(tmp_path: Path):
    raw = tmp_path / "0_Raw"
    (raw / "2_成本统计收集").mkdir(parents=True)
    (raw / "3.1_责任成本").mkdir()
    (raw / "2_成本统计收集" / "1_2026年08月").mkdir()
    (raw / "2_成本统计收集" / "1_2026年08月" / "a.xlsm").write_bytes(b"")

    assert find_named_dir(tmp_path, "Raw") == raw
    assert find_named_dir(raw, "成本统计收集").name == "2_成本统计收集"
    assert find_named_dir(raw, "2_成本统计收集").name == "2_成本统计收集"
    assert find_named_dir(raw, "责任成本").name == "3.1_责任成本"

    resolved = resolve_under(tmp_path, "Raw", "成本统计收集")
    assert resolved.name == "2_成本统计收集"

    months = find_month_dirs(resolved)
    assert months[0][0] == "2026年08月"
    assert month_label_from_dirname("1_2026年08月") == "2026年08月"

    files = iter_excel_files(resolved)
    assert len(files) == 1
    print("find/resolve/month OK")


def test_html_key():
    assert dir_key("0_html") == "html"
    assert dir_key("html") == "html"
    print("html key OK")


def test_sum_key():
    assert dir_key("0_Sum") == "sum"
    assert dir_key("Sum") == "sum"
    print("sum key OK")


def test_std_key():
    assert dir_key("0_Std") == "std"
    assert dir_key("Std") == "std"
    assert dir_key("1_Std") == "std"
    print("std key OK")


if __name__ == "__main__":
    from tempfile import TemporaryDirectory

    test_strip_prefix()
    test_dir_key()
    with TemporaryDirectory() as d:
        test_find_named_and_resolve(Path(d))
    test_html_key()
    test_sum_key()
    test_std_key()
    print("all path tests OK")
