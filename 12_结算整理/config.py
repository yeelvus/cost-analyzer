# -*- coding: utf-8 -*-
from pathlib import Path

DEFAULT_SETTLEMENT_ROOT = Path(
    "/Volumes/PC-SunYongtao/工作临时文件存放包/运行考核部/业主结算额"
)
DEFAULT_SETTLEMENT_OUTPUT_DIR = Path(
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/03_Sum/13_业主结算额"
)

PROJECT_CODE_RE = r"^THCC-\d{4}-\d{3}$"
JSON_NAME_TMPL = "{code}.json"
EXCEL_NAME = "业主结算额_汇总.xlsx"

# 汇率：折合泰铢。JSON 里填了「汇率」就用填写值；留空则按币种默认。
DEFAULT_RATES = {
    "THB": 1.0,
    "泰铢": 1.0,
    "CNY": 5.0,
    "人民币": 5.0,
    "RMB": 5.0,
    "USD": 32.0,
    "美元": 32.0,
}
