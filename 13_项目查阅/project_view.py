# -*- coding: utf-8 -*-
"""按项目编码、项目名称整理 03_Sum 与 02_Std，供查阅。

只读业务数据。结果写到 01成本统计/19_项目查阅，不写回 02_Std、03_Sum、18_output。
项目匹配只用内存中的结果，不保存匹配库。

输出：
  00_按项目总汇总.xlsx
      项目总览：每个项目一行，跨业务的金额指标
      明细行数：该项目各明细表有多少行
  01_按项目/{项目编码}_{项目名称}.xlsx
      本项目、目录，以及该项目在各业务里的明细 sheet
"""

from __future__ import annotations

import re
import shutil
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from common.layout import domain_dir  # noqa: E402
from common.project import ProjectMatcher  # noqa: E402
from settings import DATA_ROOT  # noqa: E402

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# THCC-2024-002、LSBM-01、XCY-02。其他写法不当作正式编码。
FORMAL_CODE_RE = re.compile(r"^(?:THCC-\d{4}-\d{2,3}|LSBM-\d{2}|XCY-\d{2})$", re.IGNORECASE)
FILE_CODE_RE = re.compile(
    r"^((?:THCC-\d{4}-\d{2,3})|(?:LSBM-\d{2})|(?:XCY-\d{2}))(?:_|\.|$)",
    re.IGNORECASE,
)
MONTH_RE = re.compile(r"(20\d{2})\D{0,3}(\d{1,2})")
# 汇总表末行的「合计 / 【总计】 / 【全部项目合计】」，不是项目。
TOTAL_LABEL_RE = re.compile(r"^[【\[]?(?:全部(?:项目)?)?(?:总?合计|总计|小计)[】\]]?$")
BUDGET_NAME_RE = re.compile(r"^标段人材机汇总-(.+)_processed$")
ILLEGAL_FILENAME_RE = re.compile(r'[\\/:*?"<>|]+')
ILLEGAL_SHEET_RE = re.compile(r"[:\\/?*\[\]]+")
EMPTY_TEXT = {"", "nan", "none", "nat", "null", "<na>"}

OUTPUT_DIRNAME = "19_项目查阅"
SUMMARY_NAME = "00_按项目总汇总.xlsx"
PROJECT_DIRNAME = "01_按项目"

# 03_Sum 里要带进项目文件的表。金额指标从这些表汇总到「项目总览」。
# 币种拆分、不分类副本、科目宽表（项目在列上）不重复装入，说明写在「读取说明」。
SUM_TABLES: tuple[dict, ...] = (
    {
        "domain": "成本",
        "domain_key": "成本统计收集",
        "filename": "项目成本统计数据_汇总.xlsx",
        "source_sheet": "成本统计数据",
        "sheet": "成本_月度",
        "metrics": (
            ("成本_统计月份", "月份", "last"),
            ("成本_预计总成本", "预计总成本", "last"),
            ("成本_毛利", "毛利", "last"),
            ("成本_毛利率", "毛利率", "last"),
            ("成本_实际已发生", "实际已发生金额", "last"),
            ("成本_业主预计结算额", "业主预计结算额", "last"),
            ("成本_总包合同额", "总包合同额", "last"),
            ("成本_分包合同额", "分包合同额", "last"),
            ("成本_分包结算额", "分包结算额", "last"),
            ("成本_材料采购", "材料采购实际已发生", "last"),
            ("成本_人工费", "人工费汇总", "last"),
            ("成本_机具费", "机具费汇总", "last"),
        ),
    },
    {
        "domain": "预算",
        "domain_key": "预算部分",
        "filename": "预算人材机_汇总.xlsx",
        "source_sheet": "总汇总",
        "sheet": "预算_汇总",
        "metrics": (
            ("预算_币种", "币种", "join"),
            ("预算_标段人材机汇总", "标段人材机汇总", "sum"),
            ("预算_钢筋", "钢筋", "sum"),
            ("预算_混凝土", "混凝土", "sum"),
        ),
    },
    {
        "domain": "预算",
        "domain_key": "预算部分",
        "filename": "预算人材机_汇总.xlsx",
        "source_sheet": "材料汇总",
        "sheet": "预算_材料",
        "metrics": (),
    },
    {
        "domain": "合同",
        "domain_key": "合同管理",
        "filename": "合同管理_总包分包_汇总.xlsx",
        "source_sheet": "总包合同",
        "sheet": "合同_总包明细",
        "metrics": (),
    },
    {
        "domain": "合同",
        "domain_key": "合同管理",
        "filename": "合同管理_总包分包_汇总.xlsx",
        "source_sheet": "分包合同",
        "sheet": "合同_分包明细",
        "metrics": (),
    },
    {
        "domain": "合同",
        "domain_key": "合同管理",
        "filename": "合同管理_总包分包_汇总.xlsx",
        "source_sheet": "1_总包分包汇总(全部)",
        "sheet": "合同_汇总",
        "metrics": (
            ("合同_总包合同额_泰铢", "总包合同额(泰铢)", "sum"),
            ("合同_分包合同额_泰铢", "分包合同额(泰铢)", "sum"),
            ("合同_台账分包合同额_泰铢", "台账分包合同额(泰铢)", "sum"),
        ),
    },
    {
        "domain": "分包台账",
        "domain_key": "分包合同",
        "filename": "分包合同台账_汇总.xlsx",
        "source_sheet": "项目分包合同汇总",
        "sheet": "分包台账_汇总",
        "metrics": (("分包台账_合同额", "分包合同汇总金额", "sum"),),
    },
    {
        "domain": "分包台账",
        "domain_key": "分包合同",
        "filename": "分包合同台账_汇总.xlsx",
        "source_sheet": "项目年月合同额汇总",
        "sheet": "分包台账_按月",
        "metrics": (),
    },
    {
        "domain": "分包台账",
        "domain_key": "分包合同",
        "filename": "分包合同台账_汇总.xlsx",
        "source_sheet": "项目付款汇总_泰铢",
        "sheet": "分包台账_付款",
        "metrics": (("分包台账_已付款_泰铢", "分包已付款（泰铢）", "sum"),),
    },
    {
        "domain": "分包台账",
        "domain_key": "分包合同",
        "filename": "分包合同台账_汇总.xlsx",
        "source_sheet": "合同变化汇总",
        "sheet": "分包台账_变化",
        "metrics": (),
    },
    {
        "domain": "采购",
        "domain_key": "采购台账",
        "filename": "采购部采购台账_std_sum.xlsx",
        "source_sheet": "合并数据",
        "sheet": "采购_明细",
        "metrics": (),
    },
    {
        "domain": "采购",
        "domain_key": "采购台账",
        "filename": "采购部采购台账_std_sum.xlsx",
        "source_sheet": "项目汇总",
        "sheet": "采购_分类",
        "metrics": (),
    },
    {
        "domain": "采购",
        "domain_key": "采购台账",
        "filename": "采购部采购台账_std_sum.xlsx",
        "source_sheet": "项目金额汇总",
        "sheet": "采购_金额",
        "metrics": (("采购_总金额", "总采购金额", "sum"),),
    },
    {
        "domain": "采购",
        "domain_key": "采购台账",
        "filename": "采购部采购台账_std_sum.xlsx",
        "source_sheet": "采购总金额",
        "sheet": "采购_按月",
        "metrics": (),
    },
    {
        "domain": "机具",
        "domain_key": "机具费",
        "filename": "机具费_汇总.xlsx",
        "source_sheet": "汇总",
        "sheet": "机具_明细",
        "metrics": (),
    },
    {
        "domain": "机具",
        "domain_key": "机具费",
        "filename": "机具费_汇总.xlsx",
        "source_sheet": "项目汇总机具费",
        "sheet": "机具_汇总",
        "metrics": (("机具_费用合计_泰铢", "费用合计(泰铢)", "sum"),),
    },
    {
        "domain": "机具",
        "domain_key": "机具费",
        "filename": "机具费_汇总.xlsx",
        "source_sheet": "项目月份金额宽表",
        "sheet": "机具_按月",
        "metrics": (),
    },
    {
        "domain": "人工",
        "domain_key": "人工费",
        "filename": "人工费_汇总.xlsx",
        "source_sheet": "原始数据",
        "sheet": "人工_月度明细",
        "metrics": (),
    },
    {
        "domain": "人工",
        "domain_key": "人工费",
        "filename": "人工费_汇总.xlsx",
        "source_sheet": "按项目汇总",
        "sheet": "人工_汇总",
        "metrics": (("人工_总人工费", "总人工费", "sum"),),
    },
    {
        "domain": "人工",
        "domain_key": "人工费",
        "filename": "人工费_汇总.xlsx",
        "source_sheet": "总人工费",
        "sheet": "人工_按月",
        "metrics": (),
    },
    {
        "domain": "机票",
        "domain_key": "人工费",
        "filename": "机票费_汇总.xlsx",
        "source_sheet": "基础汇总",
        "sheet": "机票_明细",
        "metrics": (),
    },
    {
        "domain": "机票",
        "domain_key": "人工费",
        "filename": "机票费_汇总.xlsx",
        "source_sheet": "按项目汇总",
        "sheet": "机票_汇总",
        "metrics": (("机票_费用合计_泰铢", "机票费用合计（泰铢）", "sum"),),
    },
    {
        "domain": "机票",
        "domain_key": "人工费",
        "filename": "机票费_汇总.xlsx",
        "source_sheet": "月度透视",
        "sheet": "机票_按月",
        "metrics": (),
    },
    {
        "domain": "财务",
        "domain_key": "财务付款",
        "filename": "流水_财务付款数据_汇总.xlsx",
        "source_sheet": "总体汇总",
        "sheet": "财务_汇总",
        "metrics": (
            ("财务_付款金额", "付款金额", "sum"),
            ("财务_计入成本", "计入成本_付款金额", "sum"),
        ),
    },
    {
        "domain": "财务",
        "domain_key": "财务付款",
        "filename": "流水_财务付款数据_汇总.xlsx",
        "source_sheet": "项目按月宽表",
        "sheet": "财务_按月",
        "metrics": (),
    },
    {
        "domain": "财务",
        "domain_key": "财务付款",
        "filename": "流水_财务付款数据_汇总.xlsx",
        "source_sheet": "项目一级科目汇总",
        "sheet": "财务_科目",
        "metrics": (),
    },
    {
        "domain": "结算",
        "domain_key": "业主结算额",
        "filename": "业主结算额_汇总.xlsx",
        "source_sheet": "业主结算",
        "sheet": "结算_明细",
        "metrics": (
            ("结算_金额_泰铢", "结算金额_泰铢", "sum"),
            ("结算_已付款_泰铢", "已付款金额_泰铢", "sum"),
        ),
    },
    {
        "domain": "责任成本",
        "domain_key": "责任成本",
        "filename": "责任成本_汇总.xlsx",
        "source_sheet": "责任成本汇总",
        "sheet": "责任成本_汇总",
        "metrics": (
            ("责任成本_合计", "合计金额", "sum"),
            ("责任成本_合同金额", "合同金额汇总", "sum"),
        ),
    },
    {
        "domain": "商混",
        "domain_key": "商混站",
        "filename": "搅拌站混凝土_汇总.xlsx",
        "source_sheet": "计算明细(全)",
        "sheet": "商混_明细",
        "metrics": (),
    },
    {
        "domain": "商混",
        "domain_key": "商混站",
        "filename": "搅拌站混凝土_汇总.xlsx",
        "source_sheet": "客户强度明细汇总",
        "sheet": "商混_强度",
        "metrics": (),
    },
    {
        "domain": "商混",
        "domain_key": "商混站",
        "filename": "搅拌站混凝土_汇总.xlsx",
        "source_sheet": "客户总额汇总",
        "sheet": "商混_汇总",
        "metrics": (
            ("商混_总方量", "总方量", "sum"),
            ("商混_金额", "金额", "sum"),
        ),
    },
    {
        "domain": "商混",
        "domain_key": "商混站",
        "filename": "搅拌站混凝土_汇总.xlsx",
        "source_sheet": "月度方量汇总",
        "sheet": "商混_方量按月",
        "metrics": (),
    },
    {
        "domain": "商混",
        "domain_key": "商混站",
        "filename": "搅拌站混凝土_汇总.xlsx",
        "source_sheet": "月度金额汇总",
        "sheet": "商混_金额按月",
        "metrics": (),
    },
    {
        "domain": "欠款",
        "domain_key": "材料欠款",
        "filename": "材料采购欠款_汇总.xlsx",
        "source_sheet": "欠款汇总",
        "sheet": "欠款",
        "metrics": (),
    },
)

NOTE_LINES = [
    "本文件由 13_项目查阅/00_按项目查阅汇总.py 生成，只读 03_Sum 和 02_Std。",
    "不修改 02_Std、03_Sum、18_output，也不保存项目匹配库。",
    "分组规则：行上已有正式项目编码（THCC / LSBM / XCY）时，按编码归入，不再按名称改编码。",
    "没有正式编码时，用现有项目匹配库把名称对到编码；对不上的留在「未编码」，项目名称保持原值。",
    "项目总览的名称、状态、项目经理来自 0_项目基本信息情况/项目基本信息表.xlsx 的「在建项目」。",
    "成本_* 金额沿用成本统计汇总原表，约等于万泰铢，不能与合同、采购、机具、人工、财务的泰铢直接相加。",
    "其余金额沿用各自汇总表的单位。合同、机具、分包台账付款、结算、机票列名里已标明泰铢。",
    "成本指标取该项目最新统计月份的一行，避免把多个月加总。其他金额按项目求和。",
    "项目文件里的明细保留来源表的原始项目名称，方便和原表对照。",
    "02_Std 只补充 03_Sum 里没有的逐行明细：预算人材机分表、责任成本构成、财务流水明细、分包签订（每个项目取最新月份）。",
    "不装入的内容：采购/机具/人工的中间拆分文件（与 03_Sum 明细重复），财务学习池，成本月报 xlsm 原件，",
    "合同的币种拆分表和不分类副本，财务「一级二级科目宽表」（项目在列上，与财务_科目重复），机具处理报告。",
    "分包签订明细是该项目在 02_Std/05_分包合同 里出现过的最新一个月，不是把每月快照叠在一起。",
    "来源表最后的合计、总计、小计、【全部项目合计】整行跳过，避免和各项目相加时重复计算。",
    "重新运行会整目录刷新 19_项目查阅。",
]


@dataclass
class LoadedTable:
    spec: dict
    frame: pd.DataFrame
    source_file: str


@dataclass
class ProjectBook:
    key: str
    code: str
    name: str
    full_name: str = ""
    status: str = ""
    manager: str = ""
    metrics: dict = field(default_factory=dict)
    counts: dict = field(default_factory=dict)
    budget_files: list = field(default_factory=list)
    responsibility_files: list = field(default_factory=list)
    finance_files: list = field(default_factory=list)
    filename: str = ""


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    text = str(value).replace("\x00", "").strip()
    if text.lower() in EMPTY_TEXT:
        return ""
    return text


def clean_code(value) -> str:
    return clean_text(value).replace(" ", "")


def is_total_label(value) -> bool:
    text = clean_text(value).replace(" ", "")
    return bool(text) and bool(TOTAL_LABEL_RE.match(text))


def formal_code(value) -> str:
    code = clean_code(value).upper()
    if FORMAL_CODE_RE.match(code):
        return code
    return ""


def norm_name(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).casefold()


def month_key(value) -> tuple[int, int]:
    match = MONTH_RE.search(clean_text(value))
    if not match:
        return (0, 0)
    month = int(match.group(2))
    if not 1 <= month <= 12:
        return (0, 0)
    return (int(match.group(1)), month)


def code_from_filename(filename: str) -> str:
    match = FILE_CODE_RE.match(Path(filename).name)
    if not match:
        return ""
    return formal_code(match.group(1))


def finance_loose_name(filename: str) -> str:
    stem = Path(filename).stem
    for prefix in ("未识别编码_", "需手动核对_"):
        if stem.startswith(prefix):
            rest = stem[len(prefix):]
            return re.sub(r"_财务付款数据$", "", rest).strip()
    return ""


def budget_loose_name(filename: str) -> str:
    match = BUDGET_NAME_RE.match(Path(filename).stem)
    return match.group(1).strip() if match else ""


def safe_filename(code: str, name: str) -> str:
    if code:
        base = f"{code}_{name or '未命名'}"
    else:
        base = f"未编码_{name or '未命名'}"
    base = ILLEGAL_FILENAME_RE.sub("_", base).strip(" ._")
    return (base or "未命名") + ".xlsx"


def safe_sheet_name(name: str, used: set[str]) -> str:
    base = ILLEGAL_SHEET_RE.sub("_", clean_text(name) or "Sheet")
    base = base[:31] or "Sheet"
    if base not in used:
        used.add(base)
        return base
    number = 2
    while True:
        suffix = f"_{number}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate
        number += 1


def output_dir() -> Path:
    return Path(DATA_ROOT) / OUTPUT_DIRNAME


def assert_output_is_safe(path: Path) -> None:
    resolved = path.resolve()
    banned = [
        domain_dir(DATA_ROOT, "std", "成本统计收集").parent,
        domain_dir(DATA_ROOT, "sum", "成本统计收集").parent,
        Path(DATA_ROOT) / "18_output",
    ]
    for root in banned:
        root_resolved = root.resolve()
        if resolved == root_resolved or root_resolved in resolved.parents:
            raise SystemExit(f"拒绝写入已有数据目录：{resolved}")


def load_master() -> dict[str, dict]:
    path = Path(DATA_ROOT) / "0_项目基本信息情况" / "项目基本信息表.xlsx"
    frame = pd.read_excel(path, sheet_name="在建项目")
    master: dict[str, dict] = {}
    for _, row in frame.iterrows():
        code = formal_code(row.get("项目编码"))
        name = clean_text(row.get("项目名称"))
        if not code or not name:
            continue
        master[code] = {
            "name": name,
            "full": clean_text(row.get("项目全称")),
            "status": clean_text(row.get("项目状态")),
            "manager": clean_text(row.get("项目经理")),
        }
    return master


def make_resolver(matcher: ProjectMatcher | None):
    cache: dict[str, tuple[str, str]] = {}

    def resolve(name: str) -> tuple[str, str]:
        name = clean_text(name)
        if not name:
            return "", ""
        if name in cache:
            return cache[name]
        hit = matcher.match(name, fuzzy=True) if matcher is not None else None
        source = "" if not hit else str(hit.get("source") or "")
        if not hit or source in {"unmatched_keep", "unmatched"}:
            result = ("", name)
        else:
            code = formal_code(hit.get("project_code"))
            standard = clean_text(hit.get("project_name")) or name
            result = (code, standard) if code else ("", name)
        cache[name] = result
        return result

    return resolve


def row_key(code, name, resolve) -> str:
    if is_total_label(code) or is_total_label(name):
        return ""
    formal = formal_code(code)
    if formal:
        return formal
    resolved_code, _resolved_name = resolve(clean_text(name))
    if resolved_code:
        return resolved_code
    plain = clean_text(name)
    if not plain:
        return ""
    return "未编码:" + norm_name(plain)


def assign_keys(frame: pd.DataFrame, resolve) -> pd.DataFrame:
    out = frame.copy()
    code_col = "项目编码" if "项目编码" in out.columns else None
    name_col = "项目名称" if "项目名称" in out.columns else None
    if code_col is None and name_col is None:
        out["_key"] = ""
        return out
    codes = out[code_col] if code_col else pd.Series([""] * len(out), index=out.index)
    names = out[name_col] if name_col else pd.Series([""] * len(out), index=out.index)
    out["_key"] = [row_key(code, name, resolve) for code, name in zip(codes, names)]
    return out


def _sort_for_last(part: pd.DataFrame) -> pd.DataFrame:
    for column in ("月份", "年月", "年-月", "统计月份"):
        if column in part.columns:
            return part.assign(_mk=part[column].map(month_key)).sort_values("_mk", kind="mergesort")
    return part


def aggregate_metrics(frame: pd.DataFrame, metrics: tuple) -> pd.DataFrame:
    if frame.empty or not metrics or "_key" not in frame.columns:
        return pd.DataFrame(columns=["_key", *[item[0] for item in metrics]])
    rows = []
    for key, part in frame.groupby("_key", sort=False):
        if not key:
            continue
        ordered = _sort_for_last(part)
        record = {"_key": key}
        for out_name, source_name, how in metrics:
            if source_name not in part.columns:
                record[out_name] = pd.NA
                continue
            if how == "sum":
                record[out_name] = pd.to_numeric(part[source_name], errors="coerce").sum(min_count=1)
            elif how == "join":
                values = []
                for value in part[source_name].tolist():
                    text = clean_text(value)
                    if text and text not in values:
                        values.append(text)
                record[out_name] = "、".join(values)
            else:
                series = ordered[source_name]
                picked = series.map(lambda value: value if clean_text(value) or _is_number(value) else pd.NA).dropna()
                record[out_name] = picked.iloc[-1] if not picked.empty else pd.NA
        rows.append(record)
    if not rows:
        return pd.DataFrame(columns=["_key", *[item[0] for item in metrics]])
    return pd.DataFrame(rows)


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and not pd.isna(value)


def aggregate_arrears(frame: pd.DataFrame) -> pd.DataFrame:
    """欠款按货币分开。泰铢进金额列，其他货币写成文字，避免把人民币加进泰铢。"""
    columns = ["_key", "欠款_泰铢", "欠款_其他货币"]
    if frame.empty or "_key" not in frame.columns:
        return pd.DataFrame(columns=columns)
    amount_col = "欠款金额（泰铢）"
    if amount_col not in frame.columns:
        return pd.DataFrame(columns=columns)
    rows = []
    for key, part in frame.groupby("_key", sort=False):
        if not key:
            continue
        thb = 0.0
        thb_seen = False
        others: list[str] = []
        for _, row in part.iterrows():
            amount = pd.to_numeric(row.get(amount_col), errors="coerce")
            currency = clean_text(row.get("货币")) or "泰铢"
            if pd.isna(amount):
                continue
            if currency in {"泰铢", "THB", "thb"}:
                thb += float(amount)
                thb_seen = True
            else:
                others.append(f"{currency}:{amount:g}")
        rows.append(
            {
                "_key": key,
                "欠款_泰铢": thb if thb_seen else pd.NA,
                "欠款_其他货币": "；".join(others),
            }
        )
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=columns)


def excel_cell(value):
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", value)
    return value


def prepare_excel_frame(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.drop(columns=["_key", "_月份序"], errors="ignore").copy()
    out.columns = [
        excel_cell(clean_text(column).replace("\n", " ")) or f"列{index}"
        for index, column in enumerate(out.columns, start=1)
    ]
    renamed = []
    seen: dict[str, int] = {}
    for column in out.columns:
        count = seen.get(column, 0) + 1
        seen[column] = count
        renamed.append(column if count == 1 else f"{column}_{count}")
    out.columns = renamed
    for column in out.columns:
        if out[column].dtype == object:
            out[column] = out[column].map(excel_cell)
    return out


def read_table(path: Path, sheet: str) -> pd.DataFrame | None:
    if not path.is_file():
        print(f"缺少文件，跳过：{path}")
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as exc:
        print(f"读取失败，跳过：{path} [{sheet}] {exc}")
        return None


def load_sum_tables(resolve) -> list[LoadedTable]:
    loaded = []
    for spec in SUM_TABLES:
        path = domain_dir(DATA_ROOT, "sum", spec["domain_key"]) / spec["filename"]
        frame = read_table(path, spec["source_sheet"])
        if frame is None:
            continue
        frame = assign_keys(frame, resolve)
        loaded.append(LoadedTable(spec=spec, frame=frame, source_file=str(path)))
        print(f"  {spec['domain']} / {spec['source_sheet']}：{len(frame)} 行")
    return loaded


def keep_latest_month(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "_key" not in frame.columns or "_月份序" not in frame.columns:
        return frame
    latest = frame.groupby("_key")["_月份序"].transform("max")
    return frame.loc[frame["_月份序"].eq(latest)].drop(columns=["_月份序"])


def load_subcontract_lines(resolve) -> pd.DataFrame:
    folder = domain_dir(DATA_ROOT, "std", "分包合同")
    files = []
    if folder.is_dir():
        for path in folder.glob("分包签订合同汇总_*.xlsx"):
            if path.name.startswith("~$"):
                continue
            match = re.search(r"(20\d{2})-(\d{2})", path.name)
            if match:
                files.append((int(match.group(1)) * 100 + int(match.group(2)), path))
    if not files:
        print("  分包签订：没有月度文件")
        return pd.DataFrame()
    frames = []
    for order, path in sorted(files):
        try:
            frame = pd.read_excel(path)
        except Exception as exc:
            print(f"读取失败，跳过：{path} {exc}")
            continue
        frame["_月份序"] = order
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    merged = assign_keys(pd.concat(frames, ignore_index=True), resolve)
    latest = keep_latest_month(merged)
    print(f"  分包签订最新月：{len(latest)} 行，来自 {len(frames)} 个月份文件")
    return latest


def _excel_files(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted(
        path
        for path in folder.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$") and not path.name.startswith(".")
    )


def index_std_files(resolve) -> tuple[dict[str, dict[str, list[Path]]], dict[str, str]]:
    indexed: dict[str, dict[str, list[Path]]] = {}
    hints: dict[str, str] = {}

    def add(kind: str, key: str, path: Path, loose_name: str = "") -> None:
        if not key:
            return
        indexed.setdefault(key, {}).setdefault(kind, []).append(path)
        if loose_name and key.startswith("未编码:"):
            hints.setdefault(key, loose_name)

    budget_dir = domain_dir(DATA_ROOT, "std", "预算部分")
    for path in _excel_files(budget_dir):
        code = code_from_filename(path.name)
        if code:
            add("budget", code, path)
            continue
        loose = budget_loose_name(path.name)
        resolved, _name = resolve(loose)
        add("budget", resolved or (f"未编码:{norm_name(loose)}" if loose else ""), path, loose)

    responsibility_dir = domain_dir(DATA_ROOT, "std", "责任成本")
    for path in _excel_files(responsibility_dir):
        add("responsibility", code_from_filename(path.name), path)

    finance_dir = domain_dir(DATA_ROOT, "std", "财务付款") / "B_财务付款_std" / "流水"
    for path in _excel_files(finance_dir):
        code = code_from_filename(path.name)
        if code:
            add("finance", code, path)
            continue
        loose = finance_loose_name(path.name)
        resolved, _name = resolve(loose)
        add("finance", resolved or (f"未编码:{norm_name(loose)}" if loose else ""), path, loose)

    counts = {kind: sum(len(item.get(kind, [])) for item in indexed.values()) for kind in ("budget", "responsibility", "finance")}
    print(
        f"  02_Std 文件：预算 {counts['budget']}，责任成本 {counts['responsibility']}，财务流水 {counts['finance']}"
    )
    return indexed, hints


def display_name(key: str, master: dict, name_votes: dict[str, dict[str, int]]) -> str:
    if key in master:
        return master[key]["name"]
    votes = name_votes.get(key) or {}
    if votes:
        return max(votes, key=lambda item: (votes[item], len(item)))
    if key.startswith("未编码:"):
        return key.split(":", 1)[1]
    return key


def collect_name_votes(tables: list[LoadedTable], extra_frames: list[pd.DataFrame]) -> dict[str, dict[str, int]]:
    votes: dict[str, dict[str, int]] = {}
    frames = [item.frame for item in tables if "项目名称" in item.frame.columns]
    frames.extend(frame for frame in extra_frames if "项目名称" in frame.columns)
    for frame in frames:
        if "_key" not in frame.columns:
            continue
        grouped = frame.loc[frame["_key"].astype(str).ne(""), ["_key", "项目名称"]]
        for key, name in grouped.itertuples(index=False):
            text = clean_text(name)
            if not text:
                continue
            bucket = votes.setdefault(str(key), {})
            bucket[text] = bucket.get(text, 0) + 1
    return votes


def build_projects(master, tables, arrears, subcontract, std_index, name_votes) -> dict[str, ProjectBook]:
    projects: dict[str, ProjectBook] = {}

    def ensure(key: str) -> ProjectBook | None:
        if not key:
            return None
        book = projects.get(key)
        if book is not None:
            return book
        info = master.get(key, {})
        code = "" if key.startswith("未编码:") else key
        book = ProjectBook(
            key=key,
            code=code,
            name=display_name(key, master, name_votes),
            full_name=info.get("full", ""),
            status=info.get("status", ""),
            manager=info.get("manager", ""),
        )
        projects[key] = book
        return book

    for code in master:
        ensure(code)

    metric_frames = []
    for item in tables:
        metrics = item.spec.get("metrics") or ()
        if metrics:
            metric_frames.append(aggregate_metrics(item.frame, metrics))
        for key, count in item.frame["_key"].value_counts().items():
            book = ensure(str(key))
            if book is not None and int(count):
                book.counts[item.spec["sheet"]] = int(count)
    if not arrears.empty:
        metric_frames.append(arrears)
        for key, count in arrears["_key"].value_counts().items():
            book = ensure(str(key))
            if book is not None:
                book.counts["欠款"] = int(count)
    if not subcontract.empty:
        for key, count in subcontract["_key"].value_counts().items():
            book = ensure(str(key))
            if book is not None and int(count):
                book.counts["分包_签订明细"] = int(count)

    for frame in metric_frames:
        if frame.empty:
            continue
        for _, row in frame.iterrows():
            book = ensure(str(row["_key"]))
            if book is None:
                continue
            for column, value in row.items():
                if column == "_key" or _blank_metric(value):
                    continue
                book.metrics[column] = value

    for key, kinds in std_index.items():
        book = ensure(key)
        if book is None:
            continue
        book.budget_files = list(kinds.get("budget", []))
        book.responsibility_files = list(kinds.get("responsibility", []))
        book.finance_files = list(kinds.get("finance", []))
    return projects


def _blank_metric(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        return False
    return clean_text(value) == "" and not _is_number(value)


def metric_columns(projects: dict[str, ProjectBook]) -> list[str]:
    ordered = []
    for spec in SUM_TABLES:
        for out_name, _source, _how in spec.get("metrics") or ():
            if out_name not in ordered:
                ordered.append(out_name)
    for extra in ("欠款_泰铢", "欠款_其他货币"):
        if extra not in ordered:
            ordered.append(extra)
    present = set()
    for book in projects.values():
        present.update(book.metrics)
    return [column for column in ordered if column in present]


def count_columns(projects: dict[str, ProjectBook]) -> list[str]:
    ordered = [spec["sheet"] for spec in SUM_TABLES]
    for extra in ("分包_签订明细", "财务_流水明细", "责任_明细", "预算_人材机分表"):
        if extra not in ordered:
            ordered.append(extra)
    present = set()
    for book in projects.values():
        present.update(book.counts)
    return [column for column in ordered if column in present]


def domain_count(book: ProjectBook) -> int:
    domains = set()
    sheet_domain = {spec["sheet"]: spec["domain"] for spec in SUM_TABLES}
    sheet_domain["分包_签订明细"] = "分包台账"
    sheet_domain["财务_流水明细"] = "财务"
    sheet_domain["责任_明细"] = "责任成本"
    for sheet in book.counts:
        if sheet in sheet_domain:
            domains.add(sheet_domain[sheet])
    if book.budget_files:
        domains.add("预算")
    return len(domains)


def summary_frames(projects: dict[str, ProjectBook]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    metrics = metric_columns(projects)
    counts = count_columns(projects)
    overview_rows = []
    count_rows = []
    for book in sorted(projects.values(), key=_project_sort_key):
        overview = {
            "项目编码": book.code,
            "项目名称": book.name,
            "项目全称": book.full_name,
            "项目状态": book.status,
            "项目经理": book.manager,
            "有数据的业务数": domain_count(book),
            "项目文件": f"{PROJECT_DIRNAME}/{book.filename}" if book.filename else "",
        }
        for column in metrics:
            overview[column] = book.metrics.get(column, pd.NA)
        overview_rows.append(overview)
        count_row = {"项目编码": book.code, "项目名称": book.name}
        for column in counts:
            count_row[column] = book.counts.get(column, 0)
        count_rows.append(count_row)
    overview = pd.DataFrame(overview_rows)
    counted = pd.DataFrame(count_rows)
    unnamed = overview.loc[overview["项目编码"].astype(str).eq("")].copy() if not overview.empty else overview
    return overview, counted, unnamed


def _project_sort_key(book: ProjectBook) -> tuple:
    return (1 if not book.code else 0, book.code, book.name)


def _unique_filename(book: ProjectBook, used: set[str]) -> str:
    base = safe_filename(book.code, book.name)
    if base not in used:
        used.add(base)
        return base
    stem = base[:-5]
    number = 2
    while True:
        candidate = f"{stem}_{number}.xlsx"
        if candidate not in used:
            used.add(candidate)
            return candidate
        number += 1


def _has_detail(book: ProjectBook) -> bool:
    return bool(book.counts or book.budget_files or book.responsibility_files or book.finance_files)


def iter_project_sheets(book: ProjectBook, tables: list[LoadedTable], subcontract: pd.DataFrame):
    for item in tables:
        part = item.frame.loc[item.frame["_key"].eq(book.key)]
        if part.empty:
            continue
        yield item.spec["sheet"], prepare_excel_frame(part), item.spec["domain"], item.source_file, item.spec["source_sheet"]
    if not subcontract.empty and "_key" in subcontract.columns:
        part = subcontract.loc[subcontract["_key"].eq(book.key)]
        if not part.empty:
            source = str(domain_dir(DATA_ROOT, "std", "分包合同"))
            yield "分包_签订明细", prepare_excel_frame(part), "分包台账", source, "分包签订合同汇总_该项目最新月"


def read_std_sheet(path: Path, sheet: str) -> pd.DataFrame | None:
    try:
        frame = pd.read_excel(path, sheet_name=sheet)
    except Exception as exc:
        print(f"读取失败，跳过：{path} [{sheet}] {exc}")
        return None
    if frame.empty:
        return None
    return prepare_excel_frame(frame)


def append_std_sheets(book: ProjectBook, used: set[str]):
    sheets = []
    for index, path in enumerate(book.budget_files, start=1):
        prefix = "预算_" if len(book.budget_files) == 1 else f"预算{index}_"
        try:
            names = pd.ExcelFile(path).sheet_names
        except Exception as exc:
            print(f"读取失败，跳过：{path} {exc}")
            continue
        for name in names:
            frame = read_std_sheet(path, name)
            if frame is None:
                continue
            title = safe_sheet_name(f"{prefix}{name}", used)
            sheets.append((title, frame, "预算", str(path), name))
    for path in book.responsibility_files:
        try:
            names = pd.ExcelFile(path).sheet_names
        except Exception as exc:
            print(f"读取失败，跳过：{path} {exc}")
            continue
        for name in names:
            frame = read_std_sheet(path, name)
            if frame is None:
                continue
            title = safe_sheet_name("责任_明细" if name == "责任成本" else f"责任_{name}", used)
            sheets.append((title, frame, "责任成本", str(path), name))
    for path in book.finance_files:
        frame = read_std_sheet(path, "明细数据")
        if frame is None:
            continue
        title = safe_sheet_name("财务_流水明细", used)
        sheets.append((title, frame, "财务", str(path), "明细数据"))
    return sheets


def cover_frame(book: ProjectBook, metrics: list[str]) -> pd.DataFrame:
    rows = [
        ("项目编码", book.code),
        ("项目名称", book.name),
        ("项目全称", book.full_name),
        ("项目状态", book.status),
        ("项目经理", book.manager),
    ]
    for column in metrics:
        value = book.metrics.get(column, "")
        rows.append((column, "" if _blank_metric(value) else value))
    return pd.DataFrame(rows, columns=["项目", "内容"])


def polish(worksheet) -> None:
    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions
    bold = Font(bold=True)
    for cell in worksheet[1]:
        cell.font = bold
    for column in worksheet.columns:
        letter = column[0].column_letter
        header = clean_text(column[0].value)
        worksheet.column_dimensions[letter].width = min(32, max(12, len(header) + 4))


def write_workbook(path: Path, sheets: list[tuple[str, pd.DataFrame]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets:
            frame.to_excel(writer, sheet_name=name[:31], index=False)
        for worksheet in writer.book.worksheets:
            polish(worksheet)


def notes_frame() -> pd.DataFrame:
    return pd.DataFrame({"说明": NOTE_LINES})


def write_outputs(projects, tables, subcontract, metrics) -> Path:
    destination = output_dir()
    assert_output_is_safe(destination)
    staging = destination / "_writing"
    if staging.exists():
        shutil.rmtree(staging)
    project_dir = staging / PROJECT_DIRNAME
    project_dir.mkdir(parents=True)
    written = 0
    used_filenames: set[str] = set()
    for book in sorted(projects.values(), key=_project_sort_key):
        if not _has_detail(book):
            continue
        used = {"本项目", "目录"}
        catalog = []
        body = []
        for sheet, frame, domain, source, source_sheet in iter_project_sheets(book, tables, subcontract):
            title = safe_sheet_name(sheet, used)
            body.append((title, frame))
            catalog.append((title, domain, len(frame), source, source_sheet))
        budget_rows = 0
        responsibility_rows = 0
        for title, frame, domain, source, source_sheet in append_std_sheets(book, used):
            body.append((title, frame))
            catalog.append((title, domain, len(frame), source, source_sheet))
            if domain == "预算":
                budget_rows += len(frame)
            elif domain == "责任成本":
                responsibility_rows += len(frame)
            elif domain == "财务":
                book.counts["财务_流水明细"] = book.counts.get("财务_流水明细", 0) + len(frame)
        if budget_rows:
            book.counts["预算_人材机分表"] = budget_rows
        if responsibility_rows:
            book.counts["责任_明细"] = responsibility_rows
        if not body:
            continue
        book.filename = _unique_filename(book, used_filenames)
        catalog_frame = pd.DataFrame(
            [
                {
                    "序号": index,
                    "工作表": title,
                    "业务": domain,
                    "行数": rows,
                    "来源文件": source,
                    "来源工作表": source_sheet,
                }
                for index, (title, domain, rows, source, source_sheet) in enumerate(catalog, start=1)
            ]
        )
        sheets = [("本项目", cover_frame(book, metrics)), ("目录", catalog_frame), *body]
        write_workbook(project_dir / book.filename, sheets)
        written += 1
        if written % 20 == 0:
            print(f"  已写项目文件 {written}")

    overview, counted, unnamed = summary_frames(projects)
    write_workbook(
        staging / SUMMARY_NAME,
        [
            ("项目总览", overview),
            ("明细行数", counted),
            ("未编码", unnamed if not unnamed.empty else pd.DataFrame(columns=list(overview.columns))),
            ("读取说明", notes_frame()),
        ],
    )

    final_projects = destination / PROJECT_DIRNAME
    final_summary = destination / SUMMARY_NAME
    if final_projects.exists():
        shutil.rmtree(final_projects)
    if final_summary.exists():
        final_summary.unlink()
    shutil.move(str(project_dir), str(final_projects))
    shutil.move(str(staging / SUMMARY_NAME), str(final_summary))
    staging.rmdir()
    print(f"项目文件 {written} 个")
    return destination


def main() -> None:
    print("读取项目基本信息…")
    master = load_master()
    print(f"  正式项目 {len(master)} 个")
    print("加载项目匹配库（只读，不保存）…")
    matcher = ProjectMatcher()
    resolve = make_resolver(matcher)
    print("读取 03_Sum…")
    tables = load_sum_tables(resolve)
    arrears_table = next((item for item in tables if item.spec["sheet"] == "欠款"), None)
    arrears = aggregate_arrears(arrears_table.frame) if arrears_table is not None else pd.DataFrame()
    print("读取 02_Std 里汇总表没有的明细…")
    subcontract = load_subcontract_lines(resolve)
    std_index, name_hints = index_std_files(resolve)
    votes = collect_name_votes(tables, [subcontract])
    for key, hinted_name in name_hints.items():
        bucket = votes.setdefault(key, {})
        bucket.setdefault(hinted_name, 1)
    projects = build_projects(master, tables, arrears, subcontract, std_index, votes)
    metrics = metric_columns(projects)
    print(f"写出 {OUTPUT_DIRNAME}…")
    destination = write_outputs(projects, tables, subcontract, metrics)
    coded = sum(1 for book in projects.values() if book.code)
    loose = sum(1 for book in projects.values() if not book.code)
    with_detail = sum(1 for book in projects.values() if _has_detail(book))
    print(f"完成：{destination}")
    print(f"总览 {len(projects)} 行，其中正式编码 {coded}，未编码 {loose}，有明细文件 {with_detail}")


if __name__ == "__main__":
    main()
