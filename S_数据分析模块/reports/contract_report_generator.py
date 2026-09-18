
# --- path bootstrap (S_数据分析模块) ---
import sys
from pathlib import Path as _PathForBootstrap
try:
    from S_数据分析模块._bootstrap import ensure_code_root, ensure_script_dir
    ensure_code_root(_PathForBootstrap(__file__))
    ensure_script_dir(__file__)
except Exception:
    _p = _PathForBootstrap(__file__).resolve()
    for _c in [_p.parent, *_p.parents]:
        if (_c / "settings.py").exists():
            if str(_c) not in sys.path:
                sys.path.insert(0, str(_c))
            break
    if str(_p.parent) not in sys.path:
        sys.path.insert(0, str(_p.parent))
# --- end bootstrap ---

import re
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import warnings

import pandas as pd

from contract_report_pipeline import (
    build_project_summary,
    build_provenance_rows,
    cleanup_legacy_output_files,
    export_index_file,
    export_project_report,
    export_provenance_excel,
    load_responsibility_cost_mapping,
    replicate_outputs,
)
from contract_report_loaders import (
    coerce_float,
    load_budget_summary,
    load_concrete,
    load_contracts_from_folder,
    load_cost_statistics,
    load_equipment,
    load_finance_summary,
    load_labor,
    load_labor_monthly,
    load_material_arrears,
    load_procurement,
    load_project_name_map,
    load_project_manager_map,
    load_subcontract_ledger,
)

warnings.filterwarnings("ignore", category=FutureWarning)

try:
    from settings import DATA_ROOT
    BASE_DIR = Path(DATA_ROOT)
except Exception:
    BASE_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")

INPUT_PATHS = {
    "合同管理": BASE_DIR / "5_合同管理📑",
    "采购明细": BASE_DIR / "0_Sum/9_采购台账/采购部采购台账_std_sum.xlsx",
    "材料欠款": BASE_DIR / "0_Sum/9.1_材料欠款/泰国子公司材料采购欠款汇总_标准化.xlsx",
    "商混混凝土": [BASE_DIR / "0_Sum/9.1_商混站/搅拌站混凝土汇总报表.xlsx"],
    "机具使用费": BASE_DIR / "0_Sum/10_机具费/机具费_汇总.xlsx",
    "人工费汇总": BASE_DIR / "0_Sum/11_人工费/人工费_汇总.xlsx",
    "财务分包": BASE_DIR / "03_Sum/12_财务付款/银行账_财务付款数据_汇总.xlsx",
    "分包合同台账": BASE_DIR / "0_Std/5.1_分包合同/分包签订合同汇总_2026-07.xlsx",
    "成本统计": BASE_DIR / "0_Sum/2_成本统计收集/项目统计数据汇总_全月份.xlsx",
    "项目经理": BASE_DIR / "0_项目基本信息情况/项目基本信息表.xlsx",
    "预算金额": BASE_DIR / "0_Sum/3_预算部分/预算人材机汇总.xlsx",
    "责任成本": BASE_DIR / "0_Std/3.1_责任成本",
}

FORCED_EQUIPMENT_FILE = None

OUTPUT_PATH = BASE_DIR / "18_output"
PROJECT_REPORTS_DIRNAME = "项目合同汇总"
PROJECT_MANAGER_DIRNAME = "项目经理"
PROJECT_STATUS_DIRNAME = "项目状态"


def calculate_responsibility_cost_summary(responsibility_data) -> Tuple[float, str]:
    """兼容旧调用：供 config_driven_report_builder 动态导入。"""
    if not responsibility_data:
        return 0.0, ""

    total_amount = 0.0
    currency = ""

    for _, df in responsibility_data.items():
        if df.empty:
            continue

        amount_col = None
        curr_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if "金额" in str(col):
                if amount_col is None:
                    amount_col = col
                if "泰铢" in str(col) or "thb" in col_lower:
                    amount_col = col
                    curr_col = "THB"
                    break
                if "人民币" in str(col) or "cny" in col_lower or "rmb" in col_lower:
                    amount_col = col
                    curr_col = "CNY"

        if amount_col is None:
            continue

        try:
            col_sum = df[amount_col].apply(coerce_float).sum()
            total_amount += col_sum
            if curr_col is None:
                curr_col = "CNY" if "人民币" in str(amount_col) else "THB"
            if not currency:
                currency = curr_col
        except Exception:
            continue

    return total_amount, currency if currency else "THB"


def detect_latest_year_month(df) -> str:
    if df.empty:
        return ""

    if {"年", "月"}.issubset(df.columns):
        ym_df = df[["年", "月"]].copy()
        ym_df["年"] = ym_df["年"].apply(lambda x: float(x) if str(x).strip() not in {"", "nan", "None"} else float("nan"))
        ym_df["月"] = ym_df["月"].apply(lambda x: float(x) if str(x).strip() not in {"", "nan", "None"} else float("nan"))
        ym_df = ym_df.dropna(subset=["年", "月"])
        if not ym_df.empty:
            ym_df = ym_df.sort_values(["年", "月"], ascending=[False, False])
            latest = ym_df.iloc[0]
            return f"{int(latest['年'])}年{int(latest['月']):02d}月"

    if "年-月" in df.columns:
        values = df["年-月"].dropna().astype(str)
        pairs = []
        for value in values:
            match = re.search(r"(\d{4})\D?(\d{1,2})", value)
            if match:
                pairs.append((int(match.group(1)), int(match.group(2))))
        if pairs:
            year, month = max(pairs)
            return f"{year}年{month:02d}月"

    return ""


def pick_latest_file(
    root: Path,
    pattern: str,
    recursive: bool = True,
    exclude_substrings: Optional[List[str]] = None,
) -> Optional[Path]:
    if not root or not Path(root).exists():
        return None
    root = Path(root)
    iterator = root.rglob(pattern) if recursive and "**" not in pattern else root.glob(pattern)
    exclude = [s.lower() for s in (exclude_substrings or [])]
    files = []
    for p in iterator:
        if not p.is_file() or p.name.startswith("~$") or p.name.startswith("."):
            continue
        haystack = str(p).lower()
        if any(token in haystack for token in exclude):
            continue
        files.append(p)
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


def pick_latest_existing(
    candidates: List[Tuple[Path, str, bool]],
    exclude_substrings: Optional[List[str]] = None,
) -> Optional[Path]:
    for root, pattern, recursive in candidates:
        latest = pick_latest_file(root, pattern, recursive=recursive, exclude_substrings=exclude_substrings)
        if latest:
            return latest
    return None


def _first_existing_dir(candidates: List[Path]) -> Optional[Path]:
    for path in candidates:
        if path and path.is_dir():
            return path
    return None


def _source_label(value) -> str:
    if isinstance(value, list):
        names = [p.name for p in value if isinstance(p, Path)]
        return "; ".join(names) if names else "未知"
    if isinstance(value, Path):
        return value.name
    return "未知"


def resolve_input_paths() -> Dict:
    resolved: Dict = {}
    for key, value in INPUT_PATHS.items():
        if isinstance(value, list):
            existing = [p for p in value if isinstance(p, Path) and p.exists()]
            if existing:
                resolved[key] = existing
        elif isinstance(value, Path) and (value.exists() or key in {"合同管理", "责任成本"}):
            resolved[key] = value

    contract_dir = resolved.get("合同管理") or INPUT_PATHS["合同管理"]
    contract_latest = pick_latest_existing([
        (contract_dir, "合同管理_总包分包全量汇总.xlsx", False),
        (contract_dir, "合同管理_总包分包*.xlsx", False),
        (contract_dir, "总包合同+分包合同汇总_*.xlsx", True),
        (contract_dir, "合同管理_总包分包汇总_*.xlsx", True),
    ])
    if contract_latest:
        resolved["合同管理汇总"] = contract_latest

    finance_latest = pick_latest_existing(
        [
            (BASE_DIR / "03_Sum/12_财务付款", "银行账_财务付款数据_汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款", "流水_财务付款数据_汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款", "*财务付款数据_汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款/银行账", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款/流水", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "0_Sum/12_财务付款/银行账", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "0_Sum/12_财务付款/流水", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款", "财务付款数据汇总.xlsx", False),
            (BASE_DIR / "03_Sum/12_财务付款", "*财务付款数据汇总.xlsx", True),
            (BASE_DIR / "12-财务数据💰/A_财务付款数据/C_财务付款_sum/银行账", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "12-财务数据💰/A_财务付款数据/C_财务付款_sum/流水", "*财务付款数据汇总.xlsx", False),
            (BASE_DIR / "12-财务数据💰/A_财务付款数据/C_财务付款_sum", "财务付款数据汇总.xlsx", False),
            (BASE_DIR / "12-财务数据💰/A_财务付款数据/C_财务付款_sum", "*财务付款数据汇总.xlsx", True),
            (BASE_DIR / "12-财务数据💰/A_分包对账与付款/B_分包对账与付款_sum", "*分包台账汇总_*.xlsx", False),
            (BASE_DIR / "12-财务数据💰", "*分包台账汇总_*.xlsx", True),
        ],
        exclude_substrings=["历史未分组", "备份", "修正前"],
    )
    if finance_latest:
        resolved["财务分包"] = finance_latest

    budget_latest = pick_latest_existing([
        (BASE_DIR / "3_预算部分/2_人材机_std", "预算人材机汇总*.xlsx", False),
        (BASE_DIR / "3_预算部分/5.1_主材汇总（广联达软件导出）", "预算人材机汇总*.xlsx", False),
        (BASE_DIR / "3_预算部分", "预算人材机汇总*.xlsx", True),
    ])
    if budget_latest:
        resolved["预算金额"] = budget_latest

    if FORCED_EQUIPMENT_FILE and FORCED_EQUIPMENT_FILE.exists():
        resolved["机具使用费"] = FORCED_EQUIPMENT_FILE
    else:
        equip_latest = pick_latest_existing(
            [
                (BASE_DIR / "10_机具费🚜/D_机具费_std_sum", "机具费汇总_*汇总.xlsx", False),
                (BASE_DIR / "10_机具费🚜/D_机具费_std_sum", "*.xlsx", False),
                (BASE_DIR / "Z_report/机具费分析", "机具租赁统计汇总_*.xlsx", False),
                (BASE_DIR / "10_机具费🚜/F_机具费_report", "机具租赁统计汇总_*.xlsx", False),
                (BASE_DIR / "10_机具费🚜/B_机具费_sum", "*租赁机具台账_汇总.xlsx", False),
                (BASE_DIR / "10_机具费🚜/B_机具费_sum", "*机具台账*_汇总.xlsx", False),
                (BASE_DIR / "10_机具费🚜/B_机具费_sum", "机具费汇总_*.xlsx", False),
                (BASE_DIR / "10_机具费🚜/B_机具费_sum", "*_已汇总_*.xlsx", False),
            ],
            exclude_substrings=["split", "raw_sum"],
        )
        if equip_latest:
            resolved["机具使用费"] = equip_latest
        else:
            resolved.pop("机具使用费", None)

    cost_stats_latest = pick_latest_existing([
        (BASE_DIR / "2_成本统计收集/B_成本统计_sum", "项目统计数据汇总_*.xlsx", False),
        (BASE_DIR / "2_成本统计收集/B_成本统计_sum", "*.xlsx", False),
        (BASE_DIR / "2_成本统计收集", "项目统计数据汇总_*.xlsx", True),
    ])
    if cost_stats_latest:
        resolved["成本统计"] = cost_stats_latest

    labor_default = INPUT_PATHS.get("人工费汇总")
    if isinstance(labor_default, Path) and labor_default.exists():
        resolved["人工费汇总"] = labor_default
    else:
        labor_latest = pick_latest_existing([
            (BASE_DIR / "0_Sum/11_人工费", "人工费_汇总.xlsx", False),
            (BASE_DIR / "11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std", "人工费_汇总.xlsx", False),
            (BASE_DIR / "11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std", "人工费汇总_*.xlsx", False),
            (BASE_DIR / "11_人工费🙎🏻‍♂️/A_1_人工费台账/B_人工费台账_std", "人工费汇总*.xlsx", False),
            (BASE_DIR / "11_人工费🙎🏻‍♂️", "综合汇总_*.xlsx", True),
            (BASE_DIR / "11_人工费🙎🏻‍♂️/子公司人力资源发", "综合汇总.xlsx", True),
            # 兜底：规避 emoji 目录名差异导致的精确路径失效。
            (BASE_DIR, "11_人工费*/**/人工费汇总_*.xlsx", False),
            (BASE_DIR, "11_人工费*/**/人工费汇总*.xlsx", False),
            (BASE_DIR, "11_人工费*/**/综合汇总*.xlsx", False),
            (BASE_DIR, "**/人工费汇总_*.xlsx", False),
            (BASE_DIR, "**/人工费汇总*.xlsx", False),
            (BASE_DIR / "Z_report", "人工费_汇总.xlsx", False),
        ], exclude_substrings=["分析"])
        if labor_latest:
            resolved["人工费汇总"] = labor_latest
        else:
            resolved.pop("人工费汇总", None)

    procurement_default = INPUT_PATHS.get("采购明细")
    proc_latest = pick_latest_existing(
        [
            (BASE_DIR / "0_Sum/9_采购台账", "采购部采购台账_std_sum.xlsx", False),
            (BASE_DIR / "0_Sum/9_采购台账", "*_std_sum.xlsx", False),
            (BASE_DIR / "9_采购台账📦/E_采购部采购台账_sum", "采购部采购台账_std_sum.xlsx", False),
            (BASE_DIR / "9_采购台账📦/E_采购部采购台账_sum", "*_std_sum.xlsx", False),
            (BASE_DIR / "9_采购台账📦/E_采购部采购台账_sum", "*_采购台账_std_merged.xlsx", False),
            (BASE_DIR / "9_采购台账📦/E_采购部采购台账_sum", "*.xlsx", False),
            (BASE_DIR / "9_采购台账📦", "*_采购台账_std_merged.xlsx", True),
        ],
        exclude_substrings=["~$"],
    )
    if proc_latest:
        resolved["采购明细"] = proc_latest
    elif isinstance(procurement_default, Path) and procurement_default.exists():
        resolved["采购明细"] = procurement_default

    arrears_default = INPUT_PATHS.get("材料欠款")
    arrears_latest = pick_latest_existing([
        (BASE_DIR / "0_Sum/9.1_材料欠款", "*欠款汇总*.xlsx", False),
        (BASE_DIR / "9.1_材料欠款/B_材料欠款_std", "*材料采购欠款汇总*_标准化*.xlsx", False),
        (BASE_DIR / "9.1_材料欠款/B_材料欠款_std", "*欠款汇总*.xlsx", False),
        (BASE_DIR / "9.1_材料欠款", "*材料采购欠款汇总*_标准化*.xlsx", True),
        (BASE_DIR / "9.1_材料欠款", "*欠款汇总*.xlsx", True),
    ])
    if arrears_latest:
        resolved["材料欠款"] = arrears_latest
    elif isinstance(arrears_default, Path) and arrears_default.exists():
        resolved["材料欠款"] = arrears_default

    concrete_default = INPUT_PATHS.get("商混混凝土", [])
    concrete_default = [p for p in concrete_default if isinstance(p, Path) and p.exists()]
    if concrete_default:
        resolved["商混混凝土"] = concrete_default
    else:
        concrete_latest = pick_latest_existing(
            [
                (BASE_DIR / "0_Sum/9.1_商混站", "搅拌站混凝土汇总报表.xlsx", False),
                (BASE_DIR / "9.1_商混站🚛/B_商混站_std", "搅拌站混凝土汇总报表.xlsx", False),
                (BASE_DIR / "9.1_商混站🚛/B_商混站_std", "搅拌站混凝土汇总报表_*.xlsx", False),
                (BASE_DIR / "9.1_商混站🚛/B_商混站_std", "*.xlsx", False),
                (BASE_DIR / "9.1_商混站🚛", "混凝土汇总报表_*.xlsx", True),
                (BASE_DIR / "9.1_商混站🚛", "*混凝土*汇总*.xlsx", True),
            ],
            exclude_substrings=["仅有"],
        )
        if concrete_latest:
            resolved["商混混凝土"] = [concrete_latest]
        else:
            resolved["商混混凝土"] = []

    sub_contract_latest = pick_latest_existing(
        [
            (BASE_DIR / "0_Std/5.1_分包合同", "分包签订合同汇总_*.xlsx", False),
            (BASE_DIR / "0_Sum/5.1_分包合同", "分包合同台账_汇总.xlsx", False),
            (BASE_DIR / "5.1_分包合同/2_分包合同台账_std", "分包签订合同汇总_*.xlsx", False),
            (BASE_DIR / "5.1_分包合同/2_分包合同台账_std", "*.xlsx", False),
            (BASE_DIR / "5.1_分包合同/3_分包合同_sum", "分包签订合同汇总_*.xlsx", False),
        ],
        exclude_substrings=["钢结构"],
    )
    if sub_contract_latest:
        resolved["分包合同台账"] = sub_contract_latest

    resp_dir = _first_existing_dir([
        BASE_DIR / "0_Std/3.1_责任成本",
        BASE_DIR / "3.1_责任成本/2_责任成本_std",
        BASE_DIR / "3_预算部分/6.1_责任成本书/汇总版",
        BASE_DIR / "3.1_责任成本/1_责任成本_raw",
    ])
    if resp_dir:
        resolved["责任成本"] = resp_dir
    else:
        resolved.pop("责任成本", None)

    return resolved


def _collect_year_month_candidates(text: str) -> List[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    if not text:
        return candidates

    for year, month in re.findall(r"(20\d{2})\D{0,2}(1[0-2]|0?[1-9])", text):
        candidates.append((int(year), int(month)))

    for year, month in re.findall(r"(20\d{2})(0[1-9]|1[0-2])\d{0,2}", text):
        candidates.append((int(year), int(month)))

    return candidates


def infer_latest_year_month_from_texts(texts: List[str]) -> str:
    all_candidates: List[Tuple[int, int]] = []
    for text in texts:
        all_candidates.extend(_collect_year_month_candidates(str(text)))

    if not all_candidates:
        return ""

    year, month = max(all_candidates)
    return f"{year}年{month:02d}月"


def infer_cutoff_from_path(path: Path) -> str:
    if not path:
        return ""
    return infer_latest_year_month_from_texts([path.name, str(path)])


def infer_cutoff_from_paths(paths: List[Path]) -> str:
    if not paths:
        return ""
    texts: List[str] = []
    for p in paths:
        texts.append(str(p.name))
        texts.append(str(p))
    return infer_latest_year_month_from_texts(texts)


def infer_cutoff_from_month_column(df, col_name: str = "月份") -> str:
    if df.empty or col_name not in df.columns:
        return ""
    values = df[col_name].dropna().astype(str).tolist()
    return infer_latest_year_month_from_texts(values)


def _safe_load(label: str, loader, *args, default=None):
    try:
        if args and args[0] is None:
            print(f"⚠ 跳过{label}：未解析到有效文件")
            return default
        if args and isinstance(args[0], Path) and not args[0].exists():
            print(f"⚠ 跳过{label}：文件不存在 {args[0]}")
            return default
        result = loader(*args)
        print(f"✓ {label} 已加载")
        return result
    except Exception as exc:
        print(f"⚠ {label} 读取失败，已跳过：{exc}")
        traceback.print_exc()
        return default


def run_cli():
    resolved_inputs = resolve_input_paths()
    default_contract_dir = resolved_inputs.get("合同管理") or INPUT_PATHS["合同管理"]
    default_procurement = resolved_inputs.get("采购明细")
    default_material_arrears = resolved_inputs.get("材料欠款")
    default_concrete = resolved_inputs.get("商混混凝土") or []
    default_equipment = resolved_inputs.get("机具使用费")
    default_labor = resolved_inputs.get("人工费汇总")
    default_finance = resolved_inputs.get("财务分包")
    default_sub_contract = resolved_inputs.get("分包合同台账")
    default_cost_stats = resolved_inputs.get("成本统计")
    default_manager_map = resolved_inputs.get("项目经理")
    default_budget = resolved_inputs.get("预算金额")
    default_responsibility_cost_dir = resolved_inputs.get("责任成本")
    default_output = OUTPUT_PATH

    if not isinstance(default_labor, Path) or not default_labor.exists():
        print("⚠ 未找到人工费汇总文件，将按空表继续。请检查 11_人工费*/A_1_人工费台账/B_人工费台账_std/。")
        default_labor = None

    print("\n=== 当前数据源路径（自动解析） ===")
    for key, value in resolved_inputs.items():
        if isinstance(value, list):
            for item in value:
                print(f"{key}: {item}")
        else:
            print(f"{key}: {value}")

    missing = [key for key in ("合同管理", "成本统计", "项目经理") if key not in resolved_inputs]
    if missing:
        print(f"⚠ 关键数据源缺失：{', '.join(missing)}")

    cost_stats_cutoff = infer_cutoff_from_path(default_cost_stats) if isinstance(default_cost_stats, Path) else ""
    finance_cutoff = infer_cutoff_from_path(default_finance) if isinstance(default_finance, Path) else ""
    procurement_cutoff = infer_cutoff_from_path(default_procurement) if isinstance(default_procurement, Path) else ""
    material_arrears_cutoff = infer_cutoff_from_path(default_material_arrears) if isinstance(default_material_arrears, Path) else ""
    concrete_cutoff = infer_cutoff_from_paths(default_concrete)
    equipment_cutoff = infer_cutoff_from_path(default_equipment) if isinstance(default_equipment, Path) else ""
    labor_cutoff_by_name = infer_cutoff_from_path(default_labor) if isinstance(default_labor, Path) else ""

    responsibility_candidates = []
    if default_responsibility_cost_dir:
        responsibility_candidates = list(load_responsibility_cost_mapping(default_responsibility_cost_dir).values())
    responsibility_cutoff = infer_cutoff_from_paths(responsibility_candidates)

    data_source_info = {
        "项目成本统计汇总表": {"文件名": _source_label(default_cost_stats), "统计截止日期": cost_stats_cutoff or "未知"},
        "财务分包付款明细": {"文件名": _source_label(default_finance), "统计截止日期": finance_cutoff or "未知"},
        "材料采购明细": {"文件名": _source_label(default_procurement), "统计截止日期": procurement_cutoff or "未知"},
        "材料欠款明细": {"文件名": _source_label(default_material_arrears), "统计截止日期": material_arrears_cutoff or "未知"},
        "商混站明细": {"文件名": _source_label(default_concrete), "统计截止日期": concrete_cutoff or "未知"},
        "机具使用费明细": {"文件名": _source_label(default_equipment), "统计截止日期": equipment_cutoff or "未知"},
        "人工费汇总": {"文件名": _source_label(default_labor), "统计截止日期": labor_cutoff_by_name or "未知"},
        "责任成本": {"文件名": _source_label(default_responsibility_cost_dir), "编制日期": responsibility_cutoff or "未知"},
    }

    contract_dir = default_contract_dir
    output_root = default_output
    project_reports_dir = output_root / PROJECT_REPORTS_DIRNAME
    project_reports_dir.mkdir(parents=True, exist_ok=True)

    datasets = _safe_load("合同管理", load_contracts_from_folder, contract_dir, default={"总包合同": pd.DataFrame(), "分包合同": pd.DataFrame()}) or {
        "总包合同": pd.DataFrame(),
        "分包合同": pd.DataFrame(),
    }
    procurement_df = _safe_load("采购明细", load_procurement, default_procurement, default=pd.DataFrame())
    material_arrears_df = _safe_load("材料欠款", load_material_arrears, default_material_arrears, default=pd.DataFrame())
    concrete_df = _safe_load("商混混凝土", load_concrete, default_concrete, default=pd.DataFrame())
    equipment_df = _safe_load("机具使用费", load_equipment, default_equipment, default=pd.DataFrame())
    labor_df = _safe_load("人工费汇总", load_labor, default_labor, default=pd.DataFrame())
    labor_month_df = _safe_load("人工费按月", load_labor_monthly, default_labor, default=pd.DataFrame())
    labor_cutoff = detect_latest_year_month(labor_month_df if labor_month_df is not None else pd.DataFrame())
    if labor_cutoff:
        data_source_info["人工费汇总"]["统计截止日期"] = labor_cutoff
    finance_loaded = _safe_load("财务付款", load_finance_summary, default_finance, default=(pd.DataFrame(), pd.DataFrame()))
    finance_summary_df, finance_category_df = finance_loaded if finance_loaded else (pd.DataFrame(), pd.DataFrame())
    cost_stats_df = _safe_load("成本统计", load_cost_statistics, default_cost_stats, default=pd.DataFrame())
    cost_stats_cutoff_from_data = infer_cutoff_from_month_column(cost_stats_df if cost_stats_df is not None else pd.DataFrame(), "月份")
    if cost_stats_cutoff_from_data:
        data_source_info["项目成本统计汇总表"]["统计截止日期"] = cost_stats_cutoff_from_data
    project_name_map = _safe_load("项目名称映射", load_project_name_map, default_manager_map, default={}) or {}
    manager_map = _safe_load("项目经理映射", load_project_manager_map, default_manager_map, default={}) or {}
    budget_df = _safe_load("预算金额", load_budget_summary, default_budget, default=pd.DataFrame())
    responsibility_cost_mapping = (
        load_responsibility_cost_mapping(default_responsibility_cost_dir)
        if default_responsibility_cost_dir
        else {}
    )
    sub_contract_df = (
        _safe_load("分包合同台账", load_subcontract_ledger, default_sub_contract, default=pd.DataFrame())
        if default_sub_contract
        else None
    )

    procurement_df = procurement_df if procurement_df is not None else pd.DataFrame()
    material_arrears_df = material_arrears_df if material_arrears_df is not None else pd.DataFrame()
    concrete_df = concrete_df if concrete_df is not None else pd.DataFrame()
    equipment_df = equipment_df if equipment_df is not None else pd.DataFrame()
    labor_df = labor_df if labor_df is not None else pd.DataFrame()
    labor_month_df = labor_month_df if labor_month_df is not None else pd.DataFrame()
    finance_summary_df = finance_summary_df if finance_summary_df is not None else pd.DataFrame()
    finance_category_df = finance_category_df if finance_category_df is not None else pd.DataFrame()
    cost_stats_df = cost_stats_df if cost_stats_df is not None else pd.DataFrame()
    budget_df = budget_df if budget_df is not None else pd.DataFrame()

    main_df = datasets.get("总包合同", pd.DataFrame())
    sub_df = sub_contract_df if sub_contract_df is not None and not sub_contract_df.empty else datasets.get("分包合同", pd.DataFrame())

    summary_rows: List[Dict] = []
    for (
        code,
        project_name,
        summary_row,
        main_rows,
        sub_rows,
        procurement_rows,
        extra_proc_rows,
        concrete_rows,
        equipment_rows,
        labor_rows,
        labor_month_rows,
        finance_category_rows,
    ) in build_project_summary(main_df, sub_df, procurement_df, material_arrears_df, concrete_df, equipment_df, labor_df, labor_month_df, finance_summary_df, finance_category_df, budget_df):
        official_project_name = project_name_map.get(str(code).strip().upper(), "")
        if official_project_name:
            project_name = official_project_name
            summary_row["项目名称"] = official_project_name

        export_project_report(
            project_reports_dir,
            code,
            project_name,
            summary_row,
            main_rows,
            sub_rows,
            procurement_rows,
            extra_proc_rows,
            concrete_rows,
            equipment_rows,
            labor_rows,
            labor_month_rows,
            finance_category_rows,
            cost_stats_df,
            budget_df,
            responsibility_cost_mapping,
            data_source_info,
        )
        summary_rows.append(summary_row)

    export_index_file(output_root, summary_rows, cost_stats_df, budget_df)
    (
        input_rows,
        cost_stats_fields,
        procurement_fields,
        concrete_fields,
        equipment_fields,
        labor_fields,
        finance_fields,
        budget_fields,
    ) = build_provenance_rows(
        contract_dir=contract_dir,
        procurement_path=default_procurement,
        concrete_paths=default_concrete,
        equipment_path=default_equipment,
        labor_path=default_labor,
        finance_path=default_finance,
        cost_stats_path=default_cost_stats,
        manager_map_path=default_manager_map,
        budget_path=default_budget,
    )

    export_provenance_excel(
        output_root,
        input_rows,
        cost_stats_fields,
        procurement_fields,
        concrete_fields,
        equipment_fields,
        labor_fields,
        finance_fields,
        budget_fields,
    )

    replicate_outputs(
        project_reports_dir,
        output_root,
        summary_rows,
        manager_map,
        PROJECT_MANAGER_DIRNAME,
        PROJECT_STATUS_DIRNAME,
    )

    cleanup_legacy_output_files(
        output_root,
        project_reports_dir,
        summary_rows,
        PROJECT_MANAGER_DIRNAME,
        PROJECT_STATUS_DIRNAME,
    )

    if labor_cutoff:
        print(f"✓ 人工费数据最新统计月份：{labor_cutoff}")
    print(f"\n✓ 完成！已生成 {len(summary_rows)} 个项目的合同汇总")
    print(f"📁 输出根目录：{output_root}")
    print(f"📁 项目合同汇总：{project_reports_dir}")
    print(f"📁 项目经理目录：{output_root / PROJECT_MANAGER_DIRNAME}")
    print(f"📁 项目状态目录：{output_root / PROJECT_STATUS_DIRNAME}")
    print(f"📄 索引文件：{output_root / '项目汇总索引.xlsx'}")
    print(f"📄 字段映射：{output_root / '数据来源与字段映射.xlsx'}")


if __name__ == "__main__":
    run_cli()
