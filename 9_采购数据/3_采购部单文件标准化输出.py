import os
import sys
import re
import importlib.util
from datetime import datetime, date as date_type
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


# ============= 路径配置 =============
# 文件夹按关键词动态查找，忽略 A_ / B_ 等前缀字母
_路径_spec = importlib.util.spec_from_file_location("_采购路径", str(Path(__file__).parent / "_采购路径.py"))
if _路径_spec is None or _路径_spec.loader is None:
    raise ImportError("无法加载路径模块: _采购路径.py")
_路径_mod = importlib.util.module_from_spec(_路径_spec)
_路径_spec.loader.exec_module(_路径_mod)
RAW_DATA_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/B_采购部采购台账_std/采购部采购台账"
OUTPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/C_采购部采购台账_simplify"
PROJECT_INFO_FILE = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
UNIT_MAP_UTILS_PATH = "/Users/wyattsun/Desktop/code/0_项目基本信息/unit_map_utils.py"

# ── common 模块加载（单例，避免孤立实例）──────────────────────────────────
_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

from settings import MATCH_DB_PATH

PROJECT_MATCH_DB_FILE = str(MATCH_DB_PATH)
PROJECT_MATCHER_PATH = str(_CODE_ROOT / "0_项目基本信息" / "project_name_matcher.py")


def _load_project_matcher_module(file_path: str):
    spec = importlib.util.spec_from_file_location("_project_name_matcher_runtime", file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载项目匹配模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_project_matcher = _load_project_matcher_module(PROJECT_MATCHER_PATH)
load_or_init_db = _project_matcher.load_or_init_db
save_db = _project_matcher.save_db
match_project = _project_matcher.match_project
mark_unmatched = _project_matcher.mark_unmatched
add_mapping = _project_matcher.add_mapping
get_project_lookup = _project_matcher.get_project_lookup
is_unmatched_alias = _project_matcher.is_unmatched_alias
prompt_unmatched_project_mappings = _project_matcher.prompt_unmatched_project_mappings


# ============= 标准化流水线配置 =============
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
STANDARDIZATION_PIPELINE_DIR = os.path.join(CURRENT_DIR, "物品名称规格标准化")
PIPELINE_CONTEXT = None


# ============= 列配置 =============
STANDARD_COLUMNS = [
    "材料归类", "物品名称", "规格", "单位", "数量",
    "单价", "不含税金额", "7%增值税", "含税金额", "货币种类", "汇率",
    "签订日期", "供货商名称", "备注",
]

CORE_REQUIRED_COLUMNS = ["物品名称", "单位", "数量"]

FINAL_OUTPUT_COLUMNS = [
    "项目编码", "项目名称", "材料归类", "物品名称", "规格", "单位", "数量",
    "单价", "不含税金额", "7%增值税", "含税金额", "货币种类", "汇率", "年月",
    "标准日期", "供货商名称", "备注", "签订日期",
]

CENTER_ALIGN_COLUMNS = {"单位", "数量", "货币种类", "汇率"}
LEFT_ALIGN_COLUMNS = {"签订日期"}
AMOUNT_COLUMNS = {"单价", "不含税金额", "7%增值税", "含税金额"}


COLUMN_MAPPINGS = {
    "材料归类": ["材料归类"],
    "物品名称": ["物品名称", "采购物资名称", "物资名称"],
    "规格": ["规格", "规格型号", "型号", "规格型号/参数"],
    "单位": ["标准单位", "单位"],
    "数量": ["数量"],
    "单价": ["单价", "不含税单价", "含税单价"],
    "不含税金额": ["不含税金额", "金额"],
    "7%增值税": ["7%增值税", "税额", "增值税", "税金"],
    "含税金额": ["含税金额", "总金额", "价税合计"],
    "货币种类": ["货币种类", "币种"],
    "汇率": ["汇率"],
    "签订日期": ["签订日期", "日期"],
    "供货商名称": ["供货商名称", "供货单位", "供应商", "供应商名称"],
    "备注": ["备注"],
}


def _load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

# 加载单位标准化模块（含 Tkinter 交互对话框与 unit_map.json 读写）
_unit_map_utils = _load_module_from_path("unit_map_utils", UNIT_MAP_UTILS_PATH)
standardize_unit = _unit_map_utils.standardize_unit


def load_project_matcher_funcs():
    return {
        "load_or_init_db": load_or_init_db,
        "save_db": save_db,
        "match_project": match_project,
        "mark_unmatched": mark_unmatched,
        "add_mapping": add_mapping,
        "get_project_lookup": get_project_lookup,
        "is_unmatched_alias": is_unmatched_alias,
        "prompt_unmatched_project_mappings": prompt_unmatched_project_mappings,
    }


def load_project_config():
    try:
        df = pd.read_excel(PROJECT_INFO_FILE, sheet_name=0)
    except Exception as exc:
        print(f"警告：无法加载项目配置 {PROJECT_INFO_FILE}: {exc}")
        return []

    required_cols = ["项目编码", "项目名称"]
    if any(col not in df.columns for col in required_cols):
        print("警告：项目配置缺少必要列，后续仅按文件名输出项目名称")
        return []

    projects = []
    for _, row in df.iterrows():
        code = str(row["项目编码"]).strip() if pd.notna(row["项目编码"]) else ""
        name = str(row["项目名称"]).strip() if pd.notna(row["项目名称"]) else ""
        if code or name:
            projects.append({"project_code": code, "project_name": name})
    return projects


def extract_project_name(filename):
    base_name = os.path.splitext(filename)[0]
    if "_" in base_name:
        return base_name.split("_", 1)[0].strip()
    match = re.match(r"^(.+?项目)", base_name)
    if match:
        return match.group(1).strip()
    return base_name.strip()


def find_matching_project(extracted_name, project_config, match_project, match_db):
    hit = match_project(extracted_name, match_db)
    if hit:
        return {
            "project_code": hit.get("project_code", ""),
            "project_name": hit.get("project_name", extracted_name),
        }

    for project in project_config:
        if project["project_name"] == extracted_name:
            return project

    for project in project_config:
        if extracted_name in project["project_name"] or project["project_name"] in extracted_name:
            return project

    return None


def list_excel_files(root_dir):
    excel_paths = []
    for root, _, files in os.walk(root_dir):
        for name in sorted(files):
            if name.startswith("~$"):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in {".xlsx", ".xls", ".xlsm"}:
                excel_paths.append(os.path.join(root, name))
    return sorted(excel_paths)


def preload_project_matches_for_files(excel_files, project_config, matcher_funcs, match_db):
    """批量收集未识别项目名并一次性调用匹配模块，减少逐文件弹窗。"""
    unmatched_projects = []
    for file_path in excel_files:
        extracted_name = extract_project_name(os.path.basename(file_path))
        matched = find_matching_project(extracted_name, project_config, matcher_funcs["match_project"], match_db)
        if matched:
            continue
        if extracted_name not in unmatched_projects:
            unmatched_projects.append(extracted_name)

    if not unmatched_projects:
        return

    print(f"\n发现 {len(unmatched_projects)} 个未识别项目，打开批量匹配窗口...")
    batch_results = matcher_funcs["prompt_unmatched_project_mappings"](
        unmatched_projects,
        match_db,
        window_title="采购-批量项目匹配",
    )
    project_lookup = matcher_funcs["get_project_lookup"](match_db)
    for raw_name, std_name in batch_results.items():
        if std_name:
            matcher_funcs["add_mapping"](match_db, raw_name, std_name)
            code, _ = project_lookup.get(std_name, ("", ""))
            print(f"  ✓ 匹配: '{raw_name}' -> [{code}] {std_name}")
        else:
            matcher_funcs["mark_unmatched"](match_db, raw_name)
            print(f"  - 留空: '{raw_name}'")


def normalize_headers(headers):
    normalized = []
    for col in headers:
        stripped_col = str(col).strip() if col else ""
        for std_col, variations in COLUMN_MAPPINGS.items():
            if stripped_col in variations:
                normalized.append(std_col)
                break
        else:
            normalized.append(stripped_col)
    return normalized


def find_header_row(sheet):
    required_header_hits = ["物品名称", "单位", "数量"]

    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        if row and row[0] and str(row[0]).strip() in ["序号", "No.", "编号"]:
            row_values = [str(val).strip() for val in row if val is not None]

            # 先判断关键列是否齐全，兼容历史台账表头（如“采购物资名称/规格型号/总金额/日期”）
            required_matched = []
            for key in required_header_hits:
                aliases = COLUMN_MAPPINGS.get(key, [key])
                if any(alias in row_values for alias in aliases):
                    required_matched.append(key)

            if len(required_matched) == len(required_header_hits):
                return row_idx

            matched = [
                col for col in STANDARD_COLUMNS
                if any(alias in row_values for alias in COLUMN_MAPPINGS.get(col, [col]))
            ]
            if len(matched) >= len(STANDARD_COLUMNS) // 2:
                return row_idx
    return None


def is_summary_row(formula_sheet, row_idx, amount_col):
    amount_cell = formula_sheet.cell(row_idx, amount_col)
    if amount_cell.data_type != "f":
        return False

    formula_str = str(amount_cell.value).upper()
    if "SUM" not in formula_str:
        return False

    col_letter = get_column_letter(amount_col)
    if f"{col_letter}:{col_letter}" in formula_str:
        return True

    range_pattern = f"\\$?{col_letter}\\$?\\d+:\\$?{col_letter}\\$?\\d+"
    return re.search(range_pattern, formula_str) is not None


def parse_date(date_str):
    if date_str is None:
        return (None, None, None)

    if isinstance(date_str, datetime):
        return (date_str.year, date_str.month, date_str.day)
    if isinstance(date_str, date_type):
        return (date_str.year, date_str.month, date_str.day)

    s = str(date_str).strip()
    if not s:
        return (None, None, None)

    match = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})(?:\s+\d{1,2}:\d{2}:\d{2})?", s)
    if match:
        y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return (y, m, d)

    match = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[\s T]\d{1,2}:\d{2}:\d{2})?", s)
    if match:
        y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return (y, m, d)

    match = re.match(r"^(\d{4})\.(\d{1,2})(?:\.(\d{1,2}))?", s)
    if match:
        y, m = int(match.group(1)), int(match.group(2))
        d = int(match.group(3)) if match.group(3) else None
        if 1 <= m <= 12 and (d is None or 1 <= d <= 31):
            return (y, m, d)

    return (None, None, None)


def is_numeric_value(value):
    """判断值是否可视为数值（用于数量列校验）。"""
    if value is None:
        return False
    if isinstance(value, (int, float)) and pd.notna(value):
        return True

    s = str(value).strip()
    if not s:
        return False
    s = s.replace(",", "")
    return re.fullmatch(r"[+-]?\d+(?:\.\d+)?", s) is not None


def _to_price_key(value):
    """将单价标准化为可比较的浮点键，无法转换返回 None。"""
    num = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(num):
        return None
    return round(float(num), 6)


def _to_date_ts(date_text):
    """将标准日期字符串转为时间戳，无法解析返回 None。"""
    if date_text is None:
        return None
    s = str(date_text).strip()
    if not s:
        return None
    ts = pd.to_datetime(s, errors="coerce")
    if pd.isna(ts):
        return None
    return ts


def _within_price_tolerance(base_price, candidate_price, tolerance=0.50):
    """候选单价是否在基准单价的允许偏差范围内（默认50%）。"""
    if base_price is None or candidate_price is None:
        return False
    try:
        base = float(base_price)
        cand = float(candidate_price)
    except Exception:
        return False

    if base == 0:
        return abs(cand - base) < 1e-9

    return abs(cand - base) / abs(base) <= tolerance


# 全局单位参考库：{物品名称: 最常见单位}，跨文件兜底使用
_GLOBAL_UNIT_REF: dict = {}


def build_global_unit_reference(std_root: str) -> dict:
    """扫描 STD 目录中已输出的 Excel，统计每个物品名称出现最多的单位，返回 {物品名称: 单位} 字典。"""
    ref: dict = {}
    from collections import Counter
    if not os.path.isdir(std_root):
        return ref
    for root, _, files in os.walk(std_root):
        for name in files:
            if name.startswith("~$") or os.path.splitext(name)[1].lower() not in {".xlsx", ".xls"}:
                continue
            fpath = os.path.join(root, name)
            try:
                df = pd.read_excel(fpath, sheet_name=0, usecols=["物品名称", "单位"], dtype=str)
            except Exception:
                continue
            df = df.dropna(subset=["物品名称"])
            df["物品名称"] = df["物品名称"].str.strip()
            df["单位"] = df["单位"].fillna("").str.strip()
            for item_name, grp in df[df["单位"].ne("")].groupby("物品名称"):
                counter = ref.setdefault(item_name, Counter())
                counter.update(grp["单位"].tolist())
    # 每个物品名称取出现最多的单位
    return {k: v.most_common(1)[0][0] for k, v in ref.items()}


def fill_missing_units_by_similarity(df, global_unit_ref: dict | None = None):
    """按 物品名称+单价 回填空单位；多候选时优先同项目，再按日期最近。
    单价匹配范围：同物品名称所有行的平均单价 ±50%。
    同文件无候选时，用跨文件全局单位参考库兜底。
    """
    if df.empty or "单位" not in df.columns:
        return df, 0

    required_cols = {"物品名称", "单价", "单位", "项目名称", "标准日期"}
    if not required_cols.issubset(df.columns):
        return df, 0

    out_df = df.copy()
    out_df["_name_key"] = out_df["物品名称"].astype(str).str.strip()
    out_df["_unit_key"] = out_df["单位"].astype(str).str.strip()
    out_df["_price_key"] = out_df["单价"].apply(_to_price_key)
    out_df["_date_ts"] = out_df["标准日期"].apply(_to_date_ts)

    # 预计算每个物品名称的平均单价（所有行，排除0和空）
    avg_price_map: dict = {}
    for name_key, grp in out_df[out_df["_name_key"].ne("") & out_df["_price_key"].notna()].groupby("_name_key"):
        valid_prices = grp["_price_key"].dropna()
        valid_prices = valid_prices[valid_prices != 0]
        if not valid_prices.empty:
            avg_price_map[name_key] = float(valid_prices.mean())

    # 候选池：有单位 + 有物品名称 + 单价可解析
    candidate_df = out_df[
        out_df["_unit_key"].ne("")
        & out_df["_name_key"].ne("")
        & out_df["_price_key"].notna()
    ]

    # 待回填：单位为空 + 有物品名称 + 单价可解析
    target_idx = out_df[
        out_df["_unit_key"].eq("")
        & out_df["_name_key"].ne("")
        & out_df["_price_key"].notna()
    ].index

    if candidate_df.empty and (not global_unit_ref) and len(target_idx) == 0:
        return df, 0
    if len(target_idx) == 0:
        return df, 0

    filled_count = 0
    for idx in target_idx:
        row = out_df.loc[idx]
        avg_price = avg_price_map.get(row["_name_key"])
        same_pool = candidate_df[
            (candidate_df["_name_key"] == row["_name_key"])
            & candidate_df["_price_key"].apply(lambda p: _within_price_tolerance(avg_price, p, 0.50))
        ]
        if same_pool.empty:
            # 兜底：从全局单位参考库查找
            if global_unit_ref:
                fallback_unit = global_unit_ref.get(str(row["_name_key"]).strip(), "")
                if fallback_unit:
                    out_df.at[idx, "单位"] = fallback_unit
                    out_df.at[idx, "_unit_key"] = fallback_unit
                    filled_count += 1
            continue

        target_project = str(row.get("项目名称", "")).strip()
        target_ts = row.get("_date_ts", None)

        ranked = []
        for _, c in same_pool.iterrows():
            same_project_rank = 0 if str(c.get("项目名称", "")).strip() == target_project else 1
            cand_ts = c.get("_date_ts", None)
            if target_ts is not None and cand_ts is not None:
                date_gap = abs((cand_ts - target_ts).days)
            else:
                date_gap = 10**9
            ranked.append((same_project_rank, date_gap, str(c["单位"]).strip()))

        ranked.sort(key=lambda x: (x[0], x[1]))
        best_unit = ranked[0][2] if ranked else ""
        if best_unit:
            out_df.at[idx, "单位"] = best_unit
            out_df.at[idx, "_unit_key"] = best_unit
            filled_count += 1

    out_df = out_df.drop(columns=["_name_key", "_unit_key", "_price_key", "_date_ts"], errors="ignore")
    return out_df, filled_count


def infer_date_for_project(project_data):
    result = list(project_data)

    for i in range(1, len(result)):
        row_idx, year, month, day = result[i]
        if year is None:
            for j in range(i - 1, -1, -1):
                _, py, pm, pd_ = result[j]
                if py is not None:
                    result[i] = (row_idx, py, pm, pd_)
                    break

    for i in range(len(result) - 2, -1, -1):
        row_idx, year, month, day = result[i]
        if year is None:
            for j in range(i + 1, len(result)):
                _, ny, nm, nd = result[j]
                if ny is not None:
                    result[i] = (row_idx, ny, nm, nd)
                    break

    return result


def format_date_str(year, month, day):
    if year is None:
        return ""
    if day is None:
        return f"{year}-{month:02d}"
    return f"{year}-{month:02d}-{day:02d}"


def get_pipeline_context():
    """缓存流水线模块与规则模块，避免每个文件重复import。"""
    global PIPELINE_CONTEXT
    if PIPELINE_CONTEXT is not None:
        return PIPELINE_CONTEXT

    pipeline_dir = Path(STANDARDIZATION_PIPELINE_DIR)
    pipeline_script = pipeline_dir / "0_统一标准化流水线.py"
    if not pipeline_script.exists():
        PIPELINE_CONTEXT = {"ok": False, "reason": f"未找到标准化流水线脚本: {pipeline_script}"}
        return PIPELINE_CONTEXT

    spec = importlib.util.spec_from_file_location("_standardization_pipeline", str(pipeline_script))
    if spec is None or spec.loader is None:
        PIPELINE_CONTEXT = {"ok": False, "reason": f"无法加载标准化流水线: {pipeline_script}"}
        return PIPELINE_CONTEXT

    pipeline_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(pipeline_module)

    rule_names = getattr(pipeline_module, "DEFAULT_RULES", [])
    if not rule_names:
        PIPELINE_CONTEXT = {"ok": False, "reason": "标准化流水线 DEFAULT_RULES 为空"}
        return PIPELINE_CONTEXT

    rule_modules = []
    for rule_name in rule_names:
        rule_file = pipeline_dir / rule_name
        if not rule_file.exists():
            continue
        try:
            rule_module = pipeline_module.load_module_from_file(rule_file)
            rule_modules.append((rule_file.name, rule_module))
        except Exception:
            continue

    PIPELINE_CONTEXT = {
        "ok": True,
        "pipeline_module": pipeline_module,
        "rule_modules": rule_modules,
    }
    return PIPELINE_CONTEXT


def apply_standardization_pipeline(excel_path):
    context = get_pipeline_context()
    if not context.get("ok"):
        print(f"  ⚠️ {context.get('reason', '标准化流水线不可用，已跳过')}")
        return

    workbook = openpyxl.load_workbook(excel_path)
    pipeline_module = context["pipeline_module"]
    for rule_name, rule_module in context["rule_modules"]:
        try:
            stats = pipeline_module.run_rule_on_workbook(rule_module, workbook, rule_name)
            print(
                f"  ✓ {rule_name}: "
                f"命中 {stats['matched']} 行, 更新 {stats['updated']} 行, 未识别 {stats['unresolved']} 行"
            )
        except Exception as exc:
            print(f"  ✗ 规则执行失败 {rule_name}: {exc}")

    workbook.save(excel_path)


def extract_rows_from_single_excel(file_path, project_code, project_name):
    formula_wb = openpyxl.load_workbook(file_path, data_only=False)
    formula_sheet = formula_wb.active
    if formula_sheet is None:
        formula_wb.close()
        reason = "未找到活动工作表"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason

    header_row_idx = find_header_row(formula_sheet)
    if header_row_idx is None:
        formula_wb.close()
        reason = "未识别到表头"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason

    headers = normalize_headers([cell.value for cell in formula_sheet[header_row_idx]])
    header_idx_map = {name: idx for idx, name in enumerate(headers)}

    missing_core = [col for col in CORE_REQUIRED_COLUMNS if col not in headers]
    if missing_core:
        formula_wb.close()
        reason = f"缺少核心列: {', '.join(missing_core)}"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason

    try:
        amount_col = header_idx_map["含税金额"] + 1
    except ValueError:
        formula_wb.close()
        reason = "缺少含税金额列"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason
    except KeyError:
        formula_wb.close()
        reason = "缺少含税金额列"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason

    summary_rows = set()
    for row_idx in range(header_row_idx + 1, formula_sheet.max_row + 1):
        if is_summary_row(formula_sheet, row_idx, amount_col):
            summary_rows.add(row_idx)

    formula_wb.close()

    value_wb = openpyxl.load_workbook(file_path, data_only=True)
    value_sheet = value_wb.active
    if value_sheet is None:
        value_wb.close()
        reason = "未找到活动工作表"
        print(f"  ⚠️ {reason}，跳过")
        return [], reason

    key_data_cols = ["物品名称", "规格", "标准单位", "数量", "单价", "不含税金额"]
    key_data_indices = [header_idx_map[col] for col in key_data_cols if col in header_idx_map]

    valid_rows = []
    row_values_map = {}
    for row_idx, row_vals in enumerate(
        value_sheet.iter_rows(min_row=header_row_idx + 1, max_row=value_sheet.max_row, values_only=True),
        start=header_row_idx + 1,
    ):
        if row_idx in summary_rows:
            continue

        if all(v is None or str(v).strip() == "" for v in row_vals):
            continue

        key_data_empty = True
        for idx_ in key_data_indices:
            v = row_vals[idx_] if idx_ < len(row_vals) else None
            if v is not None and str(v).strip() != "":
                key_data_empty = False
                break

        if key_data_empty:
            continue

        valid_rows.append(row_idx)
        row_values_map[row_idx] = row_vals

    rows = []
    for row_idx in valid_rows:
        row_vals = row_values_map[row_idx]
        row_obj = {"项目编码": project_code, "项目名称": project_name}

        for col_name in STANDARD_COLUMNS:
            if col_name == "单位":
                unit_idx = header_idx_map.get("单位")
                if unit_idx is not None and unit_idx < len(row_vals):
                    unit_val = row_vals[unit_idx]
                    row_obj[col_name] = standardize_unit(unit_val)
                else:
                    row_obj[col_name] = ""
            elif col_name in header_idx_map:
                source_idx = header_idx_map[col_name]
                row_obj[col_name] = row_vals[source_idx] if source_idx < len(row_vals) else ""
            else:
                row_obj[col_name] = ""

        # 纠错：若数量为非数字且单位为空，将数量内容挪到单位列。
        qty_val = row_obj.get("数量", "")
        unit_val = row_obj.get("单位", "")
        if str(unit_val).strip() == "":
            qty_text = str(qty_val).strip()
            if qty_text and not is_numeric_value(qty_val):
                row_obj["单位"] = qty_text
                row_obj["数量"] = ""

        rows.append(row_obj)

    value_wb.close()
    return rows, ""


def enrich_dates(df):
    if df.empty:
        df["标准日期"] = ""
        df["年月"] = ""
        return df

    parsed = [parse_date(v) for v in df["签订日期"].tolist()]
    project_data = [(idx, y, m, d) for idx, (y, m, d) in enumerate(parsed)]
    inferred = infer_date_for_project(project_data)

    std_dates = [""] * len(df)
    yms = [""] * len(df)
    for idx, y, m, d in inferred:
        std_dates[idx] = format_date_str(y, m, d)
        yms[idx] = f"{y}-{m:02d}" if y and m else ""

    df["标准日期"] = std_dates
    df["年月"] = yms
    return df


def write_output(df, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    out_df = df.copy()
    for col in FINAL_OUTPUT_COLUMNS:
        if col not in out_df.columns:
            out_df[col] = ""
    out_df = out_df[FINAL_OUTPUT_COLUMNS]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="合并数据", index=False)

    wb = openpyxl.load_workbook(output_path)
    ws = wb["合并数据"]

    ws.row_dimensions[1].height = 40
    ws.freeze_panes = "A2"

    for col_idx, header in enumerate(FINAL_OUTPUT_COLUMNS, start=1):
        cell = ws.cell(1, col_idx)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

        if header in AMOUNT_COLUMNS:
            for row_idx in range(2, ws.max_row + 1):
                c = ws.cell(row_idx, col_idx)
                if isinstance(c.value, (int, float)):
                    c.number_format = "0.00"

        if header in CENTER_ALIGN_COLUMNS:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row_idx, col_idx).alignment = Alignment(horizontal="center", vertical="center")
        elif header in LEFT_ALIGN_COLUMNS:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row_idx, col_idx).alignment = Alignment(horizontal="left", vertical="center")

    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    wb.save(output_path)


def build_output_path(input_path):
    rel_path = os.path.relpath(input_path, RAW_DATA_ROOT)
    rel_dir = os.path.dirname(rel_path)
    base_name, _ = os.path.splitext(os.path.basename(rel_path))
    output_name = f"{base_name}.xlsx"
    return os.path.join(OUTPUT_ROOT, rel_dir, output_name)


def process_single_file(file_path, project_config, matcher_funcs, match_db):
    file_name = os.path.basename(file_path)
    print(f"\n处理文件: {file_path}")

    extracted_name = extract_project_name(file_name)
    matched = find_matching_project(extracted_name, project_config, matcher_funcs["match_project"], match_db)

    if matched:
        project_code = matched.get("project_code", "")
        project_name = matched.get("project_name", extracted_name)
        matcher_funcs["add_mapping"](match_db, extracted_name, project_name)
    else:
        project_code = ""
        project_name = extracted_name
        matcher_funcs["mark_unmatched"](match_db, extracted_name)

    try:
        rows, extract_reason = extract_rows_from_single_excel(file_path, project_code, project_name)
    except Exception as exc:
        print(f"  ✗ 处理失败: {exc}")
        return False, f"处理异常: {exc}"

    if not rows:
        print("  ⚠️ 无有效数据，跳过输出")
        if extract_reason:
            return False, extract_reason
        return False, "无有效数据"

    df = pd.DataFrame(rows)

    # 统一处理本文件扫描到的未知单位（一次弹窗批量选择），并回填标准单位列。
    resolved_count = _unit_map_utils.resolve_pending_unit_mappings()
    if resolved_count and "单位" in df.columns:
        df["单位"] = df["单位"].apply(lambda v: standardize_unit(v, auto_prompt=False))

    # 仅保留金额合计不为0的行。
    amount_df = df[[c for c in ["单价", "不含税金额", "7%增值税", "含税金额"] if c in df.columns]].apply(
        pd.to_numeric, errors="coerce"
    )
    row_sum = amount_df.fillna(0).sum(axis=1)
    df = df[row_sum.ne(0)]
    if df.empty:
        print("  ⚠️ 金额合计均为0，跳过输出")
        return False, "金额合计均为0"

    df = enrich_dates(df)

    # 单位补全：单位为空时，按同物品名称+单价回填，优先同项目再按日期最近。
    df, filled_unit_count = fill_missing_units_by_similarity(df, global_unit_ref=_GLOBAL_UNIT_REF)
    if filled_unit_count:
        print(f"  ✓ 单位自动补全: {filled_unit_count} 行")

    output_path = build_output_path(file_path)
    write_output(df, output_path)
    apply_standardization_pipeline(output_path)

    print(f"  ✓ 已输出: {output_path}")
    return True, ""


def main():
    global _GLOBAL_UNIT_REF
    if not os.path.isdir(RAW_DATA_ROOT):
        raise FileNotFoundError(f"输入目录不存在: {RAW_DATA_ROOT}")

    print(f"项目匹配模块: {PROJECT_MATCHER_PATH}")

    os.makedirs(OUTPUT_ROOT, exist_ok=True)

    project_config = load_project_config()
    matcher_funcs = load_project_matcher_funcs()
    match_db = matcher_funcs["load_or_init_db"](PROJECT_MATCH_DB_FILE, project_config)

    # 预先扫描 STD 目录，建立跨文件单位参考库（已有输出结果）
    _GLOBAL_UNIT_REF = build_global_unit_reference(OUTPUT_ROOT)
    print(f"全局单位参考库: 已加载 {len(_GLOBAL_UNIT_REF)} 个物品名称")

    excel_files = list_excel_files(RAW_DATA_ROOT)
    if not excel_files:
        print("未发现Excel文件，程序结束")
        return

    preload_project_matches_for_files(excel_files, project_config, matcher_funcs, match_db)

    success_count = 0
    skip_count = 0
    skipped_details = []

    print(f"扫描到 {len(excel_files)} 个Excel文件，开始逐个标准化输出...")
    for file_path in excel_files:
        ok, reason = process_single_file(file_path, project_config, matcher_funcs, match_db)
        if ok:
            success_count += 1
        else:
            skip_count += 1
            skipped_details.append((file_path, reason or "未知原因"))

    matcher_funcs["save_db"](PROJECT_MATCH_DB_FILE, match_db)

    print("\n处理完成")
    print(f"成功输出: {success_count} 个")
    print(f"跳过/失败: {skip_count} 个")
    if skipped_details:
        print("\n跳过/失败文件明细：")
        for idx, (path, reason) in enumerate(skipped_details, start=1):
            print(f"  {idx}. {path}  -> {reason}")
    print(f"输出根目录: {OUTPUT_ROOT}")


if __name__ == "__main__":
    main()
