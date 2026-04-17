import json
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


def normalize_project_code(value) -> str:
    if pd.isna(value):
        return ""
    code = str(value).strip().upper()
    if code.lower() in {"nan", "none"}:
        return ""
    return code


def normalize_project_name(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def coerce_float(value) -> float:
    if pd.isna(value):
        return 0.0
    try:
        return float(str(value).replace(",", ""))
    except Exception:
        return 0.0


def split_project_fields(project: str) -> Tuple[str, str]:
    parts = str(project).split("_", 1)
    code = parts[0].strip()
    name = parts[1].strip() if len(parts) > 1 else ""
    return code, name


def detect_procurement(row: pd.Series) -> bool:
    keywords = ["采购", "材料", "买卖", "设备", "物资"]
    text = f"{row.get('合同类型', '')}{row.get('合同子文件夹', '')}"
    return any(key in str(text) for key in keywords)


def ensure_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = None
    return df


def convert_amount_to_thb(amount, currency: str, rate, is_wan_thb: bool = False) -> float:
    base = coerce_float(amount)
    if is_wan_thb:
        base *= 10000
    if base == 0:
        return 0.0
    currency_upper = str(currency).strip().upper()
    if currency_upper in {"THB", "泰铢", "TBH", "THAILAND BAHT", "NA", ""}:
        return base
    fx = coerce_float(rate)
    if fx == 0:
        fx = 5.0 if currency_upper in {"CNY", "人民币"} else 1.0
    return base * fx


def scale_wan_amount_columns(df: pd.DataFrame, amount_keywords: List[str], exclude_cols: Optional[List[str]] = None) -> pd.DataFrame:
    df = df.copy()
    exclude_set = set(exclude_cols or [])
    for col in df.columns:
        if col in exclude_set:
            continue
        col_text = str(col)
        if not any(keyword in col_text for keyword in amount_keywords):
            continue
        df[col] = df[col].apply(lambda x: coerce_float(x) * 10000)
    return df


def load_project_manager_map(manager_path: Path) -> Dict[str, str]:
    if not manager_path.exists():
        return {}

    def _read_candidates() -> List[pd.DataFrame]:
        result: List[pd.DataFrame] = []
        for header in (0, 1):
            try:
                df_tmp = pd.read_excel(manager_path, sheet_name=0, header=header)
                if not df_tmp.empty:
                    result.append(df_tmp)
            except Exception:
                continue
        return result

    def _pick_col(columns, candidates: List[str]) -> Optional[str]:
        for cand in candidates:
            if cand in columns:
                return cand
        return None

    for df in _read_candidates():
        cols = [str(c).strip() for c in df.columns]
        df.columns = cols

        manager_col = _pick_col(cols, ["项目经理", "项目经理人", "经理", "负责人"])
        if manager_col is None:
            for col in cols:
                if "经理" in col:
                    manager_col = col
                    break
        if manager_col is None:
            continue

        code_col = _pick_col(cols, ["项目编码", "项目代码", "项目编号", "编码"])
        name_col = _pick_col(cols, ["项目名称", "项目全称", "项目简称"])

        mapping: Dict[str, str] = {}
        for _, row in df.iterrows():
            code = normalize_project_code(row.get(code_col, "")) if code_col else ""

            if not code and name_col:
                name_val = str(row.get(name_col, "")).strip()
                match = re.search(r"(THCC-\d{4}-\d{3})", name_val, re.IGNORECASE)
                if match:
                    code = match.group(1).upper()

            if not code:
                continue

            manager = str(row.get(manager_col, "")).strip() or "其他"
            mapping[code] = manager

        if mapping:
            return mapping

    return {}


def load_project_name_map(manager_path: Path) -> Dict[str, str]:
    """从项目基本信息表读取 `项目编码 -> 项目名称` 映射。"""
    if not manager_path.exists():
        return {}

    def _read_candidates() -> List[pd.DataFrame]:
        result: List[pd.DataFrame] = []
        for header in (0, 1):
            try:
                df_tmp = pd.read_excel(manager_path, sheet_name=0, header=header)
                if not df_tmp.empty:
                    result.append(df_tmp)
            except Exception:
                continue
        return result

    def _pick_col(columns, candidates: List[str]) -> Optional[str]:
        for cand in candidates:
            if cand in columns:
                return cand
        return None

    for df in _read_candidates():
        cols = [str(c).strip() for c in df.columns]
        df.columns = cols

        code_col = _pick_col(cols, ["项目编码", "项目代码", "项目编号", "编码"])
        name_col = _pick_col(cols, ["项目名称", "项目全称", "项目简称"])
        if code_col is None or name_col is None:
            continue

        mapping: Dict[str, str] = {}
        for _, row in df.iterrows():
            code = normalize_project_code(row.get(code_col, ""))
            name = normalize_project_name(row.get(name_col, ""))
            if code and name:
                mapping[code] = name

        if mapping:
            return mapping

    return {}


def read_txt_file(file_path: Path) -> str:
    content = ""
    for enc in ["utf-8", "gbk"]:
        try:
            with open(file_path, "r", encoding=enc, errors="ignore") as f:
                content = f.read().strip()
            break
        except Exception:
            continue
    return content


def read_contract_data(folder_path: Path) -> Dict:
    default_data = {"对方单位公司名称": "", "合同编号": "", "合同签订日期": "", "币种": "", "合同金额": 0.0}
    json_path = folder_path / "合同信息.json"
    if json_path.exists():
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)
            if "合同金额" in json_data:
                try:
                    json_data["合同金额"] = float(json_data["合同金额"])
                except Exception:
                    json_data["合同金额"] = 0.0
            return {**default_data, **json_data}
        except Exception:
            pass

    data = default_data.copy()
    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".txt"):
            continue
        content = read_txt_file(folder_path / filename)
        if not content:
            continue
        field = filename[:-4]
        if field == "合同金额":
            upper = content.upper()
            if any(x in upper for x in ["CNY", "人民币"]):
                data["币种"] = "CNY"
            elif any(x in upper for x in ["THB", "泰铢"]):
                data["币种"] = "THB"
            nums = re.findall(r"\d[\d,\.]*\d", content)
            clean = [float(n.replace(",", "")) for n in nums if re.fullmatch(r"-?\d+\.?\d*", n.replace(",", ""))]
            data["合同金额"] = sum(clean) if clean else 0.0
        else:
            data[field] = content
    return data


def detect_attachment_files(folder_path: Path) -> Dict[str, str]:
    if not folder_path.is_dir():
        return {"有PDF": "", "有Word": "", "有Excel": ""}
    has_pdf = has_word = has_excel = False
    for item in folder_path.iterdir():
        if item.name.startswith(".") or not item.is_file():
            continue
        lower_name = item.name.lower()
        if lower_name.endswith(".pdf"):
            has_pdf = True
        elif lower_name.endswith((".doc", ".docx")):
            has_word = True
        elif lower_name.endswith((".xls", ".xlsx")):
            has_excel = True
    return {"有PDF": "✓" if has_pdf else "", "有Word": "✓" if has_word else "", "有Excel": "✓" if has_excel else ""}


def load_contracts_from_folder(root_dir: Path) -> Dict[str, pd.DataFrame]:
    target_keywords = ["总包合同", "分包合同"]
    mapping_path = root_dir / "project_mapping.json"
    project_mapping: Dict[str, str] = {}
    if mapping_path.exists():
        try:
            loaded = json.loads(mapping_path.read_text(encoding="utf-8"))
            first_val = next(iter(loaded.values())) if loaded else None
            if isinstance(first_val, list):
                for standard, raws in loaded.items():
                    for raw in raws:
                        project_mapping[raw] = standard
            else:
                project_mapping = loaded
        except Exception:
            project_mapping = {}

    records = {kw: [] for kw in target_keywords}

    def _looks_like_project_dir(name: str) -> bool:
        if not name or "_" not in name:
            return False
        prefix = name.split("_", 1)[0].strip().upper()
        return ("-" in prefix and any(ch.isdigit() for ch in prefix))

    def _append_project_contracts(hit_kw: str, project_dir: Path, project_status: str):
        raw_project_name = project_dir.name
        if raw_project_name not in project_mapping:
            project_mapping[raw_project_name] = raw_project_name
        project_name = project_mapping[raw_project_name]

        contract_folder = project_dir
        subfolder_data: Dict[str, Dict] = {}
        for sub_root, _, __ in os.walk(contract_folder):
            rel_path = Path(sub_root).relative_to(contract_folder)
            sub_key = str(rel_path) if str(rel_path) != "." else contract_folder.name
            abs_sub_path = Path(sub_root)
            data = read_contract_data(abs_sub_path)
            subfolder_data[sub_key] = data

        for sub_key, fields in subfolder_data.items():
            if all(str(v).strip() == "" or v == 0.0 for k, v in fields.items() if k != "币种"):
                continue
            sub_abs_path = (contract_folder / sub_key).resolve()
            attachments = detect_attachment_files(sub_abs_path)
            rec = {
                "项目状态": project_status,
                "项目名称": project_name,
                "合同子文件夹": sub_key,
                "对方单位公司名称": fields.get("对方单位公司名称", ""),
                "合同编号": fields.get("合同编号", ""),
                "合同签订日期": fields.get("合同签订日期", ""),
                "币种": fields.get("币种", ""),
                "合同金额": fields.get("合同金额", 0.0),
                "汇率": fields.get("汇率", ""),
                "备注": fields.get("备注", ""),
                **attachments,
            }
            for k, v in fields.items():
                if k not in rec:
                    rec[k] = v
            rec["金额_含税优先"] = rec.get("合同金额", 0.0)
            rec["金额_THB"] = convert_amount_to_thb(rec.get("合同金额", 0.0), rec.get("币种", ""), rec.get("汇率", ""))
            rec["项目编码"], rec["项目简称"] = split_project_fields(project_name)
            rec["是否采购类合同"] = detect_procurement(pd.Series(rec))
            records[hit_kw].append(rec)

    # 优先解析新结构：合同类型根目录 -> 项目状态 -> 项目 -> 合同子文件夹
    for contract_root in root_dir.iterdir() if root_dir.exists() else []:
        if not contract_root.is_dir():
            continue
        root_name = contract_root.name
        if "签证" in root_name:
            continue
        if not any(kw in root_name for kw in target_keywords):
            continue

        hit_kw = next(kw for kw in target_keywords if kw in root_name)
        for status_or_project in contract_root.iterdir():
            if not status_or_project.is_dir():
                continue

            # 兼容“合同类型根目录 -> 项目”与“合同类型根目录 -> 状态 -> 项目”两种层级
            if _looks_like_project_dir(status_or_project.name):
                _append_project_contracts(hit_kw, status_or_project, "")
                continue

            project_status = status_or_project.name
            for project_dir in status_or_project.iterdir():
                if not project_dir.is_dir() or not _looks_like_project_dir(project_dir.name):
                    continue
                _append_project_contracts(hit_kw, project_dir, project_status)

    # 旧结构兜底：项目目录下直接包含“总包合同/分包合同”目录
    if all(len(records[kw]) == 0 for kw in target_keywords):
        for current_root, dirs, _ in os.walk(root_dir):
            folder_name = os.path.basename(current_root)
            if not any(kw in folder_name for kw in target_keywords):
                continue
            parent_name = os.path.basename(os.path.dirname(current_root))
            if any(kw in parent_name for kw in target_keywords):
                continue
            dirs[:] = []

            hit_kw = next(kw for kw in target_keywords if kw in folder_name)
            parent_dir = Path(current_root).parent
            raw_project_name = parent_dir.name
            if raw_project_name not in project_mapping:
                project_mapping[raw_project_name] = raw_project_name
            project_name = project_mapping[raw_project_name]
            grandparent_dir = parent_dir.parent
            project_status = grandparent_dir.name if grandparent_dir != root_dir else ""

            contract_folder = Path(current_root)
            subfolder_data: Dict[str, Dict] = {}
            for sub_root, _, __ in os.walk(contract_folder):
                rel_path = Path(sub_root).relative_to(contract_folder)
                sub_key = str(rel_path) if str(rel_path) != "." else folder_name
                abs_sub_path = Path(sub_root)
                data = read_contract_data(abs_sub_path)
                subfolder_data[sub_key] = data

            for sub_key, fields in subfolder_data.items():
                if all(str(v).strip() == "" or v == 0.0 for k, v in fields.items() if k != "币种"):
                    continue
                sub_abs_path = (contract_folder / sub_key).resolve()
                attachments = detect_attachment_files(sub_abs_path)
                rec = {
                    "项目状态": project_status,
                    "项目名称": project_name,
                    "合同子文件夹": sub_key,
                    "对方单位公司名称": fields.get("对方单位公司名称", ""),
                    "合同编号": fields.get("合同编号", ""),
                    "合同签订日期": fields.get("合同签订日期", ""),
                    "币种": fields.get("币种", ""),
                    "合同金额": fields.get("合同金额", 0.0),
                    "汇率": fields.get("汇率", ""),
                    "备注": fields.get("备注", ""),
                    **attachments,
                }
                for k, v in fields.items():
                    if k not in rec:
                        rec[k] = v
                rec["金额_含税优先"] = rec.get("合同金额", 0.0)
                rec["金额_THB"] = convert_amount_to_thb(rec.get("合同金额", 0.0), rec.get("币种", ""), rec.get("汇率", ""))
                rec["项目编码"], rec["项目简称"] = split_project_fields(project_name)
                rec["是否采购类合同"] = detect_procurement(pd.Series(rec))
                records[hit_kw].append(rec)

    base_columns = [
        "项目状态", "项目名称", "项目编码", "项目简称", "合同子文件夹", "合同金额", "币种", "汇率", "金额_THB", "金额_含税优先", "是否采购类合同",
    ]

    result = {}
    for kw, rows in records.items():
        result[kw] = pd.DataFrame(rows) if rows else pd.DataFrame(columns=base_columns)
        result[kw] = ensure_columns(result[kw], base_columns)
    return result


def load_procurement(procurement_path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(procurement_path)
    preferred_sheets = ["合并数据", "采购明细", "明细", "项目汇总"]
    sheet = next((s for s in preferred_sheets if s in xl.sheet_names), xl.sheet_names[0])
    df = xl.parse(sheet)

    if "项目编码" not in df.columns:
        for alt in ["标准项目编码", "project_code"]:
            if alt in df.columns:
                df["项目编码"] = df[alt]
                break

    if "含税金额" not in df.columns:
        for alt in ["总采购金额", "金额_THB", "金额", "不含税金额"]:
            if alt in df.columns:
                df["含税金额"] = df[alt]
                break

    if "标准项目名称" not in df.columns:
        for alt in ["项目名称", "project_name"]:
            if alt in df.columns:
                df["标准项目名称"] = df[alt]
                break

    if "项目编码" in df.columns and "项目状态" in df.columns:
        blank_code_mask = df["项目编码"].isna() | (df["项目编码"].astype(str).str.strip() == "")
        status_code = df["项目状态"].astype(str).str.extract(r"(THCC-\d{4}-\d{3})", expand=False)
        df.loc[blank_code_mask & status_code.notna(), "项目编码"] = status_code[blank_code_mask & status_code.notna()]

    df = df.dropna(subset=["项目编码"])
    df["项目编码"] = df["项目编码"].apply(normalize_project_code)
    df = df[df["项目编码"] != ""]

    if "标准项目名称" in df.columns:
        name_series = df["标准项目名称"]
    elif "项目名称" in df.columns:
        name_series = df["项目名称"]
    else:
        name_series = pd.Series([""] * len(df), index=df.index)
    df["标准项目名称"] = name_series.apply(normalize_project_name)

    df["金额_THB"] = df["含税金额"].apply(coerce_float)

    if "物品名称" not in df.columns:
        for alt in ["材料名称", "名称", "品名"]:
            if alt in df.columns:
                df["物品名称"] = df[alt]
                break

    return df


def load_concrete(concrete_paths: List[Path]) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for path in concrete_paths:
        if not path.exists():
            print(f"警告：商混文件不存在，已跳过: {path}")
            continue

        try:
            xl = pd.ExcelFile(path)
            preferred_sheets = ["客户总额汇总", "计算明细(全)", "计算明细", "汇总", xl.sheet_names[0]]
            sheet = next((s for s in preferred_sheets if s in xl.sheet_names), xl.sheet_names[0])
            df = xl.parse(sheet)
        except Exception as e:
            print(f"警告：读取商混文件失败 {path}: {e}")
            continue

        if "金额" in df.columns:
            df["金额_THB"] = df["金额"].apply(coerce_float)
        elif "金额（泰铢）" in df.columns:
            df["金额_THB"] = df["金额（泰铢）"].apply(coerce_float)
        elif "计算总金额" in df.columns:
            df["金额_THB"] = df["计算总金额"].apply(coerce_float)
        elif "金额（万泰铢）" in df.columns:
            df["金额_THB"] = df["金额（万泰铢）"].apply(lambda x: coerce_float(x) * 10000)
        else:
            df["金额_THB"] = 0.0

        if "项目编码" not in df.columns:
            df["项目编码"] = df.get("项目名称", "")

        if "标准项目名称" not in df.columns:
            df["标准项目名称"] = df.get("项目名称", "")

        df["项目编码"] = df["项目编码"].apply(normalize_project_code)
        df["标准项目名称"] = df["标准项目名称"].apply(normalize_project_name)
        frames.append(df)

    if frames:
        df_all = pd.concat(frames, ignore_index=True)
        dropna_cols = ["项目编码"]
        if "标准项目名称" in df_all.columns:
            dropna_cols.append("标准项目名称")
        elif "项目名称" in df_all.columns:
            dropna_cols.append("项目名称")

        df_all = df_all.dropna(subset=dropna_cols, how="all")
        return df_all
    return pd.DataFrame()


def load_equipment(equipment_path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(equipment_path)
    sheet = "汇总" if "汇总" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet)
    df["sheet名称"] = sheet
    df.columns = [str(c).strip() for c in df.columns]

    if "项目编码" not in df.columns:
        for alt in ["标准项目编码", "project_code"]:
            if alt in df.columns:
                df["项目编码"] = df[alt]
                break

    if "项目名称" not in df.columns:
        for alt in ["标准项目名称", "project_name"]:
            if alt in df.columns:
                df["项目名称"] = df[alt]
                break

    amount_col = None
    for col in ["费用合计(泰铢)", "费用合计", "本次费用合计", "结算费用_泰铢", "月租金额_泰铢"]:
        if col in df.columns:
            amount_col = col
            break

    if amount_col is None:
        df["金额_THB"] = 0.0
        return df

    temp = df.copy()
    temp["项目编码"] = temp["项目编码"].apply(normalize_project_code)
    temp["项目名称"] = temp["项目名称"].apply(normalize_project_name)
    temp = temp[(temp["项目编码"] != "") | (temp["项目名称"] != "")]
    temp["金额_THB"] = temp[amount_col].apply(coerce_float)
    temp["标准项目名称"] = temp.get("标准项目名称", temp["项目名称"]).apply(normalize_project_name)
    return temp


def load_labor(labor_path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(labor_path)
    labor_sheet = "按项目汇总" if "按项目汇总" in xl.sheet_names else ("项目求和" if "项目求和" in xl.sheet_names else xl.sheet_names[0])
    df = xl.parse(labor_sheet)
    df = df.rename(columns={"project_code": "项目编码"})

    if "项目编码" not in df.columns:
        for alt in ["标准项目编码", "编码"]:
            if alt in df.columns:
                df["项目编码"] = df[alt]
                break

    if "项目名称" not in df.columns:
        for alt in ["标准项目名称", "project_name"]:
            if alt in df.columns:
                df["项目名称"] = df[alt]
                break

    if "人工费汇总金额" in df.columns and "人工费汇总（泰铢）" not in df.columns:
        df["人工费汇总（泰铢）"] = df["人工费汇总金额"].apply(coerce_float)

    if "总人工费" in df.columns and "人工费汇总（泰铢）" not in df.columns:
        df["人工费汇总（泰铢）"] = df["总人工费"].apply(coerce_float)

    if "总人工费" not in df.columns and "人工费汇总（泰铢）" in df.columns:
        df["总人工费"] = df["人工费汇总（泰铢）"]

    labor_cols = [
        "求和项:中方正式工铢", "求和项:中方劳务工铢", "求和项:子公司属地工铢", "求和项:劳务公司铢", "求和项:签证费",
    ]
    available_labor_cols = [col for col in labor_cols if col in df.columns]

    if "人工费汇总（泰铢）" in df.columns:
        df["人工费汇总（泰铢）"] = df["人工费汇总（泰铢）"].apply(coerce_float)
    if "总人工费" in df.columns:
        df["总人工费"] = df["总人工费"].apply(coerce_float)
    elif available_labor_cols:
        df["人工费汇总（泰铢）"] = df[available_labor_cols].apply(lambda row: sum(coerce_float(v) for v in row), axis=1)
    else:
        df["人工费汇总（泰铢）"] = 0.0

    if "总人工费" not in df.columns:
        df["总人工费"] = df["人工费汇总（泰铢）"]

    if "项目编码" in df.columns:
        df["项目编码"] = df["项目编码"].apply(normalize_project_code)
    if "项目名称" in df.columns:
        df["项目名称"] = df["项目名称"].apply(normalize_project_name)

    return df


def load_labor_monthly(labor_path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(labor_path)
    if "按年月项目汇总" in xl.sheet_names:
        df = xl.parse("按年月项目汇总")
    elif "按项目汇总" in xl.sheet_names:
        df = xl.parse("按项目汇总")
    else:
        return pd.DataFrame(columns=["项目编码", "项目名称", "年", "月"])

    df = df.rename(columns={"project_code": "项目编码"})
    df = scale_wan_amount_columns(
        df,
        amount_keywords=["万元"],
        exclude_cols=["项目编码", "project_code", "项目名称", "年", "月", "年-月", "target_status", "地区"],
    )
    return df


def load_subcontract_ledger(subcontract_path: Path) -> pd.DataFrame:
    if not subcontract_path.exists():
        return pd.DataFrame()

    xl = pd.ExcelFile(subcontract_path)
    preferred_sheets = ["项目汇总", "汇总", xl.sheet_names[0]]
    sheet = next((s for s in preferred_sheets if s in xl.sheet_names), xl.sheet_names[0])
    df = xl.parse(sheet)
    df.columns = [str(c).strip() for c in df.columns]

    if "项目编码" not in df.columns:
        for alt in ["标准项目编码", "project_code"]:
            if alt in df.columns:
                df["项目编码"] = df[alt]
                break

    if "项目名称" not in df.columns:
        for alt in ["标准项目名称", "project_name"]:
            if alt in df.columns:
                df["项目名称"] = df[alt]
                break

    if "合同额（泰铢）" not in df.columns:
        for alt in ["合同金额", "合同额", "含税金额"]:
            if alt in df.columns:
                df["合同额（泰铢）"] = df[alt]
                break

    df = ensure_columns(df, ["项目编码", "项目名称", "合同结算模式", "协作体名称", "工程范围", "合同额（泰铢）"])
    df = df.copy()
    df["项目编码"] = df["项目编码"].apply(normalize_project_code)
    df["项目名称"] = df["项目名称"].apply(normalize_project_name)
    df = df[(df["项目编码"] != "") | (df["项目名称"] != "")]

    df["项目简称"] = df["项目名称"]
    df["项目状态"] = ""
    df["金额_THB"] = df["合同额（泰铢）"].apply(coerce_float)
    df["是否采购类合同"] = False

    df["合同子文件夹"] = df["合同结算模式"]
    df["合同类型"] = df["合同结算模式"]
    df["合同金额"] = df["合同额（泰铢）"].apply(coerce_float)
    df["对方公司名称"] = df["协作体名称"]

    return df


def load_budget_summary(budget_excel_path: Path) -> pd.DataFrame:
    if not budget_excel_path.exists():
        return pd.DataFrame()

    try:
        xl = pd.ExcelFile(budget_excel_path)
        sheet_name = "总汇总"

        if sheet_name not in xl.sheet_names:
            print(f"警告：找不到了『{sheet_name}』 sheet")
            return pd.DataFrame()

        df = xl.parse(sheet_name)
        if "项目编码" not in df.columns:
            print("警告：找不到了『项目编码』列")
            return pd.DataFrame()

        keep_cols = ["序号", "项目编码", "项目名称", "币种"]
        amount_cols = [col for col in df.columns if col not in keep_cols and "项目名称" not in col and col != "序号"]
        keep_cols.extend(amount_cols)

        df = df[[col for col in keep_cols if col in df.columns]].copy()
        if "项目编码" in df.columns:
            df = df[df["项目编码"].astype(str) != "合计"].copy()

        for col in amount_cols:
            if col in df.columns:
                df[col] = df[col].apply(coerce_float)

        return df
    except Exception as e:
        print(f"警告：读取预算数据失败 {budget_excel_path}: {e}")
        return pd.DataFrame()


def load_cost_statistics(cost_stats_path: Path) -> pd.DataFrame:
    if not cost_stats_path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_excel(cost_stats_path)
        if "标准项目编码" not in df.columns and "项目编码" not in df.columns:
            return pd.DataFrame()

        if "标准项目编码" in df.columns:
            df["匹配编码"] = df["标准项目编码"]
        else:
            df["匹配编码"] = df["项目编码"]

        if "月份" in df.columns:
            df["_月份排序"] = pd.to_datetime(
                df["月份"].astype(str).str.extract(r"(\d{4})\D(\d{1,2})").apply(
                    lambda r: f"{r[0]}-{int(r[1]):02d}-01" if pd.notna(r[0]) and pd.notna(r[1]) else "1900-01-01",
                    axis=1,
                ),
                errors="coerce",
            )
            df_list = []
            for _, group in df.groupby("匹配编码"):
                group_sorted = group.sort_values("_月份排序", ascending=False)
                latest_row = group_sorted.iloc[0:1]
                df_list.append(latest_row)

            if df_list:
                df = pd.concat(df_list, ignore_index=True)
                print(f"✓ 成本统计数据：为每个项目提取最新月份数据（共 {len(df)} 个项目）")
            else:
                df = pd.DataFrame()
        else:
            print("警告：成本统计文件中未找到'月份'列，将使用全部数据")

        amount_columns = [
            "预计总成本", "收入", "毛利", "业主预计结算额", "分包结算额", "分包已付款", "材料采购实际已发生", "材料采购后续预估发生", "采购部采购金额", "项目部采购金额", "材料租赁费用",
            "材料租赁后续预估", "工资支出", "管理人员工资", "机具费支出", "机具费后续预估", "人员调遣支出", "人员调遣后续预估", "其他费用支出", "其他费用后续预估", "临建费用支出",
            "临建费用后续预估", "总包合同额(万泰铢)", "签证情况", "分包合同额(万泰铢)", "分包签证情况", "实际已发生金额",
        ]

        for col in df.columns:
            col_text = str(col)
            if col in amount_columns:
                continue
            if "付款财务数据" in col_text or "万泰铢" in col_text:
                amount_columns.append(col)

        for col in amount_columns:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: coerce_float(x) * 10000)

        if "_月份排序" in df.columns:
            df = df.drop(columns=["_月份排序"])

        return df
    except Exception as e:
        print(f"警告：读取成本统计数据失败 {cost_stats_path}: {e}")
        return pd.DataFrame()


def load_finance_summary(finance_path: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not finance_path.exists():
        return pd.DataFrame(), pd.DataFrame()

    try:
        df_summary = pd.read_excel(finance_path, sheet_name="项目汇总（万泰铢）")
        df_category = pd.read_excel(finance_path, sheet_name="分类汇总")
        return df_summary, df_category
    except Exception as e:
        print(f"警告：读取财务部数据失败 {finance_path}: {e}")
        return pd.DataFrame(), pd.DataFrame()
