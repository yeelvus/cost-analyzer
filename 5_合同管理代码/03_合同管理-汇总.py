# -*- coding: utf-8 -*-)
"""
合同管理文件层级提取数据 - )14（新增：列排序记忆对话框）
功能增强：
1. Excel写入前弹窗调整列顺序（拖拽）
2. 记忆用户排序，下次自动应用
3. 不调整 → 维持原顺序
"""

import os
import re
import json
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# =========================
# 【快速切换配置】
# =========================
TARGET_KEYWORDS = ["总包合同", "分包合同"]

SETTLEMENT_CHECKS = [
    {
        "name": "竣工结算资料",
        "pattern": r"^H_[A-Z0-9-]+_竣工结算资料$",
        "empty_ok": False
    },
]

# 这些列不参与统计且不输出到Excel
OUTPUT_EXCLUDED_COLUMNS = {
    "竣工结算资料",
    "合同结算额",
    "结算金额",
    "合同签订日期",
    "计划开工时间",
    "计划竣工时间",
    "开工日期",
    "生效日期",
    "联系人及电话",
    "是否有业主结算资料",
    "付款方式",
    "关联主合同编号",
}

# =========================
# 配置 + 列顺序记忆
# =========================
# 【配置】在此设置项目主目录路径（留空则弹出选择对话框）
ROOT_DIR_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑"

COLUMN_ORDER_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "column_order.txt")  # 新增：记忆列顺序（固定保存到当前代码目录）
PROJECT_CODE_DB_FILE = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"  # 项目名称/别名 -> 项目编码库
PROJECT_CODE_DB_FALLBACK = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"

# 使用配置路径或弹出选择对话框
if ROOT_DIR_PATH and os.path.isdir(ROOT_DIR_PATH):
    root_dir = ROOT_DIR_PATH
    print(f"使用配置路径：{root_dir}")
else:
    raise FileNotFoundError(f"ROOT_DIR_PATH 无效，请检查路径配置：{ROOT_DIR_PATH}")

output_dir = root_dir
keyword_str = "+".join(TARGET_KEYWORDS)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"{keyword_str}汇总_{timestamp}.xlsx"
output_path = os.path.join(output_dir, output_filename)

if os.path.exists(output_path):
    try:
        os.remove(output_path)
    except PermissionError:
        raise PermissionError(f"文件被占用，请关闭后重试：{output_path}")

# =========================
# 读取 .txt + 判断非空
# =========================
def read_txt_file(file_path):
    content = ""
    for enc in ["utf-8", "gbk"]:
        try:
            with open(file_path, "r", encoding=enc, errors="ignore") as f:
                content = f.read().strip()
            break
        except:
            continue
    return content

def read_contract_data(folder_path):
    """
    读取合同数据（兼容两种模式）：
    优先级：JSON > 多个txt文件
    返回 dict: {"对方单位公司名称": "", "合同编号": "", ...}
    """
    default_data = {"对方单位公司名称": "", "合同编号": "", "合同签订日期": "", "币种": "", "合同金额": 0.0}
    
    # 尝试读取 JSON 文件
    json_path = os.path.join(folder_path, "合同信息.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)
            # 转换金额类型
            if "合同金额" in json_data:
                try:
                    json_data["合同金额"] = float(json_data["合同金额"])
                except:
                    json_data["合同金额"] = 0.0
            result = {**default_data, **json_data}
            return result
        except Exception as e:
            print(f"  警告：JSON 读取失败 {json_path}: {e}")
    
    # 回退：读取 txt 文件
    data = default_data.copy()
    for f in os.listdir(folder_path):
        if not f.lower().endswith(".txt"):
            continue
        content = read_txt_file(os.path.join(folder_path, f))
        if not content:
            continue
        field = f[:-4]
        
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

def is_folder_non_empty(folder_path):
    if not os.path.isdir(folder_path):
        return False
    return len([f for f in os.listdir(folder_path) if not f.startswith('.')]) > 0

def natural_sort_key(text):
    """自然排序：将文本转换为可自然排序的元组
    例如：'01', '02', '10' 会按 (1, 1), (2, 1), (10, 1) 排序
    而不是 ('01', '10', '2')
    """
    import re
    def convert(s):
        return int(s) if s.isdigit() else s.lower()
    return [convert(c) for c in re.split('([0-9]+)', text)]

def get_sort_key(subfolder_name, signature_date, contract_amount):
    """获取排序键：优先数字排序，然后日期，最后金额"""
    # 检查子文件夹名称是否以数字开头
    if subfolder_name and subfolder_name[0].isdigit():
        # 以数字开头，使用自然排序
        return (0, natural_sort_key(subfolder_name), "", 0)
    else:
        # 不以数字开头，按日期、再按金额排序
        # 将日期转换为可比较的格式（字符串或日期对象）
        date_str = str(signature_date).strip() if signature_date else ""
        amount = float(contract_amount) if contract_amount and contract_amount != 0.0 else 0.0
        return (1, "", date_str, amount)

def detect_attachment_files(folder_path):
    """检测子文件夹本层的附件文件（不递归）"""
    if not os.path.isdir(folder_path):
        return {"有PDF": "", "有Word": "", "有Excel": ""}
    
    has_pdf = False
    has_word = False
    has_excel = False
    
    try:
        for item in os.listdir(folder_path):
            if item.startswith('.'):
                continue
            item_path = os.path.join(folder_path, item)
            if not os.path.isfile(item_path):
                continue
            
            lower_name = item.lower()
            if lower_name.endswith('.pdf'):
                has_pdf = True
            elif lower_name.endswith(('.doc', '.docx')):
                has_word = True
            elif lower_name.endswith(('.xls', '.xlsx')):
                has_excel = True
    except:
        pass
    
    return {
        "有PDF": "✓" if has_pdf else "",
        "有Word": "✓" if has_word else "",
        "有Excel": "✓" if has_excel else ""
    }


def load_project_code_mapping(db_path):
    """从项目编码库加载编码映射。返回 code->(标准名, 编码)"""
    code_to_meta = {}
    if not os.path.exists(db_path):
        print(f"警告：项目编码库不存在：{db_path}")
        return code_to_meta

    try:
        with open(db_path, "r", encoding="utf-8") as f:
            db = json.load(f)
    except Exception as e:
        print(f"警告：读取项目编码库失败：{e}")
        return code_to_meta

    projects = db.get("projects", {}) if isinstance(db, dict) else {}
    for std_name, meta in projects.items():
        if not isinstance(meta, dict):
            continue
        project_code = str(meta.get("project_code", "")).strip()
        if not project_code:
            continue
        code_to_meta[project_code] = (std_name, project_code)

    print(f"已加载项目编码库：{db_path}，可匹配项目编码 {len(code_to_meta)} 个")
    return code_to_meta


def extract_project_code_from_folder_name(folder_name):
    """文件夹名中下划线前的字段视为项目编码。"""
    text = str(folder_name).strip()
    if not text:
        return ""
    return text.split("_", 1)[0].strip()


def resolve_project_by_code(raw_project_name, code_to_meta):
    """仅按项目编码匹配：从文件夹名提取编码后在编码库中查找。"""
    raw = str(raw_project_name).strip()
    if not raw:
        return "", ""

    project_code = extract_project_code_from_folder_name(raw)
    if project_code and project_code in code_to_meta:
        return code_to_meta[project_code]

    # 未匹配时保留原名，并保留提取到的编码
    return raw, project_code


def infer_project_context(contract_folder_path, sub_abs_path, code_to_meta):
    """从合同子路径推断项目状态、项目名、项目编码与项目目录（兼容新旧层级）。"""
    rel_path = os.path.relpath(sub_abs_path, contract_folder_path)
    parts = [] if rel_path == "." else rel_path.split(os.sep)

    # 新结构下通常第1层是项目状态（如 A_在建项目 / D_1_已完工...）
    # 旧结构或异常结构下回退到合同类型目录名
    project_status = parts[0] if parts else os.path.basename(contract_folder_path)

    raw_project_name = ""
    project_dir = ""

    # 优先识别“项目目录”层：编码在映射库中，或形如 THCC-2025-001 的编码模式
    for idx, part in enumerate(parts):
        code = extract_project_code_from_folder_name(part)
        if code and (code in code_to_meta or re.match(r"^[A-Za-z]{2,}-\d{4}-\d+$", code)):
            raw_project_name = part
            project_dir = os.path.join(contract_folder_path, *parts[:idx + 1])
            break

    # 回退：常见为 [状态, 项目, 合同子文件夹...]
    if not raw_project_name:
        if len(parts) >= 2:
            raw_project_name = parts[1]
            project_dir = os.path.join(contract_folder_path, parts[0], parts[1])
        elif len(parts) == 1:
            raw_project_name = parts[0]
            project_dir = os.path.join(contract_folder_path, parts[0])
        else:
            raw_project_name = os.path.basename(os.path.dirname(contract_folder_path))
            project_dir = os.path.dirname(contract_folder_path)

    project_name, project_code = resolve_project_by_code(raw_project_name, code_to_meta)
    return project_status, project_name, project_code, project_dir

# =========================
# 列顺序读取（直接使用固定文件）
# =========================
def load_column_order(columns, default_order):
    """直接读取固定列顺序文件；不存在或读取失败则使用默认顺序。"""
    if os.path.exists(COLUMN_ORDER_FILE):
        try:
            with open(COLUMN_ORDER_FILE, "r", encoding="utf-8") as f:
                saved = [line.strip() for line in f.readlines() if line.strip() in columns]
            return saved + [c for c in columns if c not in saved]
        except Exception:
            return default_order
    return default_order

# =========================
# 扫描 + 提取数据（同 v13）
# =========================
all_data = {}
project_counts = {}
settlement_results = {}

print(f"开始扫描根目录：{root_dir}")
print(f"正在查找文件夹名包含 {TARGET_KEYWORDS} 的目录...\n")

# =========================
# 加载项目编码映射库
# =========================
project_db_file = PROJECT_CODE_DB_FILE if os.path.exists(PROJECT_CODE_DB_FILE) else PROJECT_CODE_DB_FALLBACK
print(f"使用项目编码库：{project_db_file}")
project_code_map = load_project_code_mapping(project_db_file)

for current_root, dirs, files in os.walk(root_dir):
    folder_name = os.path.basename(current_root)
    if not any(kw in folder_name for kw in TARGET_KEYWORDS):
        continue
    parent_name = os.path.basename(os.path.dirname(current_root))
    if any(kw in parent_name for kw in TARGET_KEYWORDS):
        continue
    dirs[:] = []

    hit_kw = next(kw for kw in TARGET_KEYWORDS if kw in folder_name)
    if hit_kw not in all_data:
        all_data[hit_kw] = []
        project_counts[hit_kw] = 0
    project_counts[hit_kw] += 1

    contract_folder_path = current_root
    print(f"[{project_counts[hit_kw]}] 发现 {hit_kw} 文件夹：{contract_folder_path}")

    # 读取 txt（同前）
    subfolder_data = {}
    for sub_root, _, sub_files in os.walk(contract_folder_path):
        rel_path = os.path.relpath(sub_root, contract_folder_path)
        sub_key = rel_path if rel_path != "." else folder_name
        abs_sub_path = sub_root
        
        # 使用新的兼容函数读取合同数据
        data = read_contract_data(abs_sub_path)
        subfolder_data[sub_key] = data

    for sub, fields in subfolder_data.items():
        if all(str(v).strip() == "" or v == 0.0 for k, v in fields.items() if k != "币种"):
            continue
        sub_abs_path = os.path.join(contract_folder_path, sub) if sub != folder_name else contract_folder_path
        sub_abs_path = os.path.normpath(sub_abs_path)

        project_status, project_name, project_code, project_dir = infer_project_context(
            contract_folder_path,
            sub_abs_path,
            project_code_map,
        )
        settlement_key = project_code if project_code else project_name

        # 检测结算资料：按“项目目录”识别，兼容新旧结构
        if settlement_key and settlement_key not in settlement_results:
            settlement_results[settlement_key] = {check["name"]: "" for check in SETTLEMENT_CHECKS}

        if settlement_key:
            for check in SETTLEMENT_CHECKS:
                # 已有“有资料”则不再重复扫描
                if settlement_results[settlement_key].get(check["name"]) == "有资料":
                    continue
                if not os.path.isdir(project_dir):
                    continue
                for item in os.listdir(project_dir):
                    item_path = os.path.join(project_dir, item)
                    if not os.path.isdir(item_path):
                        continue
                    if re.match(check["pattern"], item):
                        if is_folder_non_empty(item_path):
                            settlement_results[settlement_key][check["name"]] = "有资料"
                        break
        
        # 检测附件文件
        attachments = detect_attachment_files(sub_abs_path)
        
        rec = {
            "项目状态": project_status, "项目编码": project_code, "项目名称": project_name, "合同子文件夹": sub,
            "对方单位公司名称": fields.get("对方单位公司名称", ""), "合同编号": fields.get("合同编号", ""),
            "合同签订日期": fields.get("合同签订日期", ""), "币种": fields.get("币种", ""), "合同金额": fields.get("合同金额", 0.0),
            "有PDF": attachments["有PDF"], "有Word": attachments["有Word"], "有Excel": attachments["有Excel"],
            "_folder_path": sub_abs_path
        }
        for k, v in fields.items():
            if k not in rec:
                rec[k] = v
        all_data[hit_kw].append(rec)

print("\n扫描完成：")
for kw in TARGET_KEYWORDS:
    cnt = project_counts.get(kw, 0)
    rec = len(all_data.get(kw, []))
    print(f"  {kw}：{cnt} 个文件夹，{rec} 条记录")

# 输出未匹配项目提示
unmatched_projects = sorted({
    rec.get("项目名称", "")
    for rows in all_data.values()
    for rec in rows
    if not str(rec.get("项目编码", "")).strip()
})
if unmatched_projects:
    print("\n以下项目未在项目编码库匹配到编码：")
    for name in unmatched_projects:
        print(f"  - {name}")

# =========================
# 写入 Excel + 列排序弹窗
# =========================
fixed_cols = ["项目状态", "项目编码", "项目名称", "合同子文件夹", "对方单位公司名称", "币种", "合同金额", "有PDF", "有Word", "有Excel"]
settlement_cols = [check["name"] for check in SETTLEMENT_CHECKS if check["name"] not in OUTPUT_EXCLUDED_COLUMNS]
default_cols_order = fixed_cols + ["汇率", "金额（万泰铢）"] + settlement_cols

has_any_data = any(len(rows) > 0 for rows in all_data.values())

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    if not has_any_data:
        pd.DataFrame([{"提示": "未找到符合条件的 .txt 文件"}]).to_excel(writer, index=False, sheet_name="说明")
    else:
        # 确保总包合同在第一位
        sorted_keywords = sorted(all_data.keys(), key=lambda x: (x != "总包合同", x))
        for kw in sorted_keywords:
            rows = all_data[kw]
            if not rows:
                continue
            df = pd.DataFrame(rows).fillna("")

            dynamic_cols = sorted([
                c for c in df.columns
                if c not in fixed_cols and c != "_folder_path" and c not in OUTPUT_EXCLUDED_COLUMNS
            ])

            for col in ["汇率", "金额（万泰铢）"] + settlement_cols:
                if col not in df.columns:
                    df[col] = ""

            # 所有列
            all_columns = list(set(fixed_cols + ["汇率", "金额（万泰铢）"] + settlement_cols + dynamic_cols))
            default_order = [c for c in default_cols_order if c in all_columns] + [c for c in all_columns if c not in default_cols_order]

            # 不再弹窗，直接按固定文件读取列顺序
            final_cols_order = load_column_order(all_columns, default_order)

            output_cols = [c for c in final_cols_order if c in df.columns and c not in OUTPUT_EXCLUDED_COLUMNS]
            if not output_cols:
                continue
            df = df[output_cols]

            # 按合同签订日期降序排列（最新日期在前，空日期在最后）
            if "合同签订日期" in df.columns:
                # 将空日期转换为一个很早的日期字符串，使其排在最后
                df["_date_sort"] = df["合同签订日期"].apply(
                    lambda x: str(x).strip() if str(x).strip() else "0000-00-00"
                )
                df = df.sort_values(by="_date_sort", ascending=False)
                df = df.drop(columns=["_date_sort"])
            
            # 分类汇总
            grouped = []
            for (code, name), g in df.groupby(["项目编码", "项目名称"], sort=False):
                dates = [d for d in g["合同签订日期"] if str(d).strip()] if "合同签订日期" in g.columns else []
                min_date = min(dates) if dates else ""
                cur = "CNY" if "CNY" in g["币种"].values else g["币种"].iloc[0] if not g.empty else ""
                settlement_key = str(code).strip() if str(code).strip() else name
                # 仅按当前分组已有列构造汇总行，避免把已排除列重新加回去
                sum_row = {col: "" for col in g.columns}
                if "项目状态" in sum_row:
                    sum_row["项目状态"] = g["项目状态"].iloc[0]
                if "项目编码" in sum_row:
                    sum_row["项目编码"] = "" if code is None else str(code)
                if "项目名称" in sum_row:
                    sum_row["项目名称"] = "" if name is None else str(name)
                if "合同子文件夹" in sum_row:
                    sum_row["合同子文件夹"] = "（分类汇总）"
                if "币种" in sum_row:
                    sum_row["币种"] = cur
                if "合同签订日期" in sum_row:
                    sum_row["合同签订日期"] = min_date
                for check_name in settlement_cols:
                    if check_name in sum_row:
                        sum_row[check_name] = settlement_results.get(str(settlement_key).strip(), {}).get(check_name, "")
                for c in dynamic_cols:
                    if c not in sum_row:
                        continue
                    if any(k in c for k in ["金额", "数量", "面积"]):
                        sum_row[c] = pd.to_numeric(g[c], errors='coerce').sum()
                    elif "日期" in c:
                        v = [x for x in g[c] if str(x).strip()]
                        sum_row[c] = min(v) if v else ""
                    else:
                        sum_row[c] = ""
                grouped.append(pd.concat([g, pd.DataFrame([sum_row])], ignore_index=True))

            final = pd.concat(grouped, ignore_index=True)
            # 优先使用记录中的汇率（非空且不等于 0/1 的情况保留），否则按币种默认：CNY->5, 其他->1
            def _choose_rate(row):
                try:
                    v = row.get("汇率") if isinstance(row, dict) else row["汇率"]
                except Exception:
                    v = None
                # 有效数值且不是 0/1 则保留
                if v is not None and str(v).strip() != "":
                    try:
                        fv = float(v)
                        if fv != 0 and fv != 1:
                            return fv
                    except Exception:
                        pass
                # 回退到币种默认
                return 5 if (row.get("币种") if isinstance(row, dict) else row["币种"]) == "CNY" else 1

            # 使用 apply 时传入 axis=1，需要基于 DataFrame 行对象判断
            final["汇率"] = final.apply(lambda r: _choose_rate(r.to_dict()), axis=1)
            final["金额（万泰铢）"] = ""
            final.to_excel(writer, index=False, sheet_name=kw[:31])

# =========================
# 格式化（同前）
# =========================
wb = load_workbook(output_path)
folder_paths = {}
for kw, rows in all_data.items():
    for rec in rows:
        key = (kw, rec.get("项目名称", rec.get("总包合同", "")), rec["合同子文件夹"])
        folder_paths[key] = rec["_folder_path"]

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.max_row == 1 and ws["A1"].value == "提示":
        continue
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    highlight = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    bold = Font(bold=True)
    col_map = {}
    for idx, col_name in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
        cell = col_name[0]
        col_map[cell.value] = get_column_letter(idx)
    rate_col = col_map.get("汇率")
    money_col = col_map.get("合同金额")
    target_col = col_map.get("金额（万泰铢）")
    subfolder_col = col_map.get("合同子文件夹")

    for r_idx in range(2, ws.max_row + 1):
        sub_val = ws[f"{subfolder_col}{r_idx}"].value if subfolder_col else ""
        proj_col = col_map.get('项目名称', col_map.get('总包合同'))
        proj_val = ws[f"{proj_col}{r_idx}"].value

        if sub_val == "（分类汇总）":
            for c in ws[r_idx]:
                c.fill = highlight
                c.font = bold
            if target_col:
                start = r_idx
                while start > 2 and ws[f"{proj_col}{start-1}"].value == proj_val:
                    start -= 1
                ws[f"{target_col}{r_idx}"] = f"=SUM({target_col}{start}:{target_col}{r_idx-1})"
        else:
            key = (sheet_name, proj_val, sub_val)
            folder_path = folder_paths.get(key, "")
            if folder_path and os.path.isdir(folder_path) and subfolder_col:
                safe_path = folder_path.replace("\\", "/")
                cell = ws[f"{subfolder_col}{r_idx}"]
                cell.hyperlink = f"file:///{safe_path}"
                cell.font = Font(color="0000FF", underline="single")
                cell.style = "Hyperlink"
            if target_col and rate_col and money_col:
                ws[f"{target_col}{r_idx}"] = f"=({money_col}{r_idx}*{rate_col}{r_idx}/10000)"

    # 设置“金额（万泰铢）”列显示为千分位逗号且保留两位小数
    if target_col:
        for cell in ws[target_col][1:]:
            cell.number_format = '#,##0.00'
    
    # 设置所有行高为35
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        row_idx = row[0].row
        if row_idx is not None:
            ws.row_dimensions[row_idx].height = 35
    
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    start = None
    for r in range(2, ws.max_row + 1):
        sub = ws[f"{subfolder_col}{r}"].value if subfolder_col else ""
        if isinstance(sub, str) and "（分类汇总）" in sub:
            if start:
                for x in range(start, r):
                    ws.row_dimensions[x].outline_level = 1
            start = None
        elif start is None:
            start = r
    if start:
        for x in range(start, ws.max_row + 1):
            if "（分类汇总）" not in str(ws[f"{subfolder_col}{x}"].value):
                ws.row_dimensions[x].outline_level = 1

wb.save(output_path)
wb.close()

# =========================
# 完成提示
# =========================
if has_any_data:
    print(f"\n汇总完成！列顺序已记忆\n文件保存至：{output_path}")
else:
    print(f"\n未找到数据，已生成空文件：{output_path}")
print(f"成功生成：{output_path}")