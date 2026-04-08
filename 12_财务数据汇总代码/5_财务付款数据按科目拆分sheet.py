import os
import re
import importlib.util
from collections import OrderedDict

import pandas as pd
from openpyxl import load_workbook


INPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务部💰/A_财务付款数据/B_标准化数据/20260406_财务付款数据"
OUTPUT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务部💰/A_财务付款数据/B_财务付款数据类别拆分"
FORMAT_MODULE_PATH = "/Users/wyattsun/Desktop/19_代码/0_项目基本信息/excel_formatting_utils.py"


def load_apply_common_excel_layout(module_path):
    if not os.path.exists(module_path):
        raise FileNotFoundError(f"未找到格式模块: {module_path}")

    spec = importlib.util.spec_from_file_location("excel_formatting_utils", module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载格式模块: {module_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.apply_common_excel_layout


APPLY_COMMON_EXCEL_LAYOUT = load_apply_common_excel_layout(FORMAT_MODULE_PATH)


SUBJECT_COLUMN_ALIASES = [
    "科目",
    "会计科目",
    "财务科目",
    "费用科目",
]


CATEGORY_ALIAS_MAP = OrderedDict(
    {
        "分包款": [
            "分包款",
            "分包工程款",
            "分包付款",
            "专业分包",
            "劳务分包",
            "分包",
        ],
        "材料采购": [
            "材料采购",
            "材料款",
            "采购款",
            "物资采购",
            "原材料",
            "材料",
        ],
        "工资薪金": [
            "工资薪金",
            "工资",
            "薪金",
            "职工薪酬",
            "人员工资",
            "人工费",
            "薪酬",
        ],
        "机具使用费": [
            "机具使用费",
            "机具费",
            "机械费",
            "机械使用费",
            "设备使用费",
            "设备租赁费",
            "机具租赁费",
            "机具租赁",
        ],
    }
)


DEFAULT_CATEGORY = "管理费"
AMOUNT_HEADERS = ["应付金额", "预扣税额", "实付金额", "增值税额", "不含税额"]


def normalize_text(text):
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = s.replace("\n", "").replace("\r", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[()（）\[\]【】:：_\-]", "", s)
    return s


def find_subject_column(df):
    normalized_map = {normalize_text(col): col for col in df.columns}

    for alias in SUBJECT_COLUMN_ALIASES:
        key = normalize_text(alias)
        if key in normalized_map:
            return normalized_map[key]

    for alias in SUBJECT_COLUMN_ALIASES:
        key = normalize_text(alias)
        for norm_col, raw_col in normalized_map.items():
            if key and (key in norm_col or norm_col in key):
                return raw_col

    return None


def classify_subject(subject_value):
    norm = normalize_text(subject_value)
    if not norm:
        return DEFAULT_CATEGORY

    for category, aliases in CATEGORY_ALIAS_MAP.items():
        for alias in aliases:
            alias_norm = normalize_text(alias)
            if alias_norm and alias_norm in norm:
                return category

    return DEFAULT_CATEGORY


def is_excel_file(filename):
    if filename.startswith("~$"):
        return False
    ext = os.path.splitext(filename)[1].lower()
    return ext in {".xlsx", ".xls", ".xlsm"}


def build_unique_sheet_name(base_name, used_names):
    max_len = 31
    clean = str(base_name).strip() or "Sheet"
    clean = re.sub(r"[\\/*?:\[\]]", "_", clean)

    candidate = clean[:max_len]
    if candidate not in used_names:
        used_names.add(candidate)
        return candidate

    index = 2
    while True:
        suffix = f"_{index}"
        head = clean[: max_len - len(suffix)]
        candidate = f"{head}{suffix}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        index += 1


def split_one_sheet(df, subject_col):
    category_series = df[subject_col].apply(classify_subject)

    split_map = OrderedDict()
    for category in list(CATEGORY_ALIAS_MAP.keys()) + [DEFAULT_CATEGORY]:
        part = df[category_series == category].copy()
        if not part.empty:
            split_map[category] = part

    return split_map


def process_one_workbook(input_path, output_path):
    try:
        xl = pd.ExcelFile(input_path)
    except Exception as exc:
        print(f"  ⚠️ 读取失败: {os.path.basename(input_path)} | {exc}")
        return False

    if not xl.sheet_names:
        print(f"  ⚠️ 无sheet，跳过: {os.path.basename(input_path)}")
        return False

    per_sheet_result = []
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(input_path, sheet_name=sheet_name)
        except Exception as exc:
            print(f"    ⚠️ sheet读取失败: {sheet_name} | {exc}")
            continue

        if df.empty:
            continue

        subject_col = find_subject_column(df)
        if subject_col is None:
            print(f"    ⚠️ 未找到科目列，整表归入管理费: {sheet_name}")
            per_sheet_result.append((sheet_name, OrderedDict({DEFAULT_CATEGORY: df.copy()})))
            continue

        split_map = split_one_sheet(df, subject_col)
        if split_map:
            per_sheet_result.append((sheet_name, split_map))

    if not per_sheet_result:
        print(f"  ⚠️ 无有效数据，跳过: {os.path.basename(input_path)}")
        return False

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    used_sheet_names = set()
    single_source_sheet = len(per_sheet_result) == 1

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for source_sheet, split_map in per_sheet_result:
            for category, part_df in split_map.items():
                if single_source_sheet:
                    target_name = category
                else:
                    target_name = f"{source_sheet}_{category}"

                safe_sheet = build_unique_sheet_name(target_name, used_sheet_names)
                part_df.to_excel(writer, sheet_name=safe_sheet, index=False)

    apply_output_format(output_path)

    print(f"  ✅ 输出完成: {output_path}")
    return True


def apply_output_format(output_path):
    wb = load_workbook(output_path)
    for ws in wb.worksheets:
        APPLY_COMMON_EXCEL_LAYOUT(
            ws,
            numeric_start_col=1,
            amount_headers=AMOUNT_HEADERS,
            header_row=1,
            header_height=40,
            freeze_panes="A2",
        )
    wb.save(output_path)


def main():
    if not os.path.isdir(INPUT_ROOT):
        raise FileNotFoundError(f"输入目录不存在: {INPUT_ROOT}")

    os.makedirs(OUTPUT_ROOT, exist_ok=True)

    files = [
        name for name in sorted(os.listdir(INPUT_ROOT)) if is_excel_file(name)
    ]

    if not files:
        print("未发现可处理Excel文件")
        return

    ok_count = 0
    for name in files:
        input_path = os.path.join(INPUT_ROOT, name)
        output_path = os.path.join(OUTPUT_ROOT, name)

        print(f"\n处理文件: {name}")
        if process_one_workbook(input_path, output_path):
            ok_count += 1

    print("\n" + "=" * 60)
    print("处理结束")
    print(f"成功输出文件数: {ok_count}/{len(files)}")
    print(f"输出目录: {OUTPUT_ROOT}")


if __name__ == "__main__":
    main()
