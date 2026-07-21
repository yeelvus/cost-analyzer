# -*- coding: utf-8 -*-
"""
合同金额汇总（原 03.1-总包分包合同额统计）
按项目汇总总包/分包合同额（万泰铢），并按币种拆分 Sheet。
"""

from __future__ import annotations

import json
import os
import re
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook

from summary_config import (
    GENERAL_CONTRACT_DIR,
    SUBPACKAGE_CONTRACT_DIR,
    get_contract_root,
    get_project_db,
)


def load_project_name_by_code(match_db_file: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not os.path.exists(match_db_file):
        print(f"警告: 项目库文件不存在: {match_db_file}")
        return mapping
    try:
        with open(match_db_file, "r", encoding="utf-8") as f:
            loaded_data = json.load(f)
        projects = loaded_data.get("projects", {}) if isinstance(loaded_data, dict) else {}
        for project_name, project_info in projects.items():
            if not isinstance(project_info, dict):
                continue
            code = str(project_info.get("project_code", "")).strip()
            name = str(project_name).strip()
            if not code or not name:
                continue
            mapping.setdefault(code, name)
        print(f"已加载项目编码映射: {len(mapping)} 个 project_code")
    except Exception as e:
        print(f"警告: 读取项目库失败: {e}")
    return mapping


def extract_contract_amount_from_json(json_file: str) -> tuple[float, str, float, float]:
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        amount_raw = data.get("合同金额")
        amount = float(amount_raw) if amount_raw not in (None, "", "null", "0") else 0.0

        currency = str(data.get("币种", "")).strip() or "THB"

        rate_raw = data.get("汇率")
        if rate_raw in (None, "", "null", 0, "0", 0.0):
            exchange_rate = 1.0
            print(
                f"  ⚠️  汇率缺失或为0，已默认设为1.0 → {os.path.basename(os.path.dirname(json_file))}"
            )
        else:
            try:
                exchange_rate = float(rate_raw)
                if exchange_rate <= 0:
                    exchange_rate = 1.0
                    print(
                        f"  ⚠️  汇率为负数或0，已默认设为1.0 → {os.path.basename(os.path.dirname(json_file))}"
                    )
            except (ValueError, TypeError):
                exchange_rate = 1.0
                print(
                    f"  ⚠️  汇率格式错误，已默认设为1.0 → {os.path.basename(os.path.dirname(json_file))}"
                )

        # 与明细汇总保持一致：CNY 且汇率无效时使用默认 5（仅当 JSON 无有效汇率）
        # 此处保留原 03.1 逻辑：默认 1.0；金额换算统一用 exchange_rate
        amount_in_wan_thb = (amount * exchange_rate) / 10000
        return amount, currency, exchange_rate, amount_in_wan_thb
    except Exception as e:
        print(f"  ❌ 读取JSON失败: {json_file} | 错误: {e}")
        return 0.0, "THB", 1.0, 0.0


def calculate_contract_total(contract_folder: str) -> tuple[float, int, set, dict]:
    total_thb = 0.0
    contract_count = 0
    currencies_used: set[str] = set()
    currency_details = {"CNY": 0.0, "THB": 0.0}

    if not os.path.exists(contract_folder):
        return total_thb, contract_count, currencies_used, currency_details

    for item in os.listdir(contract_folder):
        item_path = os.path.join(contract_folder, item)
        if not os.path.isdir(item_path) or item.startswith("."):
            continue
        json_file = os.path.join(item_path, "合同信息.json")
        if os.path.exists(json_file):
            amount, currency, exchange_rate, amount_in_wan_thb = extract_contract_amount_from_json(
                json_file
            )
            total_thb += amount_in_wan_thb
            contract_count += 1
            currencies_used.add(currency)
            if currency == "CNY":
                currency_details["CNY"] += amount_in_wan_thb
            else:
                currency_details["THB"] += amount_in_wan_thb

    return total_thb, contract_count, currencies_used, currency_details


def extract_project_code(project_dir_name: str) -> str | None:
    match = re.search(r"(THCC-\d{4}-\d{3}|LSBM-\d+)", project_dir_name)
    if match:
        return match.group(1)
    # 通用：编码_名称
    prefix = project_dir_name.split("_", 1)[0].strip()
    if re.match(r"^[A-Za-z]{2,}-\d{4}-\d+$", prefix) or re.match(r"^[A-Za-z]+-\d+$", prefix):
        return prefix
    return None


def resolve_contract_dir_name(base_dir: str, logical_name: str) -> str:
    exact_path = os.path.join(base_dir, logical_name)
    if os.path.isdir(exact_path):
        return logical_name

    candidates = []
    try:
        for name in os.listdir(base_dir):
            path = os.path.join(base_dir, name)
            if os.path.isdir(path) and not name.startswith("."):
                candidates.append(name)
    except OSError:
        return logical_name

    suffix_matches = sorted([name for name in candidates if name.endswith(logical_name)])
    if suffix_matches:
        return suffix_matches[0]
    contains_matches = sorted([name for name in candidates if logical_name in name])
    if contains_matches:
        return contains_matches[0]
    return logical_name


def collect_contract_project_paths(base_dir: str, contract_dir_name: str) -> dict:
    result = {}
    contract_root = os.path.join(base_dir, contract_dir_name)
    if not os.path.exists(contract_root):
        return result

    for status_dir in os.listdir(contract_root):
        status_path = os.path.join(contract_root, status_dir)
        if not os.path.isdir(status_path) or status_dir.startswith("."):
            continue
        for project_dir in os.listdir(status_path):
            project_path = os.path.join(status_path, project_dir)
            if not os.path.isdir(project_path) or project_dir.startswith("."):
                continue
            project_code = extract_project_code(project_dir)
            if not project_code:
                continue
            result[(status_dir, project_dir, project_code)] = project_path
    return result


def scan_all_projects(
    base_dir: str | None = None,
    project_name_by_code: dict[str, str] | None = None,
    general_contract_dir: str | None = None,
    subpackage_contract_dir: str | None = None,
) -> list[dict[str, Any]]:
    base_dir = base_dir or get_contract_root()
    if project_name_by_code is None:
        project_name_by_code = load_project_name_by_code(get_project_db())

    general_contract_dir = general_contract_dir or resolve_contract_dir_name(
        base_dir, GENERAL_CONTRACT_DIR
    )
    subpackage_contract_dir = subpackage_contract_dir or resolve_contract_dir_name(
        base_dir, SUBPACKAGE_CONTRACT_DIR
    )

    print("\n开始扫描项目文件夹...")
    print(f"总包目录: {general_contract_dir}")
    print(f"分包目录: {subpackage_contract_dir}")

    projects_data: list[dict[str, Any]] = []
    general_paths = collect_contract_project_paths(base_dir, general_contract_dir)
    subpackage_paths = collect_contract_project_paths(base_dir, subpackage_contract_dir)
    all_project_keys = sorted(set(general_paths.keys()) | set(subpackage_paths.keys()))

    for project_status, project_dir, project_code in all_project_keys:
        standard_name = project_name_by_code.get(project_code, project_dir)
        name_source = "project_code_db" if project_code in project_name_by_code else "folder_name_fallback"

        general_total, general_count = 0.0, 0
        subpackage_total, subpackage_count = 0.0, 0
        general_currencies: set[str] = set()
        subpackage_currencies: set[str] = set()

        general_path = general_paths.get((project_status, project_dir, project_code))
        if general_path:
            amount, count, currencies, _details = calculate_contract_total(general_path)
            general_total += amount
            general_count += count
            general_currencies.update(currencies)

        subpackage_path = subpackage_paths.get((project_status, project_dir, project_code))
        if subpackage_path:
            amount, count, currencies, _details = calculate_contract_total(subpackage_path)
            subpackage_total += amount
            subpackage_count += count
            subpackage_currencies.update(currencies)

        if general_count > 0 or subpackage_count > 0:
            all_currencies = general_currencies | subpackage_currencies
            if "CNY" in all_currencies and "THB" in all_currencies:
                currency_type = "MIXED"
            elif "CNY" in all_currencies:
                currency_type = "CNY"
            else:
                currency_type = "THB"

            projects_data.append(
                {
                    "项目状态": project_status,
                    "项目编码": project_code,
                    "项目名称": standard_name,
                    "总包合同额(万泰铢)": round(general_total, 2),
                    "分包合同额(万泰铢)": round(subpackage_total, 2),
                    "_currency_type": currency_type,
                }
            )
            currency_label = {
                "CNY": "[人民币]",
                "THB": "[泰铢]",
                "MIXED": "[混合币种]",
            }.get(currency_type, "")
            print(
                f" {project_code} {currency_label}: 总包 {general_count}个({general_total:,.2f}万) "
                f"+ 分包 {subpackage_count}个({subpackage_total:,.2f}万) [名称匹配:{name_source}]"
            )

    print(f"\n共找到 {len(projects_data)} 个项目")
    return projects_data


def build_amount_dataframes(projects_data: list[dict]) -> dict[str, pd.DataFrame]:
    """构建金额汇总各 Sheet 的 DataFrame。"""
    if not projects_data:
        return {"说明": pd.DataFrame([{"提示": "未找到任何合同金额数据"}])}

    df_all = pd.DataFrame(projects_data)
    display_columns = ["项目状态", "项目编码", "项目名称", "总包合同额(万泰铢)", "分包合同额(万泰铢)"]
    sheets: dict[str, pd.DataFrame] = {}

    def with_total(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        total_row = {
            "项目状态": "",
            "项目编码": "",
            "项目名称": "总计",
            "总包合同额(万泰铢)": df["总包合同额(万泰铢)"].sum(),
            "分包合同额(万泰铢)": df["分包合同额(万泰铢)"].sum(),
        }
        return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    df_sheet1 = with_total(
        df_all[display_columns].copy().sort_values(by="项目编码", ascending=False)
    )
    sheets["1_总包分包汇总(全部)"] = df_sheet1

    df_cny = df_all[df_all["_currency_type"] == "CNY"]
    if not df_cny.empty:
        sheets["2_人民币项目(CNY)"] = with_total(
            df_cny[display_columns].copy().sort_values(by="项目编码", ascending=False)
        )

    df_thb = df_all[df_all["_currency_type"] == "THB"]
    if not df_thb.empty:
        sheets["3_泰铢项目(THB)"] = with_total(
            df_thb[display_columns].copy().sort_values(by="项目编码", ascending=False)
        )

    df_mixed = df_all[df_all["_currency_type"] == "MIXED"]
    if not df_mixed.empty:
        sheets["4_混合币种项目"] = with_total(
            df_mixed[display_columns].copy().sort_values(by="项目编码", ascending=False)
        )

    return sheets


def write_amount_sheets(writer: pd.ExcelWriter, sheets: dict[str, pd.DataFrame]) -> list[str]:
    written = []
    for name, df in sheets.items():
        safe = name[:31]
        df.to_excel(writer, index=False, sheet_name=safe)
        written.append(safe)
    return written


def format_amount_sheets(workbook: Workbook, sheet_names: list[str] | None = None) -> None:
    targets = sheet_names if sheet_names is not None else list(workbook.sheetnames)
    for sheet_name in targets:
        if sheet_name not in workbook.sheetnames:
            continue
        # 仅格式化金额类 sheet（避免破坏明细 sheet）
        if not (
            sheet_name.startswith("1_")
            or sheet_name.startswith("2_")
            or sheet_name.startswith("3_")
            or sheet_name.startswith("4_")
            or sheet_name == "说明"
        ):
            # 若传入明确列表则仍处理
            if sheet_names is None:
                continue

        ws = workbook[sheet_name]
        if ws.max_row < 1:
            continue
        if ws.max_row == 1 and ws["A1"].value and "提示" in str(ws["A1"].value):
            continue

        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        total_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        total_font = Font(bold=True, size=11)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        ws.row_dimensions[1].height = 40
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

        last_row = ws.max_row
        # 仅当最后一行是「总计」时高亮
        name_cell = ws.cell(row=last_row, column=3).value
        if name_cell == "总计":
            for col_idx, cell in enumerate(ws[last_row], start=1):
                cell.fill = total_fill
                cell.font = total_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if col_idx <= 3 else "right", vertical="center"
                )

        for col_idx in [4, 5]:
            if col_idx > ws.max_column:
                continue
            col_letter = get_column_letter(col_idx)
            for cell in ws[col_letter][1:]:
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"

        column_widths = {"A": 20, "B": 18, "C": 50, "D": 18, "E": 18}
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions


def export_amount_only(output_path: str, base_dir: str | None = None) -> tuple[str, dict]:
    projects_data = scan_all_projects(base_dir=base_dir)
    if not projects_data:
        print("\n未找到任何合同数据!")
        return output_path, {}

    sheets = build_amount_dataframes(projects_data)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        written = write_amount_sheets(writer, sheets)

    wb = load_workbook(output_path)
    format_amount_sheets(wb, written)
    wb.save(output_path)
    wb.close()

    df_all = pd.DataFrame(projects_data)
    stats = {
        "all": len(df_all),
        "cny": int((df_all["_currency_type"] == "CNY").sum()),
        "thb": int((df_all["_currency_type"] == "THB").sum()),
        "mixed": int((df_all["_currency_type"] == "MIXED").sum()),
        "total_general": float(df_all["总包合同额(万泰铢)"].sum()),
        "total_subpackage": float(df_all["分包合同额(万泰铢)"].sum()),
    }
    print(f"金额汇总已保存：{output_path}")
    return output_path, stats


def main() -> None:
    print("=" * 80)
    print("合同管理 - 总包分包合同额汇总工具")
    print("=" * 80)
    root = get_contract_root()
    print(f"\n合同目录: {root}")
    if not os.path.exists(root):
        print(f"\n错误: 合同目录不存在: {root}")
        return
    from summary_config import AMOUNT_OUTPUT_NAME

    output_path = os.path.join(root, AMOUNT_OUTPUT_NAME)
    _path, stats = export_amount_only(output_path, base_dir=root)
    if stats:
        print("\n" + "=" * 80)
        print("✅ 汇总完成!")
        print(f" 项目总数: {stats['all']}")
        print(f" - 人民币项目: {stats['cny']}")
        print(f" - 泰铢项目: {stats['thb']}")
        print(f" - 混合币种项目: {stats['mixed']}")
        print(f"\n 总包合同额: {stats['total_general']:,.2f} 万泰铢")
        print(f" 分包合同额: {stats['total_subpackage']:,.2f} 万泰铢")


if __name__ == "__main__":
    main()
