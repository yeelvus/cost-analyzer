#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增量项目文件夹创建工具（合同类型优先版）
功能：根据Excel表格自动补齐合同管理目录，结构为：
      合同类型 / 项目状态 / 项目目录名
作者：自动生成
日期：2026-04-07
"""

import json
import os
import re
from datetime import datetime

import pandas as pd

# ==================== 配置区 ====================
# Excel配置
EXCEL_PATH = "0_项目基本信息情况/项目基本信息表.xlsx"
HEADER_ROW = 0

# 列名配置
CODE_COLUMN = "项目编码"
NAME_COLUMN = "项目全称"
STATUS_COLUMN = "项目状态"

# 根目录（相对于脚本执行位置）
ROOT_DIR = "5_合同管理"

# 合同类型目录（固定4类）
CONTRACT_TYPES = [
    "总包合同",
    "分包合同",
    "总包签证",
    "分包签证",
]

# 日志文件夹
LOG_FOLDER = "日志"
# ==============================================


def sanitize_folder_name(name):
    """清理文件夹名称中的非法字符"""
    if not name or str(name).strip() in ["", "nan", "NaN"]:
        return ""
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name).strip())


def load_projects_from_excel(excel_path):
    """从Excel加载项目列表"""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")

    df = pd.read_excel(excel_path, header=HEADER_ROW)

    required_cols = [CODE_COLUMN, NAME_COLUMN, STATUS_COLUMN]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel缺少必要列: {missing_cols}")

    df[CODE_COLUMN] = df[CODE_COLUMN].astype(str).str.strip()
    df[NAME_COLUMN] = df[NAME_COLUMN].astype(str).apply(sanitize_folder_name)
    df[STATUS_COLUMN] = df[STATUS_COLUMN].astype(str).apply(sanitize_folder_name)
    df[STATUS_COLUMN] = df[STATUS_COLUMN].str.replace(r"\s+", "", regex=True)

    df = df[
        (df[CODE_COLUMN] != "")
        & (df[CODE_COLUMN] != "nan")
        & (df[NAME_COLUMN] != "")
        & (df[NAME_COLUMN] != "nan")
        & (df[STATUS_COLUMN] != "")
        & (df[STATUS_COLUMN] != "nan")
    ]

    return df


def resolve_contract_type_dirs(root_dir, contract_types):
    """
    解析合同类型目录的实际名称，支持前缀命名。
    规则：优先完全匹配，其次后缀匹配，再次包含匹配；都未命中则使用逻辑名。
    """
    resolved = {}
    candidates = []

    if os.path.exists(root_dir):
        for name in os.listdir(root_dir):
            path = os.path.join(root_dir, name)
            if os.path.isdir(path) and not name.startswith('.'):
                candidates.append(name)

    for logical_name in contract_types:
        exact = [name for name in candidates if name == logical_name]
        if exact:
            resolved[logical_name] = exact[0]
            continue

        suffix = sorted([name for name in candidates if name.endswith(logical_name)])
        if suffix:
            resolved[logical_name] = suffix[0]
            continue

        contains = sorted([name for name in candidates if logical_name in name])
        if contains:
            resolved[logical_name] = contains[0]
            continue

        resolved[logical_name] = logical_name

    return resolved


def detect_missing_contract_dirs(df, contract_dir_map):
    """
    检测需要补齐的目录
    返回: list[{
        'code': ...,
        'project_folder': ...,
        'status': ...,
        'missing_contract_paths': [...],
    }]
    """
    plan = []

    for _, row in df.iterrows():
        code = row[CODE_COLUMN]
        name = row[NAME_COLUMN]
        status = row[STATUS_COLUMN]
        project_folder = f"{code}_{name}"

        missing = []
        for contract_type in CONTRACT_TYPES:
            contract_dir = contract_dir_map.get(contract_type, contract_type)
            expected_path = os.path.join(ROOT_DIR, contract_dir, status, project_folder)
            if not os.path.exists(expected_path):
                missing.append(expected_path)

        if missing:
            plan.append(
                {
                    "code": code,
                    "name": name,
                    "status": status,
                    "project_folder": project_folder,
                    "missing_contract_paths": missing,
                }
            )

    return plan


def display_plan(plan):
    """显示待创建目录计划"""
    if not plan:
        print("\n" + "=" * 60)
        print("✅ 所有项目合同目录已完整，无需操作！")
        print("=" * 60)
        return False

    print("\n" + "=" * 60)
    print(f"📂 检测到 {len(plan)} 个项目需要补齐合同目录")
    print("目录结构: 合同类型 / 项目状态 / 项目目录名")
    print("=" * 60)

    for idx, item in enumerate(plan, 1):
        print(f"\n项目 #{idx}:")
        print(f"   编码: {item['code']}")
        print(f"   目录名: {item['project_folder']}")
        print(f"   状态: {item['status']}")
        print(f"   待创建: {len(item['missing_contract_paths'])} 个")
        for path in item["missing_contract_paths"]:
            print(f"      - {path}")

    print("\n" + "=" * 60)
    return True


def create_dirs(plan):
    """执行目录创建"""
    created = []

    for item in plan:
        for path in item["missing_contract_paths"]:
            os.makedirs(path, exist_ok=True)
            created.append(path)
            print(f"✅ 创建: {path}")

    return created


def save_log(plan, created):
    """保存执行日志"""
    log_dir = os.path.join(ROOT_DIR, LOG_FOLDER)
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"增量创建_合同类型优先_{timestamp}.json")

    payload = {
        "timestamp": datetime.now().isoformat(),
        "operation": "增量合同目录创建（合同类型优先）",
        "structure": "合同类型/项目状态/项目目录名",
        "contract_types": CONTRACT_TYPES,
        "project_count": len(plan),
        "folders_created": len(created),
        "plan": plan,
        "created_folders": created,
    }

    with open(log_file, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)

    return log_file


def main():
    """主函数"""
    print("=" * 60)
    print("🚀 增量项目文件夹创建工具（合同类型优先版）")
    print("=" * 60)

    try:
        print(f"\n📊 读取Excel: {EXCEL_PATH}")
        df = load_projects_from_excel(EXCEL_PATH)
        print(f"   ✅ 共加载 {len(df)} 个项目")

        contract_dir_map = resolve_contract_type_dirs(ROOT_DIR, CONTRACT_TYPES)
        print("\n🧭 合同目录识别结果:")
        for logical_name in CONTRACT_TYPES:
            print(f"   - {logical_name} -> {contract_dir_map.get(logical_name, logical_name)}")

        print("\n🔍 检测待补齐目录...")
        plan = detect_missing_contract_dirs(df, contract_dir_map)

        has_work = display_plan(plan)
        if not has_work:
            return

        while True:
            response = input("是否执行以上操作？(y/n): ").strip().lower()
            if response in ["y", "yes", "是"]:
                break
            if response in ["n", "no", "否"]:
                print("\n❌ 操作已取消")
                return
            print("⚠️  请输入 y 或 n")

        print("\n📁 开始创建目录...")
        print("-" * 60)
        created = create_dirs(plan)
        log_file = save_log(plan, created)

        print("\n" + "=" * 60)
        print("✅ 操作完成！")
        print(f"   涉及项目: {len(plan)}")
        print(f"   创建目录总数: {len(created)}")
        print(f"   操作日志: {log_file}")
        print("=" * 60)

    except FileNotFoundError as exc:
        print(f"\n❌ 错误: {exc}")
        print("💡 提示: 请确保在 01成本统计 目录下运行此脚本")
    except ValueError as exc:
        print(f"\n❌ 错误: {exc}")
    except Exception as exc:
        print(f"\n❌ 未知错误: {exc}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
