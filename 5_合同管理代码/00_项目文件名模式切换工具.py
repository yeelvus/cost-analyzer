#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目文件名模式切换工具
功能：根据项目基本信息表，批量将目录名切换为：项目编码_项目全称 或 项目编码_项目名称
目录结构：合同类型 / 项目状态 / 项目目录名
"""

import json
import os
import re
from datetime import datetime

import pandas as pd

# ==================== 配置区 ====================
EXCEL_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
HEADER_ROW = 0

CODE_COLUMN = "项目编码"
STATUS_COLUMN = "项目状态"

# 可选值："项目全称" 或 "项目名称"
NAME_MODE_COLUMN = "项目全称"

ROOT_DIR = "5_合同管理📑"
CONTRACT_TYPES = ["总包合同", "分包合同", "总包签证", "分包签证"]
LOG_FOLDER = "日志"
# ==============================================


def sanitize_folder_name(name):
    if not name or str(name).strip() in ["", "nan", "NaN"]:
        return ""
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name).strip())


def resolve_contract_type_dirs(root_dir, contract_types):
    """支持目录前缀：优先完全匹配，其次后缀匹配，再次包含匹配。"""
    resolved = {}
    candidates = []

    if os.path.exists(root_dir):
        for name in os.listdir(root_dir):
            path = os.path.join(root_dir, name)
            if os.path.isdir(path) and not name.startswith('.'):
                candidates.append(name)

    for logical_name in contract_types:
        exact = [n for n in candidates if n == logical_name]
        if exact:
            resolved[logical_name] = exact[0]
            continue

        suffix = sorted([n for n in candidates if n.endswith(logical_name)])
        if suffix:
            resolved[logical_name] = suffix[0]
            continue

        contains = sorted([n for n in candidates if logical_name in n])
        if contains:
            resolved[logical_name] = contains[0]
            continue

        resolved[logical_name] = logical_name

    return resolved


def load_projects(excel_path, name_mode_column):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")

    df = pd.read_excel(excel_path, header=HEADER_ROW)

    required_cols = [CODE_COLUMN, STATUS_COLUMN, name_mode_column]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Excel缺少必要列: {missing}")

    df[CODE_COLUMN] = df[CODE_COLUMN].astype(str).str.strip()
    df[STATUS_COLUMN] = df[STATUS_COLUMN].astype(str).apply(sanitize_folder_name)
    df[STATUS_COLUMN] = df[STATUS_COLUMN].str.replace(r"\s+", "", regex=True)
    df[name_mode_column] = df[name_mode_column].astype(str).apply(sanitize_folder_name)

    df = df[
        (df[CODE_COLUMN] != "")
        & (df[CODE_COLUMN] != "nan")
        & (df[STATUS_COLUMN] != "")
        & (df[STATUS_COLUMN] != "nan")
        & (df[name_mode_column] != "")
        & (df[name_mode_column] != "nan")
    ]

    return df


def find_project_folder_by_code(status_path, code):
    """在状态目录下按项目编码查找项目目录（格式：编码_名称）。"""
    if not os.path.exists(status_path):
        return None

    prefix = f"{code}_"
    for name in os.listdir(status_path):
        full_path = os.path.join(status_path, name)
        if not os.path.isdir(full_path):
            continue
        if name.startswith(prefix):
            return name
    return None


def build_rename_plan(df, contract_dir_map):
    """
    生成重命名计划
    返回：
    {
      'rename': [...],
      'missing': [...],
      'already_ok': [...]
    }
    """
    plan_rename = []
    plan_missing = []
    plan_ok = []

    for _, row in df.iterrows():
        code = row[CODE_COLUMN]
        status = row[STATUS_COLUMN]
        name = row[NAME_MODE_COLUMN]
        expected_folder = f"{code}_{name}"

        for logical_type in CONTRACT_TYPES:
            actual_type = contract_dir_map.get(logical_type, logical_type)
            status_path = os.path.join(ROOT_DIR, actual_type, status)

            existing = find_project_folder_by_code(status_path, code)
            if not existing:
                plan_missing.append(
                    {
                        "contract_type": actual_type,
                        "status": status,
                        "code": code,
                        "expected": expected_folder,
                    }
                )
                continue

            if existing == expected_folder:
                plan_ok.append(
                    {
                        "contract_type": actual_type,
                        "status": status,
                        "code": code,
                        "folder": existing,
                    }
                )
                continue

            src = os.path.join(status_path, existing)
            dst = os.path.join(status_path, expected_folder)
            plan_rename.append(
                {
                    "contract_type": actual_type,
                    "status": status,
                    "code": code,
                    "src_name": existing,
                    "dst_name": expected_folder,
                    "src": src,
                    "dst": dst,
                }
            )

    return {"rename": plan_rename, "missing": plan_missing, "already_ok": plan_ok}


def execute_rename_plan(plan):
    renamed = []
    skipped = []
    failed = []

    for item in plan:
        src = item["src"]
        dst = item["dst"]

        try:
            if not os.path.exists(src):
                skipped.append({**item, "reason": "源目录不存在"})
                print(f"⚠️ 跳过（源不存在）: {src}")
                continue
            if os.path.exists(dst):
                skipped.append({**item, "reason": "目标目录已存在"})
                print(f"⚠️ 跳过（目标已存在）: {dst}")
                continue

            os.rename(src, dst)
            renamed.append(item)
            print(f"✅ 重命名: {item['src_name']} -> {item['dst_name']}")
        except Exception as exc:
            failed.append({**item, "error": str(exc)})
            print(f"❌ 失败: {src} -> {dst}")
            print(f"   错误: {exc}")

    return renamed, skipped, failed


def save_log(plan_data, renamed, skipped, failed):
    log_dir = os.path.join(ROOT_DIR, LOG_FOLDER)
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"项目文件名模式切换_{timestamp}.json")

    payload = {
        "timestamp": datetime.now().isoformat(),
        "operation": "项目文件名模式切换",
        "name_mode_column": NAME_MODE_COLUMN,
        "rename_plan_count": len(plan_data["rename"]),
        "missing_count": len(plan_data["missing"]),
        "already_ok_count": len(plan_data["already_ok"]),
        "renamed_count": len(renamed),
        "skipped_count": len(skipped),
        "failed_count": len(failed),
        "plan": plan_data,
        "renamed": renamed,
        "skipped": skipped,
        "failed": failed,
    }

    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return log_file


def main():
    print("=" * 68)
    print("🔁 项目文件名模式切换工具")
    print("模式: 项目编码 + 指定名称列")
    print("=" * 68)

    try:
        print(f"\n📊 读取Excel: {EXCEL_PATH}")
        print(f"🧭 当前名称模式列: {NAME_MODE_COLUMN}")
        df = load_projects(EXCEL_PATH, NAME_MODE_COLUMN)
        print(f"✅ 有效项目数: {len(df)}")

        contract_dir_map = resolve_contract_type_dirs(ROOT_DIR, CONTRACT_TYPES)
        print("\n📁 合同目录识别:")
        for name in CONTRACT_TYPES:
            print(f"   - {name} -> {contract_dir_map.get(name, name)}")

        plan_data = build_rename_plan(df, contract_dir_map)
        rename_count = len(plan_data["rename"])
        missing_count = len(plan_data["missing"])
        ok_count = len(plan_data["already_ok"])

        print("\n" + "-" * 68)
        print(f"待重命名: {rename_count}")
        print(f"无需处理: {ok_count}")
        print(f"未找到目录: {missing_count}")

        preview = min(20, rename_count)
        if preview > 0:
            print(f"\n预览前 {preview} 条重命名:")
            for i, item in enumerate(plan_data["rename"][:preview], 1):
                print(
                    f"{i:>2}. [{item['contract_type']}/{item['status']}] "
                    f"{item['src_name']} -> {item['dst_name']}"
                )

        if rename_count == 0:
            print("\n✅ 没有需要重命名的目录")
            return

        while True:
            confirm = input("\n是否执行以上重命名？(y/n): ").strip().lower()
            if confirm in {"y", "yes", "是"}:
                break
            if confirm in {"n", "no", "否"}:
                print("\n❌ 已取消")
                return
            print("⚠️ 请输入 y 或 n")

        print("\n🚀 开始执行重命名...")
        renamed, skipped, failed = execute_rename_plan(plan_data["rename"])
        log_file = save_log(plan_data, renamed, skipped, failed)

        print("\n" + "=" * 68)
        print("✅ 执行完成")
        print(f"重命名成功: {len(renamed)}")
        print(f"跳过: {len(skipped)}")
        print(f"失败: {len(failed)}")
        print(f"日志文件: {log_file}")
        print("=" * 68)

    except Exception as exc:
        print(f"\n❌ 执行失败: {exc}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
