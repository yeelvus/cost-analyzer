#!/usr.bin/env python3
# -*- coding: utf-8 -*-
"""
项目文件名模式切换工具 v2.4（完整预览修改前后版）
"""

import os
import re
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

_PKG = Path(__file__).resolve().parents[1]
if str(_PKG) not in sys.path:
    sys.path.insert(0, str(_PKG))
from contract_paths import resolve_contract_root

# ==================== 配置区 ====================
EXCEL_PATH = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
HEADER_ROW = 0
CODE_COLUMN = "项目编码"
# 修改这里切换命名模式："项目名称" 或 "项目全称"
NAME_MODE_COLUMN = "项目名称"

ROOT_DIR = str(resolve_contract_root())
LOG_FOLDER = "日志"

# ==================== 工具函数 ====================
def sanitize_folder_name(name):
    if not name or str(name).strip() in ["", "nan", "NaN"]:
        return ""
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name).strip())


def load_projects(excel_path, name_mode_column):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")

    df = pd.read_excel(excel_path, header=HEADER_ROW)
    
    required_cols = [CODE_COLUMN, name_mode_column]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Excel缺少必要列: {missing}")

    df[CODE_COLUMN] = df[CODE_COLUMN].astype(str).str.strip()
    df[name_mode_column] = df[name_mode_column].astype(str).apply(sanitize_folder_name)

    project_map = {}
    code_set = set()

    for _, row in df.iterrows():
        code = row[CODE_COLUMN]
        name = row[name_mode_column]
        if code and code != "nan" and name and name != "nan":
            project_map[code] = name
            code_set.add(code)

    print(f"✅ 从Excel加载 {len(project_map)} 个有效项目映射")
    return project_map, code_set


def find_all_project_folders(root_dir, valid_code_set):
    project_folders = []
    for dirpath, dirnames, _ in os.walk(root_dir):
        for dirname in list(dirnames):
            if dirname.startswith('.'):
                continue
            if '_' not in dirname:
                continue
            prefix = dirname.split('_', 1)[0].strip()
            if prefix in valid_code_set:
                full_path = os.path.join(dirpath, dirname)
                project_folders.append((full_path, dirname))
    return project_folders


def build_rename_plan(project_map, all_folders):
    plan = []
    already_ok = []
    no_code_match = []

    for full_path, current_name in all_folders:
        code = current_name.split('_', 1)[0].strip()
        
        if code not in project_map:
            no_code_match.append((current_name, code))
            continue

        target_name = f"{code}_{project_map[code]}"
        
        if current_name == target_name:
            already_ok.append(current_name)
            continue

        dst_path = os.path.join(os.path.dirname(full_path), target_name)
        
        plan.append({
            "src": full_path,
            "dst": dst_path,
            "src_name": current_name,
            "dst_name": target_name,
            "code": code
        })

    return {
        "rename": plan,
        "already_ok": already_ok,
        "no_code_match": no_code_match
    }


def print_rename_preview(plan):
    """完整打印所有需要修改的文件夹（修改前 → 修改后）"""
    if not plan:
        print("没有需要重命名的文件夹。")
        return

    print(f"\n{'='*80}")
    print(f"即将执行的重命名计划（共 {len(plan)} 个）")
    print(f"{'='*80}")
    
    for i, item in enumerate(plan, 1):
        print(f"{i:3d}. {item['src_name']}")
        print(f"     → {item['dst_name']}")
        print("-" * 70)


def execute_rename(plan):
    renamed = []
    skipped = []
    failed = []

    for item in plan:
        try:
            if not os.path.exists(item["src"]):
                skipped.append(item)
                print(f"⚠️ 跳过（源不存在）: {item['src_name']}")
                continue
            if os.path.exists(item["dst"]):
                skipped.append(item)
                print(f"⚠️ 跳过（目标已存在）: {item['dst_name']}")
                continue

            os.rename(item["src"], item["dst"])
            renamed.append(item)
            print(f"✅ 重命名成功: {item['src_name']} → {item['dst_name']}")

        except Exception as e:
            failed.append(item)
            print(f"❌ 重命名失败: {item['src_name']} → {item['dst_name']}")
            print(f"   错误: {e}")

    return renamed, skipped, failed


def save_log(plan_data, renamed, skipped, failed):
    log_dir = os.path.join(ROOT_DIR, LOG_FOLDER)
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"项目重命名_{timestamp}.json")

    with open(log_file, "w", encoding="utf-8") as f:
        json.dump({
            "timestamp": datetime.now().isoformat(),
            "name_mode": NAME_MODE_COLUMN,
            "to_rename_count": len(plan_data["rename"]),
            "renamed_count": len(renamed),
            "skipped_count": len(skipped),
            "failed_count": len(failed),
            "no_code_match_count": len(plan_data["no_code_match"])
        }, f, ensure_ascii=False, indent=2)

    return log_file


def main():
    print("=" * 80)
    print("🔁 项目文件夹命名模式切换工具 v2.4（完整预览修改前后）")
    print(f"目标模式: 项目编码 + {NAME_MODE_COLUMN}")
    print("=" * 80)

    try:
        project_map, valid_code_set = load_projects(EXCEL_PATH, NAME_MODE_COLUMN)
        
        print(f"\n🔍 正在递归扫描目录: {ROOT_DIR}")
        all_folders = find_all_project_folders(ROOT_DIR, valid_code_set)
        print(f"找到 {len(all_folders)} 个匹配 Excel 项目编码的项目文件夹")

        plan_data = build_rename_plan(project_map, all_folders)

        total_to_rename = len(plan_data["rename"])
        already_ok_count = len(plan_data["already_ok"])
        no_match_count = len(plan_data["no_code_match"])

        print(f"\n匹配成功: {total_to_rename + already_ok_count} 个")
        print(f"需要重命名: {total_to_rename} 个")
        print(f"已是目标格式: {already_ok_count} 个")
        print(f"异常匹配: {no_match_count} 个")

        # ==================== 关键改进：完整预览所有修改 ====================
        if total_to_rename > 0:
            print_rename_preview(plan_data["rename"])
        else:
            print("\n✅ 当前所有项目文件夹已是目标格式，无需修改。")

        # 强制确认
        while True:
            confirm = input("\n是否执行以上重命名操作？(y/n): ").strip().lower()
            if confirm in {"y", "yes", "是"}:
                break
            if confirm in {"n", "no", "否"}:
                print("已取消操作。")
                return
            print("请输入 y 或 n")

        if total_to_rename == 0:
            print("无需执行重命名操作。")
            return

        print("\n🚀 开始执行重命名...")
        renamed, skipped, failed = execute_rename(plan_data["rename"])

        log_file = save_log(plan_data, renamed, skipped, failed)

        print("\n" + "=" * 80)
        print("🎉 处理完成！")
        print(f"成功重命名: {len(renamed)} 个")
        print(f"跳过: {len(skipped)} 个")
        print(f"失败: {len(failed)} 个")
        print(f"日志文件: {log_file}")
        print("=" * 80)

    except Exception as e:
        print(f"\n❌ 程序异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()