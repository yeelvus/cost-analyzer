#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合同目录重构迁移工具
功能：将旧结构（项目状态/项目目录/E_...F_...G_...）迁移到新结构：
      合同类型/项目状态/项目目录名
并可将非合同目录集中提取到历史目录。
"""

import json
import os
import re
import shutil
from datetime import datetime

# ==================== 配置区 ====================
ROOT_DIR = "5_合同管理📑"
LOG_FOLDER = "日志"
CENTRAL_FOLDER_NAME = "历史其他子文件夹集中区"

# 新结构四类合同目录
CONTRACT_TYPES = ["总包合同", "分包合同", "总包签证", "分包签证"]

# 是否将非合同目录集中迁移
MOVE_NON_CONTRACT_TO_CENTRAL = True

# 扫描时忽略的根目录
EXCLUDE_ROOT_FOLDERS = {
    LOG_FOLDER,
    "归档",
    ".git",
    "__pycache__",
    CENTRAL_FOLDER_NAME,
    *CONTRACT_TYPES,
}
# ==============================================


def map_contract_type(folder_name):
    """将旧子目录名映射为新合同类型"""
    text = folder_name.strip()

    # 新命名已存在时直接识别
    if text in CONTRACT_TYPES:
        return text

    if re.match(r"^E_.*_总包合同$", text):
        return "总包合同"
    if re.match(r"^F_.*_分包合同$", text):
        return "分包合同"
    if re.match(r"^F_.*_(总包变更合同\(签证\)|总包签证)$", text):
        return "总包签证"
    if re.match(r"^G_.*_(分包变更合同\(签证\)|分包签证)$", text):
        return "分包签证"

    return None


def unique_path(path):
    """若目标存在，生成不冲突路径"""
    if not os.path.exists(path):
        return path

    index = 1
    while True:
        candidate = f"{path}__dup{index}"
        if not os.path.exists(candidate):
            return candidate
        index += 1


def move_or_merge_folder(src, dst):
    """移动目录；若目标已存在则并入目标目录并保留重名项"""
    os.makedirs(os.path.dirname(dst), exist_ok=True)

    if not os.path.exists(dst):
        shutil.move(src, dst)
        return

    for child in os.listdir(src):
        src_child = os.path.join(src, child)
        dst_child = unique_path(os.path.join(dst, child))
        shutil.move(src_child, dst_child)

    if os.path.isdir(src) and not os.listdir(src):
        os.rmdir(src)


def build_migration_plan(root_dir):
    """
    构建迁移计划
    返回：
    {
      'contract_moves': [...],
      'other_moves': [...]
    }
    """
    if not os.path.exists(root_dir):
        raise FileNotFoundError(f"根目录不存在: {root_dir}")

    contract_moves = []
    other_moves = []

    central_dir = os.path.join(root_dir, CENTRAL_FOLDER_NAME)

    for status in os.listdir(root_dir):
        status_path = os.path.join(root_dir, status)

        if status in EXCLUDE_ROOT_FOLDERS:
            continue
        if not os.path.isdir(status_path):
            continue

        for project_folder in os.listdir(status_path):
            project_path = os.path.join(status_path, project_folder)
            if not os.path.isdir(project_path):
                continue

            for child in os.listdir(project_path):
                child_path = os.path.join(project_path, child)
                if not os.path.isdir(child_path):
                    continue

                contract_type = map_contract_type(child)
                if contract_type:
                    dst = os.path.join(root_dir, contract_type, status, project_folder)
                    contract_moves.append(
                        {
                            "src": child_path,
                            "dst": dst,
                            "status": status,
                            "project": project_folder,
                            "source_folder": child,
                            "contract_type": contract_type,
                        }
                    )
                elif MOVE_NON_CONTRACT_TO_CENTRAL:
                    dst = os.path.join(central_dir, status, project_folder, child)
                    dst = unique_path(dst)
                    other_moves.append(
                        {
                            "src": child_path,
                            "dst": dst,
                            "status": status,
                            "project": project_folder,
                            "source_folder": child,
                        }
                    )

    return {"contract_moves": contract_moves, "other_moves": other_moves}


def execute_plan(plan):
    """执行迁移计划"""
    moved_contract = []
    moved_other = []
    failed = []

    for item in plan["contract_moves"]:
        try:
            move_or_merge_folder(item["src"], item["dst"])
            moved_contract.append(item)
            print(f"✅ 合同迁移: {item['src']} -> {item['dst']}")
        except Exception as exc:
            failed.append({**item, "error": str(exc), "kind": "contract"})
            print(f"❌ 合同迁移失败: {item['src']} -> {item['dst']}")
            print(f"   错误: {exc}")

    for item in plan["other_moves"]:
        try:
            os.makedirs(os.path.dirname(item["dst"]), exist_ok=True)
            shutil.move(item["src"], item["dst"])
            moved_other.append(item)
            print(f"✅ 其他迁移: {item['src']} -> {item['dst']}")
        except Exception as exc:
            failed.append({**item, "error": str(exc), "kind": "other"})
            print(f"❌ 其他迁移失败: {item['src']} -> {item['dst']}")
            print(f"   错误: {exc}")

    return moved_contract, moved_other, failed


def save_log(root_dir, plan, moved_contract, moved_other, failed):
    """保存日志"""
    log_dir = os.path.join(root_dir, LOG_FOLDER)
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"合同目录重构迁移_{timestamp}.json")

    payload = {
        "timestamp": datetime.now().isoformat(),
        "operation": "合同目录重构迁移",
        "structure": "合同类型/项目状态/项目目录名",
        "contract_types": CONTRACT_TYPES,
        "move_non_contract_to_central": MOVE_NON_CONTRACT_TO_CENTRAL,
        "plan_contract_count": len(plan["contract_moves"]),
        "plan_other_count": len(plan["other_moves"]),
        "moved_contract_count": len(moved_contract),
        "moved_other_count": len(moved_other),
        "failed_count": len(failed),
        "plan": plan,
        "moved_contract": moved_contract,
        "moved_other": moved_other,
        "failed": failed,
    }

    with open(log_path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)

    return log_path


def main():
    print("=" * 60)
    print("🔄 合同目录重构迁移工具")
    print("新结构: 合同类型 / 项目状态 / 项目目录名")
    print("=" * 60)

    try:
        plan = build_migration_plan(ROOT_DIR)
        contract_count = len(plan["contract_moves"])
        other_count = len(plan["other_moves"])

        if contract_count == 0 and other_count == 0:
            print("\n✅ 未发现需要迁移的目录")
            return

        print(f"\n📦 合同目录待迁移: {contract_count}")
        print(f"📦 其他目录待迁移: {other_count}")

        preview = 20
        print(f"\n预览前 {preview} 条合同迁移:")
        for idx, item in enumerate(plan["contract_moves"][:preview], 1):
            print(
                f"{idx:>2}. [{item['status']}] {item['project']} | "
                f"{item['source_folder']} -> {item['contract_type']}/{item['project']}"
            )

        if contract_count > preview:
            print(f"... 其余 {contract_count - preview} 条未展示")

        while True:
            confirm = input("\n是否执行迁移操作？(y/n): ").strip().lower()
            if confirm in {"y", "yes", "是"}:
                break
            if confirm in {"n", "no", "否"}:
                print("\n❌ 已取消迁移")
                return
            print("⚠️  请输入 y 或 n")

        print("\n🚀 开始迁移...")
        moved_contract, moved_other, failed = execute_plan(plan)
        log_path = save_log(ROOT_DIR, plan, moved_contract, moved_other, failed)

        print("\n" + "=" * 60)
        print("✅ 迁移完成")
        print(f"合同迁移成功: {len(moved_contract)}")
        print(f"其他迁移成功: {len(moved_other)}")
        print(f"迁移失败: {len(failed)}")
        print(f"日志文件: {log_path}")
        print("=" * 60)

    except Exception as exc:
        print(f"\n❌ 执行失败: {exc}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
