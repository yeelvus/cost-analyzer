#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
签证空文件夹递归清理工具（增强版）
功能：针对「总包签证」和「分包签证」目录，
     递归删除所有空文件夹（支持多层级级联删除）
     特别支持：当某个状态文件夹（如 D_1_已完工...）下所有文件被删除后，
     自动删除该空的状态文件夹。
作者：根据最新需求优化
日期：2026-04-24
"""

import os
import json
from datetime import datetime

# ==================== 配置区 ====================
ROOT_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑"

VISA_CONTRACT_TYPES = [
    "总包签证",
    "分包签证",
]

LOG_FOLDER = "日志"
DRY_RUN = False                    # 第一次使用强烈建议设为 True
MIN_DEPTH = 1                      # 改为1，允许删除状态层文件夹（D_xxx）
MAX_ITERATIONS = 15                # 最大轮数，防止意外
# ==============================================


def resolve_contract_dirs(root_dir, logical_types):
    """解析实际的合同类型文件夹名称"""
    resolved = {}
    if not os.path.exists(root_dir):
        return resolved

    existing = [name for name in os.listdir(root_dir)
                if os.path.isdir(os.path.join(root_dir, name)) and not name.startswith('.')]

    for logical in logical_types:
        for priority in [
            lambda d: d == logical,
            lambda d: d.endswith(logical),
            lambda d: logical in d
        ]:
            matches = [d for d in existing if priority(d)]
            if matches:
                resolved[logical] = sorted(matches)[-1]
                break
        else:
            resolved[logical] = logical
    return resolved


def is_empty_dir(path: str) -> bool:
    """判断文件夹是否为空"""
    try:
        return os.path.isdir(path) and len(os.listdir(path)) == 0
    except:
        return False


def find_and_delete_empty_dirs(contract_path, dry_run=False, min_depth=1):
    """单轮扫描 + 删除，返回本次删除列表"""
    deleted_this_round = []

    for dirpath, dirnames, filenames in os.walk(contract_path, topdown=False):
        # 计算相对深度
        rel_depth = dirpath.count(os.sep) - contract_path.count(os.sep)
        if rel_depth < min_depth:
            continue

        if not filenames and not dirnames and is_empty_dir(dirpath):
            try:
                if dry_run:
                    print(f"🧪 [模拟] 将删除: {dirpath}")
                else:
                    os.rmdir(dirpath)
                    print(f"✅ 已删除: {dirpath}")
                deleted_this_round.append(dirpath)
            except OSError as e:
                print(f"❌ 删除失败 {dirpath}: {e}")

    return deleted_this_round


def main():
    print("=" * 85)
    print("🧹 签证空文件夹递归清理工具（支持状态层级删除）")
    print("   → 当状态文件夹（如 D_1_已完工...）变空后会自动删除")
    print("=" * 85)

    try:
        contract_dir_map = resolve_contract_dirs(ROOT_DIR, VISA_CONTRACT_TYPES)
        print("\n📂 合同目录映射：")
        for logical, actual in contract_dir_map.items():
            print(f"   {logical}  →  {actual}")

        all_deleted = []
        total_iterations = 0

        print("\n🔍 开始多轮递归清理...")

        while total_iterations < MAX_ITERATIONS:
            total_iterations += 1
            print(f"\n第 {total_iterations} 轮扫描:")

            iteration_deleted = []
            for logical, actual in contract_dir_map.items():
                contract_path = os.path.join(ROOT_DIR, actual)
                if not os.path.exists(contract_path):
                    print(f"⚠️  未找到目录: {actual}")
                    continue

                deleted = find_and_delete_empty_dirs(
                    contract_path, dry_run=DRY_RUN, min_depth=MIN_DEPTH
                )
                iteration_deleted.extend(deleted)

            all_deleted.extend(iteration_deleted)

            if not iteration_deleted:
                if total_iterations == 1:
                    print("\n✅  未发现任何空文件夹！")
                else:
                    print(f"\n✅  清理完成！共执行 {total_iterations} 轮。")
                break

        # 保存日志
        log_dir = os.path.join(ROOT_DIR, LOG_FOLDER)
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        mode = "模拟" if DRY_RUN else "实际"
        log_file = os.path.join(log_dir, f"签证空文件夹递归清理_{mode}_{timestamp}.json")

        payload = {
            "timestamp": datetime.now().isoformat(),
            "operation": "签证空文件夹递归清理（支持状态文件夹删除）",
            "dry_run": DRY_RUN,
            "iterations": total_iterations,
            "total_deleted": len(all_deleted),
            "deleted_folders": all_deleted,
        }

        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        print("\n" + "=" * 85)
        print("🎉 操作完成！")
        print(f"   总扫描轮数: {total_iterations}")
        print(f"   共清理空文件夹: {len(all_deleted)} 个")
        print(f"   模式: {'模拟运行' if DRY_RUN else '真实删除'}")
        print(f"   日志文件: {log_file}")
        print("=" * 85)

    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()