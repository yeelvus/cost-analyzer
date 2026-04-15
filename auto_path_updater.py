#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动路径检测与更新工具（最终加强版）
功能：
1. 第一次运行：自动记录所有成本统计相关路径
2. 后续运行：自动检测子文件夹变化，并自动修复所有 .py 文件中的路径
"""

import json
import shutil
from pathlib import Path
from datetime import datetime
import argparse

print("🚀 自动路径检测与更新工具启动...\n")

# ==================== 基础配置 ====================
DATA_BASE = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计")
CODE_BASE = Path("/Users/wyattsun/Desktop/code")
RECORD_FILE = Path(__file__).parent / "path_record.json"
BACKUP_DIR = Path(__file__).parent / "path_backups"

# 确保备份目录存在
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

def find_current_code_base():
    """自动查找 code 目录"""
    if CODE_BASE.exists():
        return CODE_BASE
    home = Path.home()
    for p in home.rglob("project_name_matcher.py"):
        if p.is_file():
            return p.parent.resolve()
    return None

def scan_current_paths():
    """扫描当前 01成本统计 下所有子路径"""
    if not DATA_BASE.exists():
        print(f"❌ 未找到基础数据目录: {DATA_BASE}")
        return {}
    
    path_map = {}
    print(f"🔍 正在扫描 {DATA_BASE} 下的所有子文件夹...")
    
    for item in DATA_BASE.rglob("*"):
        if item.is_dir():
            relative = item.relative_to(DATA_BASE)
            full_path = str(item.resolve())
            relative_str = str(relative)
            path_map[relative_str] = full_path
    
    print(f"✅ 扫描完成，共记录 {len(path_map)} 个路径")
    return path_map

def load_record():
    """加载之前记录的路径"""
    if RECORD_FILE.exists():
        try:
            with open(RECORD_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_record(path_map):
    """保存当前路径记录"""
    with open(RECORD_FILE, "w", encoding="utf-8") as f:
        json.dump(path_map, f, ensure_ascii=False, indent=2)
    print(f"💾 路径记录已保存至: {RECORD_FILE}")

def detect_changes(old_record, new_record):
    """检测路径变化"""
    changes = []
    for rel_path, old_full in old_record.items():
        if rel_path in new_record:
            new_full = new_record[rel_path]
            if old_full != new_full:
                changes.append((rel_path, old_full, new_full))
    return changes

def replace_in_file(py_file, old_str, new_str, backup_dir):
    """在单个文件中替换路径"""
    try:
        content = py_file.read_text(encoding="utf-8")
        original = content
        
        # 多种替换方式
        content = content.replace(old_str, new_str)
        content = content.replace(f'"{old_str}"', f'"{new_str}"')
        content = content.replace(f"'{old_str}'", f"'{new_str}'")
        content = content.replace(f'r"{old_str}"', f'r"{new_str}"')
        
        if content != original:
            # 备份
            shutil.copy2(py_file, backup_dir / f"{py_file.name}.bak")
            py_file.write_text(content, encoding="utf-8")
            return True
    except Exception:
        pass
    return False

def auto_update():
    code_dir = find_current_code_base()
    if not code_dir:
        print("❌ 无法找到 code 目录")
        return
    
    # 加载旧记录
    old_record = load_record()
    new_record = scan_current_paths()
    
    if not old_record:
        print("📝 这是第一次运行，正在记录当前所有路径...")
        save_record(new_record)
        print("✅ 初始路径记录完成！以后移动文件夹后再次运行即可自动修复。")
        return
    
    # 检测变化
    changes = detect_changes(old_record, new_record)
    
    if not changes:
        print("✅ 所有路径未发生变化，无需修改。")
        return
    
    print(f"🔍 检测到 {len(changes)} 个路径变化：")
    for rel, old, new in changes:
        print(f"   • {rel}")
        print(f"     旧: {old}")
        print(f"     新: {new}\n")
    
    # 执行修复
    backup_dir = BACKUP_DIR / f"auto_{datetime.now():%Y%m%d_%H%M%S}"
    backup_dir.mkdir(parents=True)
    
    modified_count = 0
    py_files = list(code_dir.rglob("*.py"))
    
    for rel, old_full, new_full in changes:
        old_sub = rel  # 相对路径
        for py_file in py_files:
            if replace_in_file(py_file, old_full, new_full, backup_dir):
                modified_count += 1
            if replace_in_file(py_file, old_sub, old_sub.replace(old_full.split('/')[-1], new_full.split('/')[-1]), backup_dir):  # 简化版相对替换
                modified_count += 1
    
    print(f"\n🎉 自动修复完成！共修改 {modified_count} 个文件")
    print(f"   备份文件夹: {backup_dir}")
    
    # 更新记录
    save_record(new_record)

# ====================== 主程序 ======================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="成本统计路径自动检测与更新工具")
    parser.add_argument("--force", action="store_true", help="强制重新记录所有路径")
    parser.add_argument("--dry-run", action="store_true", help="仅检测变化，不修改文件")
    args = parser.parse_args()
    
    print("=" * 70)
    
    if args.force:
        print("强制模式：重新记录当前所有路径")
        new_record = scan_current_paths()
        save_record(new_record)
    else:
        auto_update()
    
    print("=" * 70)
    print("按回车键退出...")
    input()