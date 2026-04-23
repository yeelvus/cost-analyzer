# -*- coding: utf-8 -*-
"""
合同文件夹自动整理工具 v4.0（精简优化版）
功能：
1. 递归识别包含“总包合同”或“分包合同”的目录
2. 对目录下的 PDF 文件：创建同名子文件夹并移动文件
3. 自动为每个子文件夹创建或保留 “合同信息.json”
4. 零删除风险，不修改已有 JSON 内容
"""

import os
import shutil
import datetime
import json
from collections import OrderedDict
from typing import Dict, List, Tuple

# ==================== 配置区 ====================
CONFIG = {
    "default_contract_root": "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑",
    "contract_group_keywords": ["总包合同", "分包合同"],
    "json_template": OrderedDict([
        ("项目名称", ""),
        ("合同编号", ""),
        ("合同类型", ""),
        ("对方公司名称", ""),
        ("合同签订日期", ""),
        ("预计竣工日期", ""),
        ("工期/交货期", ""),
        ("币种", ""),
        ("合同金额", 0.0),
        ("汇率", ""),
        ("付款方式", ""),
        ("履约保证金", ""),
        ("备注", "")
    ]),
    "hidden_items": {".*", "Thumbs.db", "desktop.ini", "__MACOSX"}
}


# ==================== 工具函数 ====================
def is_hidden(name: str) -> bool:
    """判断是否为隐藏文件或系统文件"""
    return any(name == h or (h.startswith('.') and name.startswith('.')) 
               for h in CONFIG["hidden_items"])


def log_operation(log_file: str, message: str):
    """记录操作日志"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def normalize_contract_type(folder_name: str) -> str:
    """规范化合同类型"""
    if "分包" in folder_name:
        return "分包"
    if "总包" in folder_name:
        return "总包"
    return ""


def detect_contract_type_by_path(abs_dir_path: str, base_root: str) -> str:
    """通过路径识别合同类型（从深到浅查找）"""
    try:
        rel = os.path.relpath(abs_dir_path, base_root)
        parts = rel.split(os.sep) if rel != "." else [os.path.basename(base_root)]
        
        for part in reversed(parts):
            if "分包合同" in part:
                return "分包"
            if "总包合同" in part:
                return "总包"
    except Exception:
        pass
    return ""


def ensure_json_file(folder_path: str, log_file: str, contract_type: str) -> int:
    """确保文件夹中有合同信息.json，已有文件则完全保留"""
    json_path = os.path.join(folder_path, "合同信息.json")
    
    if os.path.exists(json_path):
        print(f" └─ 保留已有: 合同信息.json")
        log_operation(log_file, f"保留JSON: {json_path}")
        return 0  # 0 表示未创建新文件
    
    # 创建新JSON
    try:
        template = CONFIG["json_template"].copy()
        template["合同类型"] = contract_type
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        
        print(f" └─ 创建: 合同信息.json")
        log_operation(log_file, f"创建JSON: {json_path}")
        return 1  # 1 表示创建了新文件
        
    except Exception as e:
        print(f" └─ 创建JSON失败: {e}")
        log_operation(log_file, f"创建JSON失败: {json_path} | {e}")
        return 0


# ==================== 核心处理函数 ====================
def process_project_folder(contract_path: str, contract_type: str, log_path: str, stats: Dict[str, int]):
    """处理单个项目文件夹"""
    print(f"\n[{contract_type}] 处理文件夹: {contract_path}")
    log_operation(log_path, f"开始处理: {contract_path}")
    stats["processed_folders"] += 1

    filenames = os.listdir(contract_path)
    valid_extensions = ('.pdf',)
    
    # 筛选有效的PDF文件
    valid_files = [
        f for f in filenames
        if os.path.isfile(os.path.join(contract_path, f))
        and not is_hidden(f)
        and f.lower().endswith(valid_extensions)
        and not f.startswith('~$')
        and f != "合同信息.json"
    ]

    if not valid_files:
        print(" └─ 无PDF文件需要处理")
        log_operation(log_path, "无PDF文件")
        return

    print(f" ├─ 发现 {len(valid_files)} 个PDF文件")
    
    for filename in valid_files:
        src = os.path.join(contract_path, filename)
        folder_name = os.path.splitext(filename)[0]
        new_folder_path = os.path.join(contract_path, folder_name)

        # 防止文件夹名与隐藏规则冲突
        if is_hidden(folder_name):
            folder_name += "_file"
            new_folder_path = os.path.join(contract_path, folder_name)

        print(f" │ ├─ 处理: {filename}")

        # 如果目标文件夹已存在，跳过移动
        if os.path.exists(new_folder_path):
            print(f" │ └─ 跳过（已存在同名文件夹）")
            log_operation(log_path, f"跳过已存在文件夹: {folder_name}")
            stats["skipped_existing_folder"] += 1
            continue

        # 创建文件夹并移动文件
        try:
            os.makedirs(new_folder_path, exist_ok=True)
            dst = os.path.join(new_folder_path, filename)
            shutil.move(src, dst)
            
            stats["moved_files"] += 1
            stats["created_folders"] += 1
            
            print(f" │ └─ 已移动 → {folder_name}/")
            log_operation(log_path, f"移动成功: {filename} → {folder_name}/")
            
            # 为新文件夹创建JSON
            created = ensure_json_file(new_folder_path, log_path, contract_type)
            stats["json_created"] += created
            
        except Exception as e:
            print(f" │ └─ 处理失败: {e}")
            log_operation(log_path, f"处理失败: {filename} | {e}")


# ==================== 主程序 ====================
if __name__ == "__main__":
    log_path = ""
    try:
        root_dir = os.path.normpath(CONFIG["default_contract_root"])
        
        if not os.path.isdir(root_dir):
            raise FileNotFoundError(f"默认路径不存在: {root_dir}")
        
        print(f"合同整理工具 v4.0 开始运行")
        print(f"根目录: {root_dir}\n")
        
        # 生成日志文件
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(root_dir, f"合同整理日志_{timestamp}.txt")
        
        log_operation(log_path, "=== 合同整理程序启动 ===")
        log_operation(log_path, f"根目录: {root_dir}")
        
        # 初始化统计
        stats: Dict[str, int] = {
            "processed_folders": 0,
            "moved_files": 0,
            "created_folders": 0,
            "json_created": 0,
            "skipped_existing_folder": 0,
        }
        
        # 遍历所有项目文件夹（仅处理合同组 → 项目状态 → 项目 这一层结构）
        print("正在扫描项目目录...\n")
        
        for group_name in os.listdir(root_dir):
            if is_hidden(group_name):
                continue
            group_path = os.path.join(root_dir, group_name)
            if not os.path.isdir(group_path):
                continue
            if not any(k in group_name for k in CONFIG["contract_group_keywords"]):
                continue
                
            contract_type = normalize_contract_type(group_name)
            if not contract_type:
                continue
                
            print(f"[合同组] {group_name}")
            
            # 遍历项目状态目录
            for status_name in os.listdir(group_path):
                if is_hidden(status_name):
                    continue
                status_path = os.path.join(group_path, status_name)
                if not os.path.isdir(status_path):
                    continue
                    
                # 遍历具体项目目录
                for project_name in os.listdir(status_path):
                    if is_hidden(project_name):
                        continue
                    project_path = os.path.join(status_path, project_name)
                    if not os.path.isdir(project_path):
                        continue
                        
                    process_project_folder(project_path, contract_type, log_path, stats)
        
        # 处理完成总结
        summary = f"""
=== 处理完成 ===
根目录: {root_dir}
处理项目文件夹: {stats["processed_folders"]} 个
移动PDF文件: {stats["moved_files"]} 个
创建子文件夹: {stats["created_folders"]} 个
创建JSON文件: {stats["json_created"]} 个
跳过已有文件夹: {stats["skipped_existing_folder"]} 个

日志文件: {os.path.basename(log_path)}
"""
        print(summary)
        log_operation(log_path, "=== 处理完成 ===")
        log_operation(log_path, summary.strip())
        print(f"日志已保存至: {log_path}")
        
    except Exception as e:
        error_msg = f"程序异常: {str(e)}"
        print(error_msg)
        if log_path:
            log_operation(log_path, f"异常终止: {error_msg}")