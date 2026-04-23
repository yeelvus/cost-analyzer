# -*- coding: utf-8 -*-
"""
合同文件夹自动整理神器 v3.7（递归版 + JSON 标准化）

递归搜索任意层级合同文件夹（支持你的多层结构）。
每个匹配文件夹：补子夹JSON；文件建夹移动 + 补JSON。
日志/统计/弹窗确认。
零删除风险。
配置前置：改正则/JSON模板/深度。

ContractVault Organizer v3.7
功能：
1. 递归搜索任意层级 F_xxx_分包合同 / E_xxx_总包合同 文件夹
2. 每个PDF → 创建同名子文件夹 → 移动
3. 所有子文件夹自动补齐标准化合同信息.json（不覆盖已有内容）
4. 生成操作日志（TXT）+ 统计总结
5. Mac兼容对话框 + 最终确认 + 零删除风险

"""

import os
import shutil
import datetime
import json
from typing import Dict
from collections import OrderedDict

# ==================== 配置区（请在此修改） ====================
CONFIG = {
    "default_contract_root": "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑",
    # 仅处理该结构下“项目目录”的直层 PDF：
    # 5_合同管理📑/A_总包合同|B_分包合同/项目状态目录/项目目录/*.pdf
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
    "remove_fields": [
        "合同状态", "关联主合同编号", "联系人及电话", "PDF文件路径", "收付款记录",
        "含税金额", "不含税金额", "生效日期"
    ],
    "old_txt_files": [  # 需要删除的旧 TXT 文件
        "对方单位公司名称.txt", "合同编号.txt", "合同金额.txt",
        "合同签订日期.txt", "计划竣工时间.txt", "开工日期.txt",
        "备注.txt", "是否有业主结算资料.txt", "结算金额.txt"
    ],
    "max_recursion_depth": 100,                           # 最大递归深度
    "hidden_items": {".*", "Thumbs.db", "desktop.ini", "__MACOSX"}
}
# ============================================================


# ==================== 工具函数 ====================
def is_hidden(name: str) -> bool:
    return any(name == h or (h.startswith('.') and name.startswith('.')) for h in CONFIG["hidden_items"])


def log_operation(log_file: str, message: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def normalize_contract_type(folder_type: str) -> str:
    if "分包" in folder_type:
        return "分包"
    if "总包" in folder_type:
        return "总包"
    return ""


def detect_contract_type_by_path(abs_dir_path: str, base_root: str) -> str:
    """按路径语义识别合同类型，不依赖目录层级。"""
    try:
        rel = os.path.relpath(abs_dir_path, base_root)
    except Exception:
        return ""

    parts = rel.split(os.sep) if rel != "." else [os.path.basename(base_root)]
    for part in reversed(parts):
        if "分包合同" in part:
            return "分包"
        if "总包合同" in part:
            return "总包"
    return ""


def detect_contract_type_by_group_name(group_name: str) -> str:
    if "分包合同" in group_name:
        return "分包"
    if "总包合同" in group_name:
        return "总包"
    return ""


def ensure_json_file(folder_path: str, log_file: str, folder_type: str) -> tuple[int, int, int]:
    """确保文件夹中存在合同信息.json；已有文件保持原样，不做任何字段修改。"""
    json_filename = "合同信息.json"
    json_path = os.path.join(folder_path, json_filename)
    deleted_txt_count = 0
    contract_type = normalize_contract_type(folder_type)

    def build_template() -> OrderedDict:
        template = OrderedDict(CONFIG["json_template"])
        template["合同类型"] = contract_type
        return template

    if os.path.exists(json_path):
        # 已有JSON完全保持原样
        print(f"     └─ 保留已有: {json_filename}")
        log_operation(log_file, f"检查-保留: {json_path}")
        return 0, 0, 0
    else:
        # 文件不存在，创建新文件
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(build_template(), f, ensure_ascii=False, indent=2)
            print(f"     └─ 创建: {json_filename}")
            log_operation(log_file, f"检查-创建: {json_path}")
            json_result = (1, 0)  # 1个新建，0个更新
        except Exception as e:
            print(f"     └─ 创建失败 {json_filename}: {e}")
            log_operation(log_file, f"检查-失败: {json_path} | {e}")
            return 0, 0, 0
    
    # 不主动删除任何旧TXT，避免附带改动
    return json_result[0], json_result[1], deleted_txt_count


def iter_project_contract_dirs(root_dir: str):
    """仅枚举 项目目录（合同组/项目状态/项目名），不递归到更深层。"""
    print(f"开始扫描项目目录（根目录: {root_dir}）...")
    for group_name in os.listdir(root_dir):
        if is_hidden(group_name):
            continue
        group_path = os.path.join(root_dir, group_name)
        if not os.path.isdir(group_path):
            continue
        if not any(k in group_name for k in CONFIG["contract_group_keywords"]):
            continue

        contract_type = detect_contract_type_by_group_name(group_name)
        if not contract_type:
            continue

        print(f"[组] {group_path}")
        for status_name in os.listdir(group_path):
            if is_hidden(status_name):
                continue
            status_path = os.path.join(group_path, status_name)
            if not os.path.isdir(status_path):
                continue
            print(f"   [状态] {status_path}")

            for project_name in os.listdir(status_path):
                if is_hidden(project_name):
                    continue
                project_path = os.path.join(status_path, project_name)
                if not os.path.isdir(project_path):
                    continue
                print(f"      [项目] {project_path}")
                yield project_path, contract_type


# ==================== 核心处理 ====================
def process_folder(contract_path: str, folder_type: str, log_path: str, stats: Dict[str, int]):
    print(f"\n{folder_type} 文件夹: {contract_path}")
    log_operation(log_path, f"处理{folder_type}文件夹: {contract_path}")
    stats["processed_folders"] += 1

    filenames = os.listdir(contract_path)
    hidden_files = [f for f in filenames if is_hidden(f)]
    valid_extensions = ('.pdf',)
    valid_files = [
        f for f in filenames
        if os.path.isfile(os.path.join(contract_path, f))
        and not is_hidden(f)
        and f.lower().endswith(valid_extensions)
        and not f.startswith('~$')
        and f != "合同信息.json"
    ]
    subfolders = [d for d in filenames if os.path.isdir(os.path.join(contract_path, d)) and not is_hidden(d)]

    if hidden_files:
        stats["skipped_hidden"] += len(hidden_files)
        msg = f"  └─ 跳过隐藏文件: {', '.join(hidden_files)}"
        print(msg)
        log_operation(log_path, msg)

    if subfolders:
        msg = f"  ├─ 已有子文件夹 {len(subfolders)} 个（按需求不处理）"
        print(msg)
        log_operation(log_path, msg)

    if valid_files:
        print(f"  ├─ 处理 {len(valid_files)} 个直接文件")
        for filename in valid_files:
            print(f"  │ └─ 移动: {filename}")
            log_operation(log_path, f"处理文件: {filename}")

            src = os.path.join(contract_path, filename)
            new_folder_name = os.path.splitext(filename)[0]
            if is_hidden(new_folder_name):
                new_folder_name += "_file"
            new_folder_path = os.path.join(contract_path, new_folder_name)

            if os.path.exists(new_folder_path):
                msg = f"跳过（已有同名归档文件夹）: {new_folder_name}"
                print(f"  │   └─ {msg}")
                log_operation(log_path, msg)
                stats["skipped_existing_folder"] += 1
                continue

            os.makedirs(new_folder_path, exist_ok=True)
            stats["created_folders"] += 1

            dst = os.path.join(new_folder_path, filename)
            try:
                shutil.move(src, dst)
                stats["moved_files"] += 1
                msg = f"移动: {src} → {dst}"
                print(f"  │   └─ {msg}")
                log_operation(log_path, msg)
            except Exception as e:
                msg = f"移动失败 {filename}: {e}"
                print(f"  │   └─ {msg}")
                log_operation(log_path, msg)
                continue

            stats["checked_folders"] += 1
            print(f"  │   └─ 检查新文件夹: {new_folder_name}")
            c, s, d = ensure_json_file(new_folder_path, log_path, folder_type)
            stats["json_created"] += c
            stats["json_updated"] += s
            stats["deleted_txt"] += d
    else:
        print("  └─ 无直接文件")


# ==================== 主程序 ====================
if __name__ == "__main__":
    log_path = ""
    try:
        # 1. 固定根目录（无弹窗）
        default_root = os.path.normpath(CONFIG["default_contract_root"])
        if not os.path.isdir(default_root):
            raise FileNotFoundError(f"默认路径不存在: {default_root}")
        root_dir = default_root
        print(f"根目录: {root_dir}")

        # 2. 查找项目目录（固定到 合同组/项目状态/项目名）
        contract_folders = list(iter_project_contract_dirs(root_dir))
        if not contract_folders:
            print("未发现可处理的项目目录（请检查 合同组/项目状态/项目名 结构）。")
            raise SystemExit(0)

        # 3. 日志路径
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(root_dir, f"操作日志_{timestamp}.txt")
        log_operation(log_path, "=== 程序开始执行 ===")
        log_operation(log_path, f"根目录: {root_dir}")
        log_operation(log_path, f"候选目录数: {len(contract_folders)}")

        # 4. 初始化统计
        stats = {
            "processed_folders": 0,
            "moved_files": 0,
            "created_folders": 0,
            "checked_folders": 0,
            "json_created": 0,
            "json_updated": 0,
            "deleted_txt": 0,
            "skipped_hidden": 0,
            "skipped_existing_folder": 0,
        }

        # 5. 处理
        print("\n开始处理...\n")
        log_operation(log_path, "开始处理...")
        for path, folder_type in contract_folders:
            process_folder(path, folder_type, log_path, stats)

        # 6. 总结
        summary = f"""
处理完成！（根目录: {root_dir}）

总计处理文件夹: {stats["processed_folders"]} 个
移动文件数: {stats["moved_files"]} 个
创建子文件夹: {stats["created_folders"]} 个
检查子文件夹: {stats["checked_folders"]} 个
创建 JSON 文件: {stats["json_created"]} 个
更新 JSON 文件: {stats["json_updated"]} 个
删除旧 TXT 文件: {stats["deleted_txt"]} 个
跳过隐藏文件: {stats["skipped_hidden"]} 个
跳过已有归档文件夹: {stats["skipped_existing_folder"]} 个

日志文件: {os.path.basename(log_path)}
"""
        print(summary)
        log_operation(log_path, "=== 处理完成 ===")
        log_operation(log_path, summary.strip())
        print(f"日志保存至: {log_path}")

    except Exception as e:
        error_msg = f"程序异常：{str(e)}"
        print(error_msg)
        if log_path:
            log_operation(log_path, f"异常终止: {error_msg}")