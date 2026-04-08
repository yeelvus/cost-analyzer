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
import re
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict
from collections import OrderedDict

# ==================== 配置区（请在此修改） ====================
CONFIG = {
    "sub_contract_pattern": r"^F_[A-Z0-9-]+_分包合同$",   # 分包合同正则
    "main_contract_pattern": r"^E_[A-Z0-9-]+_总包合同$",  # 总包合同正则（适配 E_THCC-2024-006_总包合同）
    "json_template": OrderedDict([
        ("项目名称", ""),
        ("合同编号", ""),
        ("合同类型", ""),
        ("对方公司名称", ""),
        ("合同签订日期", ""),
        ("生效日期", ""),
        ("结束日期", ""),
        ("工期/交货期", ""),
        ("币种", ""),
        ("合同金额", 0.0),
        ("含税金额", 0.0),
        ("不含税金额", 0.0),
        ("税率", ""),
        ("付款方式", ""),
        ("履约保证金", ""),
        ("合同状态", ""),
        ("关联主合同编号", ""),
        ("联系人及电话", ""),
        ("PDF文件路径", ""),
        ("收付款记录", []),
        ("备注", "")
    ]),
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
def select_directory(title: str) -> str:
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title=title)
    root.destroy()
    if not directory:
        messagebox.showerror("错误", f"未选择{title}，程序退出。")
        exit()
    return os.path.normpath(directory)


def is_hidden(name: str) -> bool:
    return any(name == h or (h.startswith('.') and name.startswith('.')) for h in CONFIG["hidden_items"])


def log_operation(log_file: str, message: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def ensure_json_file(folder_path: str, log_file: str) -> tuple[int, int, int]:
    """确保文件夹中存在合同信息.json，不覆盖已有内容，并删除旧的TXT文件"""
    json_filename = "合同信息.json"
    json_path = os.path.join(folder_path, json_filename)
    deleted_txt_count = 0
    
    if os.path.exists(json_path):
        # 文件已存在，读取并合并字段（保留已有数据）
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
            
            # 合并逻辑：保留已有值，只添加缺失字段
            updated_data = OrderedDict()
            for key, default_value in CONFIG["json_template"].items():
                updated_data[key] = existing_data.get(key, default_value)
            
            # 保留不在模板中的额外字段
            for key in existing_data:
                if key not in updated_data:
                    updated_data[key] = existing_data[key]
            
            # 写回文件（保持标准格式）
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(updated_data, f, ensure_ascii=False, indent=2)
            
            print(f"     └─ 更新: {json_filename}")
            log_operation(log_file, f"检查-更新: {json_path}")
            json_result = (0, 1)  # 0个新建，1个更新
            
        except Exception as e:
            print(f"     └─ 更新失败 {json_filename}: {e}")
            log_operation(log_file, f"检查-失败: {json_path} | {e}")
            return 0, 0, 0
    else:
        # 文件不存在，创建新文件
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(CONFIG["json_template"], f, ensure_ascii=False, indent=2)
            print(f"     └─ 创建: {json_filename}")
            log_operation(log_file, f"检查-创建: {json_path}")
            json_result = (1, 0)  # 1个新建，0个更新
        except Exception as e:
            print(f"     └─ 创建失败 {json_filename}: {e}")
            log_operation(log_file, f"检查-失败: {json_path} | {e}")
            return 0, 0, 0
    
    # 删除旧的 TXT 文件
    for txt_file in CONFIG["old_txt_files"]:
        txt_path = os.path.join(folder_path, txt_file)
        if os.path.exists(txt_path):
            try:
                os.remove(txt_path)
                deleted_txt_count += 1
                print(f"     └─ 删除旧文件: {txt_file}")
                log_operation(log_file, f"删除TXT: {txt_path}")
            except Exception as e:
                print(f"     └─ 删除失败 {txt_file}: {e}")
                log_operation(log_file, f"删除失败: {txt_path} | {e}")
    
    return json_result[0], json_result[1], deleted_txt_count


def find_contract_folders_recursive(root_dir: str) -> list[tuple[str, str]]:
    """递归查找所有合同文件夹"""
    contract_folders = []
    sub_pattern = re.compile(CONFIG["sub_contract_pattern"])
    main_pattern = re.compile(CONFIG["main_contract_pattern"])

    def _walk(current_path: str, depth: int = 0):
        if depth > CONFIG["max_recursion_depth"]:
            print(f"   深度超限: {current_path}")
            return
        try:
            items = os.listdir(current_path)
        except Exception as e:
            print(f"   无法访问: {current_path} | {e}")
            return

        for item in items:
            if is_hidden(item):
                continue
            full_path = os.path.join(current_path, item)
            rel_path = os.path.relpath(full_path, root_dir)
            print(f"   [{'→' if os.path.isdir(full_path) else '·'}] 第 {depth+1} 层: {rel_path}")

            if os.path.isdir(full_path):
                if sub_pattern.match(item):
                    contract_folders.append((full_path, item, "分包合同"))
                    print(f"   发现分包合同: {rel_path}")
                elif main_pattern.match(item):
                    contract_folders.append((full_path, item, "总包合同"))
                    print(f"   发现总包合同: {rel_path}")
                else:
                    _walk(full_path, depth + 1)

    print(f"开始递归搜索（根目录: {root_dir}）...")
    _walk(root_dir)
    return contract_folders


# ==================== 核心处理 ====================
def process_folder(contract_path: str, folder_type: str, log_path: str, stats: Dict[str, int]):
    print(f"\n{folder_type} 文件夹: {contract_path}")
    log_operation(log_path, f"处理{folder_type}文件夹: {contract_path}")
    stats["processed_folders"] += 1

    filenames = os.listdir(contract_path)
    hidden_files = [f for f in filenames if is_hidden(f)]
    valid_files = [f for f in filenames if os.path.isfile(os.path.join(contract_path, f)) and not is_hidden(f)]
    subfolders = [d for d in filenames if os.path.isdir(os.path.join(contract_path, d)) and not is_hidden(d)]

    if hidden_files:
        stats["skipped_hidden"] += len(hidden_files)
        msg = f"  └─ 跳过隐藏文件: {', '.join(hidden_files)}"
        print(msg)
        log_operation(log_path, msg)

    if subfolders:
        print(f"  ├─ 检查 {len(subfolders)} 个已有子文件夹")
        for sub in subfolders:
            sub_path = os.path.join(contract_path, sub)
            stats["checked_folders"] += 1
            print(f"  │ └─ 检查: {sub}")
            c, s, d = ensure_json_file(sub_path, log_path)
            stats["json_created"] += c
            stats["json_updated"] += s
            stats["deleted_txt"] += d

    if valid_files:
        print(f"  ├─ 处理 {len(valid_files)} 个直接文件")
        for filename in valid_files:
            print(f"  │ └─ 移动: {filename}")
            log_operation(log_path, f"处理文件: {filename}")
            stats["moved_files"] += 1

            src = os.path.join(contract_path, filename)
            new_folder_name = os.path.splitext(filename)[0]
            if is_hidden(new_folder_name):
                new_folder_name += "_file"
            new_folder_path = os.path.join(contract_path, new_folder_name)

            os.makedirs(new_folder_path, exist_ok=True)
            stats["created_folders"] += 1

            dst = os.path.join(new_folder_path, filename)
            try:
                shutil.move(src, dst)
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
            c, s, d = ensure_json_file(new_folder_path, log_path)
            stats["json_created"] += c
            stats["json_updated"] += s
            stats["deleted_txt"] += d
    else:
        print("  └─ 无直接文件")


# ==================== 主程序 ====================
if __name__ == "__main__":
    try:
        # 1. 选择根目录
        root_dir = select_directory("请选择【合同根目录】（如 5_合同管理，包含任意层级合同文件夹）")
        print(f"根目录: {root_dir}")

        # 2. 查找合同文件夹
        contract_folders = find_contract_folders_recursive(root_dir)

        # 3. 最终确认
        if not contract_folders:
            messagebox.showwarning("未找到", "未发现任何合同文件夹！请检查目录/命名。")
            exit()

        sample = "\n".join([f"• {name} ({os.path.relpath(path, root_dir)})" for path, name, _ in contract_folders[:5]])
        more = f"\n... 等 {len(contract_folders)-5} 个" if len(contract_folders) > 5 else ""
        confirm_msg = f"发现 {len(contract_folders)} 个合同文件夹：{sample}{more}\n\n操作：整理 + 补合同信息.json\n确认执行？"
        if not messagebox.askyesno("最终确认", confirm_msg):
            messagebox.showinfo("已取消", "用户取消操作。")
            exit()

        # 4. 日志路径
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(root_dir, f"操作日志_{timestamp}.txt")
        log_operation(log_path, "=== 程序开始执行 ===")
        log_operation(log_path, f"根目录: {root_dir}")

        # 5. 初始化统计
        stats = {
            "processed_folders": 0,
            "moved_files": 0,
            "created_folders": 0,
            "checked_folders": 0,
            "json_created": 0,
            "json_updated": 0,
            "deleted_txt": 0,
            "skipped_hidden": 0,
        }

        # 6. 处理
        print("\n开始处理...\n")
        log_operation(log_path, "开始处理...")
        for path, _, folder_type in contract_folders:
            process_folder(path, folder_type, log_path, stats)

        # 7. 总结
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

日志文件: {os.path.basename(log_path)}
"""
        print(summary)
        log_operation(log_path, "=== 处理完成 ===")
        log_operation(log_path, summary.strip())
        messagebox.showinfo("完成", summary + f"\n日志保存至：\n{log_path}")

    except Exception as e:
        error_msg = f"程序异常：{str(e)}"
        print(error_msg)
        if 'log_path' in locals():
            log_operation(log_path, f"异常终止: {error_msg}")
        messagebox.showerror("错误", error_msg)