# -*- coding: utf-8 -*-
"""
合同文件夹自动整理工具 v4.7（JSON字段顺序固定版）
"""
import os
import shutil
import datetime
import json
from collections import OrderedDict
from typing import Dict, Tuple

# ==================== 配置区 ====================
CONFIG = {
    "default_contract_root": "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑",
    "contract_group_keywords": ["总包合同", "分包合同"],
    "hidden_items": {".*", "Thumbs.db", "desktop.ini", "__MACOSX"},
    "contract_extensions": ('.pdf', '.doc', '.docx', '.xls', '.xlsx')
}

# ==================== JSON字段顺序定义 ====================
JSON_FIELD_ORDER = [
    "项目名称",
    "合同名称",
    "合同编号",
    "合同类型",
    "对方公司名称",
    "合同签订日期",
    "预计竣工日期",
    "工期/交货期",
    "币种",
    "汇率",
    "合同金额",
    "付款方式",
    "履约保证金",
    "备注"
]

# ==================== 解析函数 ====================
def parse_contract_and_project_name(folder_path: str) -> Tuple[str, str, str]:
    """从合同文件夹路径解析信息"""
    contract_folder = os.path.basename(folder_path)
    name_no_ext = os.path.splitext(contract_folder)[0].strip()

    print(f"   📄 处理文件夹: {contract_folder}")

    if '_' in name_no_ext:
        parts = name_no_ext.split('_', 1)
        contract_code = parts[0].strip()
        contract_name = parts[1].strip()
    else:
        contract_code = ""
        contract_name = name_no_ext

    project_name = ""
    try:
        project_folder = os.path.basename(os.path.dirname(folder_path))
        if '_' in project_folder:
            project_name = project_folder.split('_', 1)[-1].strip()
        else:
            project_name = project_folder
    except:
        pass

    print(f"   🔍 解析结果 → 项目名称:【{project_name}】 | 合同名称:【{contract_name}】 | 合同编号:【{contract_code}】")
    return project_name, contract_name, contract_code


# ==================== 工具函数 ====================
def is_hidden(name: str) -> bool:
    return any(name == h or (h.startswith('.') and name.startswith('.')) 
               for h in CONFIG["hidden_items"])


def log_operation(log_file: str, message: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def normalize_contract_type(folder_name: str) -> str:
    if "分包" in folder_name:
        return "分包"
    if "总包" in folder_name:
        return "总包"
    return ""


def get_unique_path(path: str) -> str:
    """若路径已存在，生成带序号的新路径，避免覆盖"""
    if not os.path.exists(path):
        return path

    base, ext = os.path.splitext(path)
    idx = 1
    while True:
        candidate = f"{base}_{idx}{ext}"
        if not os.path.exists(candidate):
            return candidate
        idx += 1


# ==================== JSON 更新核心 ====================
def ensure_json_file(folder_path: str, log_file: str, contract_type: str):
    """强制更新合同信息.json，并严格控制字段顺序"""
    json_path = os.path.join(folder_path, "合同信息.json")
    project_name, contract_name, contract_code = parse_contract_and_project_name(folder_path)

    # 读取已有数据
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            print(f"   └─ 读取已有JSON")
        except Exception:
            data = {}
    else:
        data = {}
        print(f"   └─ 未找到JSON，将创建")

    # 强制更新核心字段
    updated = False

    if project_name and data.get("项目名称") != project_name:
        print(f"   → 更新 项目名称 → {project_name}")
        data["项目名称"] = project_name
        updated = True

    if contract_name and data.get("合同名称") != contract_name:
        print(f"   → 更新 合同名称 → {contract_name}")
        data["合同名称"] = contract_name
        updated = True

    if contract_code and not data.get("合同编号"):
        print(f"   → 更新 合同编号 → {contract_code}")
        data["合同编号"] = contract_code
        updated = True

    if contract_type:
        data["合同类型"] = contract_type

    # 构建有序字典，保证输出顺序
    ordered_data = OrderedDict()
    for key in JSON_FIELD_ORDER:
        ordered_data[key] = data.get(key, "")

    # 把其他未知字段放到备注之后
    for key, value in data.items():
        if key not in JSON_FIELD_ORDER:
            ordered_data[key] = value

    # 写入文件
    try:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(ordered_data, f, ensure_ascii=False, indent=2)
        print(f"   ✅ JSON更新成功！（字段顺序已固定）")
        log_operation(log_file, f"JSON更新成功 → {contract_name}")
    except Exception as e:
        print(f"   ❌ JSON写入失败: {e}")


# ==================== 核心处理函数 ====================
def process_project_folder(contract_path: str, contract_type: str, log_path: str, stats: Dict):
    print(f"\n[{contract_type}] 处理项目: {contract_path}")
    log_operation(log_path, f"开始处理: {contract_path}")
    stats["processed_folders"] += 1

    # 先处理直接散落在项目目录中的合同文件：自动建合同文件夹并移动文件
    for item in os.listdir(contract_path):
        if is_hidden(item):
            continue

        item_path = os.path.join(contract_path, item)
        if not os.path.isfile(item_path):
            continue

        if not item.lower().endswith(CONFIG["contract_extensions"]):
            continue

        contract_folder_name = os.path.splitext(item)[0].strip() or "未命名合同"
        target_folder = os.path.join(contract_path, contract_folder_name)

        if os.path.exists(target_folder) and not os.path.isdir(target_folder):
            target_folder = get_unique_path(os.path.join(contract_path, f"{contract_folder_name}_合同"))

        if not os.path.exists(target_folder):
            os.makedirs(target_folder, exist_ok=True)
            stats["created_folders"] += 1
            print(f"   ├─ 新建合同文件夹: {os.path.basename(target_folder)}")
            log_operation(log_path, f"新建合同文件夹: {target_folder}")

        target_file = get_unique_path(os.path.join(target_folder, item))
        try:
            shutil.move(item_path, target_file)
            stats["moved_files"] += 1
            print(f"   ├─ 移动合同文件: {item} -> {os.path.basename(target_folder)}")
            log_operation(log_path, f"移动文件: {item_path} -> {target_file}")
        except Exception as e:
            print(f"   ❌ 移动文件失败: {item}，原因: {e}")
            log_operation(log_path, f"移动失败: {item_path}，原因: {e}")

    # 再处理合同文件夹，补齐/更新JSON
    for item in os.listdir(contract_path):
        item_path = os.path.join(contract_path, item)
        if not os.path.isdir(item_path) or is_hidden(item):
            continue

        files = os.listdir(item_path)
        has_contract_file = any(f.lower().endswith(CONFIG["contract_extensions"]) 
                               for f in files if not is_hidden(f))
        has_json = "合同信息.json" in files

        if has_contract_file or has_json:
            print(f"   ├─ 处理合同文件夹: {item}")
            ensure_json_file(item_path, log_path, contract_type)


# ==================== 主程序 ====================
if __name__ == "__main__":
    try:
        root_dir = os.path.normpath(CONFIG["default_contract_root"])
        if not os.path.isdir(root_dir):
            raise FileNotFoundError(f"根目录不存在: {root_dir}")

        print("合同整理工具 v4.7（JSON字段顺序固定版）启动")
        print(f"根目录: {root_dir}\n")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(root_dir, f"合同整理日志_{timestamp}.txt")

        stats = {"processed_folders": 0, "moved_files": 0, "created_folders": 0, "skipped_existing_folder": 0}

        for group_name in os.listdir(root_dir):
            if is_hidden(group_name): continue
            group_path = os.path.join(root_dir, group_name)
            if not os.path.isdir(group_path): continue
            if not any(k in group_name for k in CONFIG["contract_group_keywords"]): continue

            contract_type = normalize_contract_type(group_name)
            if not contract_type: continue

            print(f"[合同组] {group_name}")

            for status_name in os.listdir(group_path):
                if is_hidden(status_name): continue
                status_path = os.path.join(group_path, status_name)
                if not os.path.isdir(status_path): continue

                for project_name in os.listdir(status_path):
                    if is_hidden(project_name): continue
                    project_path = os.path.join(status_path, project_name)
                    if os.path.isdir(project_path):
                        process_project_folder(project_path, contract_type, log_path, stats)

        print("\n=== 处理完成 ===")
        print(f"共处理项目文件夹: {stats['processed_folders']} 个")
        print(f"新建合同文件夹: {stats['created_folders']} 个")
        print(f"移动合同文件: {stats['moved_files']} 个")

    except Exception as e:
        print(f"程序异常: {e}")