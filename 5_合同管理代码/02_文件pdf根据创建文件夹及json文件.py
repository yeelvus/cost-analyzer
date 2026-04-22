import os
import shutil
import json
from collections import OrderedDict

# ==================== 配置区 ====================
DEFAULT_CONTRACT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑"
# ==============================================

def organize_contract_files_v4_2(root_path):
    valid_extensions = ('.pdf', '.doc', '.docx', '.xls', '.xlsx')
    new_tasks = []        # 需要新建文件夹 + 移动文件
    update_tasks = []     # 需要更新JSON（重点加强）

    def detect_contract_type_by_path(abs_dir_path, base_root):
        """增强版合同类型检测"""
        try:
            rel = os.path.relpath(abs_dir_path, base_root)
            parts = rel.split(os.sep) if rel != "." else [os.path.basename(base_root)]
            for p in reversed(parts):
                p_lower = p.lower()
                if "总包合同" in p or ("总包" in p_lower and "合同" in p_lower):
                    return "总包"
                if "分包合同" in p or ("分包" in p_lower and "合同" in p_lower):
                    return "分包"
            return None
        except:
            return None

    def get_new_json_template(contract_type):
        return OrderedDict([
            ("项目名称", ""),
            ("合同编号", ""),
            ("合同类型", contract_type),
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
        ])

    print("🔍 正在扫描目录，请稍等...\n")

    for root, dirs, files in os.walk(root_path):
        contract_type = detect_contract_type_by_path(root, root_path)
        if not contract_type:
            continue

        current_dir_name = os.path.basename(root)
        json_files = [f for f in files if f == "合同信息.json"]
        contract_files = [f for f in files if f.lower().endswith(valid_extensions) 
                         and not f.startswith('.') and not f.startswith('~$')]

        # ==================== 核心修复 ====================
        target_dirs_to_check = []

        # 情况1：目录下有合同文件（需要归档到同名文件夹）
        for file_name in contract_files:
            folder_name = os.path.splitext(file_name)[0]
            target_dir = os.path.join(root, folder_name)
            target_dirs_to_check.append((target_dir, file_name, "new"))

        # 情况2：目录下已经有“合同信息.json”的文件夹（重点！用于仅更新JSON）
        for d in dirs:
            potential_target = os.path.join(root, d)
            if os.path.exists(os.path.join(potential_target, "合同信息.json")):
                # 尝试找到对应的合同文件（可能已经在文件夹里了）
                inner_files = os.listdir(potential_target)
                contract_file_in_folder = next((f for f in inner_files if f.lower().endswith(valid_extensions)), None)
                file_name = contract_file_in_folder if contract_file_in_folder else "(已归档文件)"
                target_dirs_to_check.append((potential_target, file_name, "update"))

        # 处理所有待检查的目标文件夹
        for target_dir, file_name, action in target_dirs_to_check:
            json_path = os.path.join(target_dir, "合同信息.json")

            if action == "new" and not os.path.exists(target_dir):
                new_tasks.append({
                    "source_file": os.path.join(root, file_name),
                    "file_name": file_name,
                    "target_dir": target_dir,
                    "type": contract_type,
                    "parent_dir": current_dir_name,
                    "action": "new"
                })
            else:
                # 所有已有合同信息.json 的文件夹都加入更新任务
                update_tasks.append({
                    "target_dir": target_dir,
                    "file_name": file_name,
                    "type": contract_type,
                    "parent_dir": current_dir_name,
                    "action": "update_json"
                })

    # 去重 update_tasks
    seen = set()
    unique_update = []
    for t in update_tasks:
        key = t["target_dir"]
        if key not in seen:
            seen.add(key)
            unique_update.append(t)

    all_tasks = new_tasks + unique_update

    if not all_tasks:
        print("❌ 未发现任何可处理的项目。")
        print("提示：请确认根目录下是否存在包含“总包合同”或“分包合同”的文件夹。")
        return

    print(f"\n{'='*25} 扫描结果 {'='*25}")
    print(f"📊 需要新建文件夹并归档的文件: {len(new_tasks)} 个")
    print(f"📊 需要更新合同信息.json 的文件夹: {len(unique_update)} 个")
    print("=" * 60)

    print("\n请选择操作类型：")
    print("1. 仅处理新建文件夹（移动文件 + 创建新JSON）")
    print("2. 仅更新所有已存在的合同信息.json（推荐先执行）")
    print("3. 处理全部（新建 + 更新）")
    print("0. 取消")
    choice = input("\n请输入选项 (0-3): ").strip()

    if choice == '0':
        print("已取消操作。")
        return
    elif choice == '1':
        filtered_tasks = new_tasks
        task_desc = "新建归档"
    elif choice == '2':
        filtered_tasks = unique_update
        task_desc = "更新现有JSON"
    elif choice == '3':
        filtered_tasks = all_tasks
        task_desc = "全部任务"
    else:
        print("无效选项，已取消。")
        return

    if not filtered_tasks:
        print(f"\n[状态] 没有符合条件的{task_desc}。")
        return

    # 显示清单
    print(f"\n{'='*20} 即将执行的任务 ({task_desc}) {'='*20}")
    for i, task in enumerate(filtered_tasks, 1):
        act = "【新建并归档】" if task.get("action") == "new" else "【仅更新JSON】"
        print(f"{i}. {act} 【{task['type']}】 目录: {task['parent_dir']}")
        print(f"   文件/文件夹: {task.get('file_name', '合同信息.json')}")
        print("-" * 60)

    confirm = input(f"\n共计 {len(filtered_tasks)} 个任务，确认执行？(y/n): ").lower().strip()
    if confirm != 'y':
        print("已取消操作。")
        return

    # ====================== 执行 ======================
    success = 0
    fail = 0
    for task in filtered_tasks:
        try:
            target_dir = task['target_dir']
            os.makedirs(target_dir, exist_ok=True)
            json_path = os.path.join(target_dir, "合同信息.json")

            if task.get("action") == "new":
                # 移动文件
                source = task['source_file']
                final_file = os.path.join(target_dir, task['file_name'])
                if os.path.exists(source) and source != final_file:
                    shutil.move(source, final_file)
                # 创建新JSON
                template = get_new_json_template(task['type'])
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(template, f, ensure_ascii=False, indent=4)
            else:
                # 更新JSON（核心）
                if not os.path.exists(json_path):
                    template = get_new_json_template(task['type'])
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(template, f, ensure_ascii=False, indent=4)
                else:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        data = json.load(f, object_pairs_hook=OrderedDict)

                    # 删除不需要的字段
                    to_remove = ["合同状态", "关联主合同编号", "联系人及电话",
                                 "PDF文件路径", "收付款记录", "含税金额",
                                 "不含税金额", "生效日期", "税率"]
                    for field in to_remove:
                        data.pop(field, None)

                    # 字段改名
                    if "结束日期" in data:
                        data["预计竣工日期"] = data.pop("结束日期")
                    if "税率" in data:
                        data["汇率"] = data.pop("税率", "")

                    # 确保字段存在
                    if "预计竣工日期" not in data:
                        data["预计竣工日期"] = ""
                    if "汇率" not in data:
                        data["汇率"] = ""

                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=4)

            print(f" ✅ 成功: {task.get('file_name', '合同信息.json')}")
            success += 1

        except Exception as e:
            print(f" ❌ 失败: {task.get('file_name', 'JSON')} → {e}")
            fail += 1

    print(f"\n{'='*70}")
    print(f"🎉 处理完成！成功: {success} 个 | 失败: {fail} 个")
    print(f"{'='*70}")


if __name__ == "__main__":
    print("=" * 70)
    print("📋 合同归档工具 v4.2 （JSON更新加强版）")
    print("=" * 70)

    default_path = DEFAULT_CONTRACT_ROOT
    print(f"当前默认路径: {default_path}")

    use_default = input("\n是否使用默认路径？(y/n，直接回车=是): ").strip().lower()
    path = default_path if use_default not in ['n', 'no', '否'] else (input("请输入路径: ").strip() or default_path)

    if os.path.exists(path):
        organize_contract_files_v4_2(path)
    else:
        print(f"❌ 路径不存在: {path}")