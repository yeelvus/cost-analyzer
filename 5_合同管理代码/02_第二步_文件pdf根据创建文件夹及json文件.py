import os
import shutil
import json
from collections import OrderedDict

# ==================== 配置区 ====================
# 默认合同根目录路径（可修改为你的实际路径）
DEFAULT_CONTRACT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理"  # 相对路径
# 或使用绝对路径：
# DEFAULT_CONTRACT_ROOT = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理"
# ==============================================

def organize_contract_files_v3(root_path):
    # 严格匹配逻:：
    # 1. 文件夹以 'E_' 开头且包含 '总包合同' -> 总包
    # 2. 文件夹以 'F_' 开头且包含 '分包合同' -> 分包
    valid_extensions = ('.pdf', '.doc', '.docx', '.xls', '.xlsx')
    
    pending_tasks = []

    # 1. 扫描符合特定命名规则的文件夹
    for root, dirs, files in os.walk(root_path):
        current_dir_name = os.path.basename(root)
        
        contract_type = None
        if current_dir_name.startswith('E_') and "总包合同" in current_dir_name:
            contract_type = "总包"
        elif current_dir_name.startswith('F_') and "分包合同" in current_dir_name:
            contract_type = "分包"
        
        # 如果命中规则
        if contract_type:
            for file_name in files:
                # 过滤隐藏文件、JSON、临时文件
                if file_name.startswith('.') or file_name.endswith('.json') or file_name.startswith('~$'):
                    continue
                
                if file_name.lower().endswith(valid_extensions):
                    file_full_path = os.path.join(root, file_name)
                    # 准备创建的子文件夹名（即文件名）
                    folder_name = os.path.splitext(file_name)[0]
                    target_folder_path = os.path.join(root, folder_name)
                    
                    # 只有当同名文件夹不存在时，才视为待处理文件
                    if not os.path.exists(target_folder_path):
                        pending_tasks.append({
                            "source_file": file_full_path,
                            "file_name": file_name,
                            "target_dir": target_folder_path,
                            "type": contract_type,
                            "parent_dir": current_dir_name
                        })

    # 2. 检查是否有待处理任务
    if not pending_tasks:
        print("\n[状态] 未发现符合条件的待处理文件。")
        print("当前匹配规则：")
        print(" - 文件夹需以 'E_' 开头并包含 '总包合同'")
        print(" - 或文件夹需以 'F_' 开头并包含 '分包合同'")
        return

    # 统计各类型数量
    total_count = len([t for t in pending_tasks if t['type'] == '总包'])
    subcontract_count = len([t for t in pending_tasks if t['type'] == '分包'])
    
    # 3. 让用户选择处理类型
    print(f"\n{'='*20} 合同归档统计 {'='*20}")
    print(f"📊 总包合同: {total_count} 个")
    print(f"📊 分包合同: {subcontract_count} 个")
    print(f"📊 合计: {len(pending_tasks)} 个")
    print("=" * 60)
    
    print("\n请选择要处理的合同类型：")
    print("1. 仅处理总包合同")
    print("2. 仅处理分包合同")
    print("3. 处理全部合同")
    print("0. 取消操作")
    
    choice = input("\n请输入选项 (0-3): ").strip()
    
    # 根据选择过滤任务
    if choice == '0':
        print("已取消操作。")
        return
    elif choice == '1':
        filtered_tasks = [t for t in pending_tasks if t['type'] == '总包']
        task_desc = "总包合同"
    elif choice == '2':
        filtered_tasks = [t for t in pending_tasks if t['type'] == '分包']
        task_desc = "分包合同"
    elif choice == '3':
        filtered_tasks = pending_tasks
        task_desc = "全部合同"
    else:
        print("无效选项，已取消操作。")
        return
    
    if not filtered_tasks:
        print(f"\n[状态] 没有符合条件的{task_desc}需要处理。")
        return
    
    # 4. 显示过滤后的清单
    print(f"\n{'='*20} 待归档{task_desc}清单 {'='*20}")
    for i, task in enumerate(filtered_tasks, 1):
        print(f"{i}. 【{task['type']}】所在目录: {task['parent_dir']}")
        print(f"   待处理文件: {task['file_name']}")
        print(f"   -> 将移动至: {os.path.basename(task['target_dir'])}/")
        print("-" * 50)
    
    # 5. 最终确认执行
    confirm = input(f"\n共计 {len(filtered_tasks)} 个任务，确认批量执行？(y/n): ").lower().strip()
    if confirm != 'y':
        print("已取消操作。")
        return
    
    # 使用过滤后的任务列表
    pending_tasks = filtered_tasks

    # 4. 执行逻辑
    success_count = 0
    fail_count = 0
    
    for task in pending_tasks:
        try:
            # 创建文件夹
            os.makedirs(task['target_dir'], exist_ok=True)
            
            # 移动文件
            final_file_path = os.path.join(task['target_dir'], task['file_name'])
            shutil.move(task['source_file'], final_file_path)
            
            # 生成 JSON
            json_path = os.path.join(task['target_dir'], "合同信息.json")
            template = OrderedDict([
                ("项目名称", ""),
                ("合同编号", ""),
                ("合同类型", task['type']),
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
                ("PDF文件路径", final_file_path),
                ("收付款记录", []),
                ("备注", "")
            ])
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, ensure_ascii=False, indent=4)
            
            print(f" ✅ [OK] 成功归档: {task['file_name']}")
            success_count += 1
        except Exception as e:
            print(f" ❌ [Err] 处理 {task['file_name']} 时失败: {e}")
            fail_count += 1

    print(f"\n{'='*60}")
    print(f"✅ 任务全部处理完毕！")
    print(f"   成功: {success_count} 个")
    if fail_count > 0:
        print(f"   失败: {fail_count} 个")
    print(f"{'='*60}")

if __name__ == "__main__":
    print("=" * 60)
    print("📋 合同归档工具 (精准匹配版)")
    print("=" * 60)
    
    # 使用默认路径
    default_path = DEFAULT_CONTRACT_ROOT
    print(f"\n📁 当前默认路径: {default_path}")
    print(f"   绝对路径: {os.path.abspath(default_path)}")
    
    # 询问是否使用默认路径
    use_default = input("\n是否使用默认路径？(y/n，直接回车使用默认): ").strip().lower()
    
    if use_default == 'n' or use_default == 'no' or use_default == '否':
        path = input("请输入合同根目录路径: ").strip()
        if not path:
            print("⚠️ 未输入路径，使用默认路径。")
            path = default_path
    else:
        path = default_path
    
    # 验证路径
    if os.path.exists(path):
        print(f"\n✅ 开始扫描目录: {os.path.abspath(path)}\n")
        organize_contract_files_v3(path)
    else:
        print(f"\n❌ 路径无效: {path}")
        print("💡 提示: 请检查路径是否正确，或修改代码中的 DEFAULT_CONTRACT_ROOT 配置。")