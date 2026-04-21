import os
import shutil
import pandas as pd
import json
from datetime import datetime

# 工具说明 (README)

# 用途：根据 Excel 里的项目状态（备注列），自动将 5_合同管理 下的项目文件夹移动到对应的分类目录。

# 核心逻辑：

# 扫描器：自动识别以 THCC-XXXX-XXX 开头的二级文件夹。

# 映射器：读取 Excel 并将其转换为 migration_config.json。

# 执行器：根据 JSON 配置移动物理文件夹，并记录详细日志。

# 恢复机制：若发生误操作，可根据同目录下的 file_migration_log.txt 还原。

# Excel 文件 (项目编码/名称/状态)
#     ↓
# 读取并转换为 JSON 配置
#     ↓
# 扫描现有文件夹结构
#     ↓
# 对比：当前位置 vs 目标位置
#     ↓
# 执行移动操作 + 记录日志

# 用户运行脚本
#   ↓
# 读取 Excel：THCC-0001-200 | 项目A | "已完成"
#   ↓
# 发现该项目当前在 "进行中" 目录
#   ↓
# 确认移动：5_合同管理/进行中/THCC-0001-200 
#        → 5_合同管理/已完成/THCC-0001-200
#   ↓
# 日志记录：[时间] 移动项目: THCC-0001-200 | 项目A 从 [进行中] -> [已完成]

# ================= 配置区域 =================
EXCEL_PATH = r'/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/在建项目-各项目成本统计表.xlsx' 
BASE_DIR = r'/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_MATCH_DB_PATH = r'/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json'
LOG_PATH = os.path.join(SCRIPT_DIR, "file_migration_log.txt")

# 运行模式：True (从Excel生成临时配置), False (从project_name_match_db.json读取配置)
REFRESH_JSON_FROM_EXCEL = False 
# ===========================================

def log_message(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(msg)

def excel_to_json():
    """读取Excel并生成带项目名称的JSON配置"""
    try:
        # 表头在第二行 (header=1)
        df = pd.read_excel(EXCEL_PATH, header=1)
        
        # 确保列名匹配，如果您的Excel列名略有不同，请在此修改
        # 预期的列：['项目编码', '项目名称', '备注']
        required_columns = ['项目编码', '项目名称', '备注']
        for col in required_columns:
            if col not in df.columns:
                log_message(f"[错误] Excel中未找到列: {col}")
                return None

        # 清洗数据：去掉编码为空的行
        df = df[required_columns].dropna(subset=['项目编码'])
        
        config_list = []
        for _, row in df.iterrows():
            config_list.append({
                "project_code": str(row['项目编码']).strip(),
                "project_name": str(row['项目名称']).strip(), # 新增项目名称
                "target_status": str(row['备注']).strip()
            })
        
        log_message(f"[系统] 已从 Excel 生成临时配置（未写入文件）。")
        return config_list
    except Exception as e:
        log_message(f"[错误] Excel转换JSON失败: {e}")
        return None


def load_config_from_project_match_db(db_path):
    """从 project_name_match_db.json 的 projects 节点读取迁移配置。"""
    if not os.path.exists(db_path):
        log_message(f"[错误] 找不到项目匹配数据库: {db_path}")
        return None

    try:
        with open(db_path, 'r', encoding='utf-8') as f:
            db = json.load(f)

        projects = db.get('projects', {}) if isinstance(db, dict) else {}
        config_list = []
        for project_name, meta in projects.items():
            if not isinstance(meta, dict):
                continue

            project_code = str(meta.get('project_code', '')).strip()
            target_status = str(meta.get('target_status', '')).strip()
            std_name = str(project_name).strip()

            # 仅加载有编码和目标状态的条目
            if project_code and target_status and project_code != '需手动核对' and target_status != '需手动核对':
                config_list.append({
                    'project_code': project_code,
                    'project_name': std_name,
                    'target_status': target_status,
                })

        log_message(f"[系统] 已从 project_name_match_db 加载 {len(config_list)} 条可执行配置。")
        return config_list
    except Exception as e:
        log_message(f"[错误] 读取 project_name_match_db 失败: {e}")
        return None

def execute_migration(config_data):
    """执行整理工作"""
    log_message("--- 开始执行文件整理任务 ---")
    
    for item in config_data:
        code = item['project_code']
        name = item['project_name']
        target = item['target_status']
        
        found_folder_name = None
        current_parent_dir = None

        # 搜索现有层级 (BASE_DIR -> 状态目录 -> 项目目录)
        for status_dir in os.listdir(BASE_DIR):
            status_path = os.path.join(BASE_DIR, status_dir)
            if os.path.isdir(status_path) and not status_dir.startswith(('.', '5_合同管理代码')):
                for sub_dir in os.listdir(status_path):
                    # 匹配逻辑：文件夹以项目编码开头
                    if sub_dir.startswith(code):
                        found_folder_name = sub_dir
                        current_parent_dir = status_dir
                        break
            if found_folder_name: break

        if found_folder_name:
            if current_parent_dir is None:
                log_message(f"[错误] 项目目录状态识别失败: {code} | {name}")
                continue
            source = os.path.join(BASE_DIR, current_parent_dir, found_folder_name)
            destination_dir = os.path.join(BASE_DIR, target)
            destination = os.path.join(destination_dir, found_folder_name)

            if current_parent_dir == target:
                log_message(f"[已在位] 项目: {code} | {name}")
                continue

            try:
                if not os.path.exists(destination_dir):
                    os.makedirs(destination_dir)
                
                shutil.move(source, destination)
                log_message(f"[成功] 移动项目: {code} | {name} 从 [{current_parent_dir}] -> [{target}]")
            except Exception as e:
                log_message(f"[错误] 移动 {code} 失败: {e}")
        else:
            log_message(f"[缺失] 本地未找到项目文件夹: {code} ({name})")

def main():
    if REFRESH_JSON_FROM_EXCEL:
        data = excel_to_json()
    else:
        data = load_config_from_project_match_db(PROJECT_MATCH_DB_PATH)
        if data is None:
            return

    if data:
        print(f"\n当前 JSON 配置预览 (前3条):")
        for item in data[:3]:
            print(f"  - {item['project_code']}: {item['project_name']} -> {item['target_status']}")
        
        confirm = input(f"\n共加载 {len(data)} 条配置。是否开始物理移动文件？(y/n): ")
        if confirm.lower() == 'y':
            execute_migration(data)
        else:
            log_message("[取消] 用户终止了操作。")

if __name__ == "__main__":
    main()