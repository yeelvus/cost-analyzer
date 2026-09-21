# -*- coding: utf-8 -*-
"""
签证文件夹智能重命名神器 v5.0
功能：
1. 选择签证主文件夹
2. 自动读取每个子文件夹内的 txt 文件（金额、名称、日期等）
3. 按你自定义的模板，一键重命名所有子文件夹
4. 金额自动去零、避免重名冲突、完全安全可逆
5. Mac / Windows 通用，带图形界面，运行即见效果
"""

import os
import re
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askdirectory
from tkinter import messagebox
import platform
import subprocess

# ==================== 配置区：你只需要改这里！===================
# 支持任意组合命名方式，下面提供6种常用模板，直接取消注释你想要的即可！

NAMING_TEMPLATE = "{金额}_{原文件夹名}"                            # 强烈推荐！金额在前，排序最爽
# NAMING_TEMPLATE = "{金额}_{日期}_{名称}_签证"
# NAMING_TEMPLATE = "{日期}_{金额}_{名称}"
# NAMING_TEMPLATE = "{金额}_{名称}_{完成日期}（已完成）"
# NAMING_TEMPLATE = "{名称}_{金额}元_{日期}"
# NAMING_TEMPLATE = "{原文件夹名}_{金额}"

# 可自由扩展 txt 文件映射（文件名 → 在模板中使用的变量名）
TXT_MAPPING = {
    "签证金额.txt": "金额",
    "签证名称.txt": "名称",
    "签证日期.txt": "日期",
    "签证完成日期.txt": "完成日期",
    "是否签证确认.txt": "确认状态",   # 示例：内容“已确认”会直接显示
    # "办理人.txt": "办理人",         # 需要的话在这里加
}

# 是否自动清理金额后面的 .00 或 .0
CLEAN_AMOUNT_ZERO = True

# 是否保留原文件夹名（强烈建议保留，信息最全）
KEEP_ORIGINAL_NAME = True
# ==============================================================

def clean_amount(amount_str: str) -> str:
    """清理金额：去掉多余的0和小数点"""
    if not amount_str:
        return ""
    try:
        num = float(amount_str)
        if num == int(num):
            return str(int(num))
        else:
            return str(num).rstrip("0").rstrip(".") if "." in str(num) else str(int(num))
    except:
        return amount_str.strip()

def format_date(date_str: str) -> str:
    """统一日期格式：支持 2025/03/15, 2025-03-15, 2025.03.15 → 2025-03-15"""
    if not date_str:
        return ""
    date_str = date_str.strip().replace("/", "-").replace(".", "-")
    # 简单处理常见格式
    match = re.search(r'\d{4}-\d{1,2}-\d{1,2}', date_str)
    if match:
        return match.group(0)
    return date_str

def avoid_name_conflict(parent_path: Path, new_name: str) -> str:
    """避免重名：如果已存在，加 (1)(2)..."""
    candidate = new_name
    counter = 1
    while (parent_path / candidate).exists():
        candidate = f"{new_name} ({counter})"
        counter += 1
    return candidate

def open_folder(path):
    """跨平台打开文件夹"""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except:
        pass

def main():
    Tk().withdraw()
    folder = askdirectory(title="请选择签证主文件夹（里面全是签证子文件夹）")
    if not folder:
        return

    base_path = Path(folder)
    renamed_count = 0
    skipped_count = 0
    details = []

    print("开始重命名文件夹...\n")

    for subfolder in base_path.iterdir():
        if not subfolder.is_dir():
            continue

        data = {"原文件夹名": subfolder.name}

        # 读取所有 txt 文件
        for txt_file, var_name in TXT_MAPPING.items():
            txt_path = subfolder / txt_file
            if txt_path.exists():
                content = txt_path.read_text(encoding="utf-8-sig").strip()
                if txt_file == "签证金额.txt" and CLEAN_AMOUNT_ZERO:
                    content = clean_amount(content)
                elif "日期" in txt_file:
                    content = format_date(content)
                data[var_name] = content if content else ""

        # 必须有金额才重命名（可改成其他条件）
        if not data.get("金额"):
            print(f"跳过（无金额）: {subfolder.name}")
            skipped_count += 1
            continue

        # 生成新名字
        try:
            new_name = NAMING_TEMPLATE.format(**data)
            if KEEP_ORIGINAL_NAME and "{原文件夹名}" not in NAMING_TEMPLATE:
                new_name = f"{new_name}_{subfolder.name}" if new_name else subfolder.name
        except KeyError as e:
            print(f"模板变量错误 {e}，跳过: {subfolder.name}")
            skipped_count += 1
            continue

        # 避免冲突
        final_name = avoid_name_conflict(base_path, new_name)
        if final_name != subfolder.name:
            subfolder.rename(base_path / final_name)
            renamed_count += 1
            details.append(f"{subfolder.name} → {final_name}")
            print(f"成功: {subfolder.name} → {final_name}")
        else:
            print(f"无需改名: {subfolder.name}")
            skipped_count += 1

    # 结果提示
    summary = f"重命名完成！\n\n成功重命名：{renamed_count} 个\n跳过：{skipped_count} 个\n\n"
    if renamed_count > 0:
        summary += "部分示例：\n" + "\n".join(details[:10])
        if len(details) > 10:
            summary += f"\n... 等共 {renamed_count} 个"
    
    messagebox.showinfo("大功告成！", summary)
    open_folder(folder)

if __name__ == "__main__":
    main()
    print("\n程序结束，按回车退出...")
    input()