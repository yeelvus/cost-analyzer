# -*- coding: utf-8 -*-
"""
签证自动化终极工具 v6.0（金额前置重命名完全兼容版）
功能：
1. 自动补齐缺失txt
2. 一键生成月报（支持任意重命名方式）
3. Excel中带超链接，双击直达文件夹
4. 识别标准：只要有“签证金额.txt”就算有效签证
"""

import os
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# 必须包含这个文件才算有效签证文件夹
KEY_REQUIRED_FILE = "签证金额.txt"

# 需要读取的所有txt文件（可无限扩展）
TXT_FILES_TO_READ = [
    "签证名称.txt",
    "签证日期.txt",
    "签证完成日期.txt",
    "签证金额.txt",
    "是否签证确认.txt",
    # "办理人.txt",        # 需要的话取消注释
    # "备注.txt",
]

# ==================== 判断是否为有效签证文件夹（新逻辑）===================
def is_valid_visa_folder(subfolder: Path) -> bool:
    return (subfolder / KEY_REQUIRED_FILE).exists()

# ==================== 第一部分：智能补全缺失txt ====================
def smart_check_and_complete():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="请选择签证主文件夹")
    if not folder:
        return
    base_path = Path(folder)
    
    total = 0
    created = 0
    created_list = []

    for sub in base_path.iterdir():
        if not sub.is_dir():
            continue
        if not is_valid_visa_folder(sub):
            continue
        
        total += 1
        for txt_name in TXT_FILES_TO_READ:
            txt_path = sub / txt_name
            if not txt_path.exists():
                txt_path.write_text("", encoding="utf-8-sig")
                created += 1
                created_list.append(f"{sub.name}/{txt_name}")

    if created == 0:
        messagebox.showinfo("完成", f"检查了 {total} 个签证文件夹\n全部齐全，无需创建！")
    else:
        msg = f"补全成功！新建 {created} 个文件：\n\n"
        for i in created_list[:20]:
            msg += f"• {i}\n"
        if len(created_list) > 20:
            msg += f"... 等共 {created} 个"
        messagebox.showinfo("补全完成", msg)

# ==================== 第二部分：生成带超链接的月报 ====================
def generate_report():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="请选择签证主文件夹生成报表")
    if not folder:
        return
    base_path = Path(folder)
    
    data = []
    for sub in base_path.iterdir():
        if not sub.is_dir():
            continue
        if not is_valid_visa_folder(sub):
            continue
        
        row = {"文件夹名称": sub.name, "路径": str(sub)}
        for txt_name in TXT_FILES_TO_READ:
            txt_path = sub / txt_name
            key = txt_name.replace(".txt", "")
            if txt_path.exists():
                content = txt_path.read_text(encoding="utf-8-sig").strip()
                row[key] = content if content else "未填写"
            else:
                row[key] = "【缺失】"
        
        # 添加最后更新时间
        txt_files = [f for f in sub.glob("*.txt")]
        if txt_files:
            latest = max(f.stat().st_mtime for f in txt_files)
            row["最后更新"] = datetime.fromtimestamp(latest).strftime("%Y-%m-%d %H:%M")
        else:
            row["最后更新"] = "未知"
        
        data.append(row)

    if not data:
        messagebox.showinfo("提示", "未找到任何包含【签证金额.txt】的签证文件夹")
        return

    df = pd.DataFrame(data)
    
    # 按签证日期排序（可失败不影响）
    try:
        df["排序辅助"] = pd.to_datetime(df["签证日期"], errors="coerce")
        df = df.sort_values("排序辅助", ascending=False).drop("排序辅助", axis=1)
    except:
        pass

    timestamp = datetime.now().strftime("%Y年%m月")
    output_file = base_path / f"{timestamp}_签证办理月度报表（带链接）.xlsx"

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 对外简洁版
        df_simple = df[["文件夹名称", "签证名称", "签证日期", "签证完成日期", "签证金额", "是否签证确认"]].copy()
        df_simple.to_excel(writer, sheet_name="对外报表", index=False, startrow=2)
        # 完整版
        df.to_excel(writer, sheet_name="详细完整版", index=False, startrow=2)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            
            # 大标题
            title = f"签证办理进度月度报表 - {timestamp}"
            ws.merge_cells(f"A1:{get_column_letter(ws.max_column)}1")
            ws["A1"] = title
            ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor="1976D2")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 45

            # 表头美化
            for cell in ws[3]:
                cell.fill = PatternFill("solid", fgColor="1565C0")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")

            # 为“文件夹名称”列添加超链接（详细版才有路径列）
            if "路径" in df.columns:
                path_col_idx = df.columns.get_loc("路径") + 1
                name_col_idx = df.columns.get_loc("文件夹名称") + 1
                for row_idx in range(4, len(df) + 4):
                    folder_path = ws.cell(row=row_idx, column=path_col_idx).value
                    cell = ws.cell(row=row_idx, column=name_col_idx)
                    if folder_path and os.path.exists(folder_path):
                        # Windows和Mac都支持file:///协议
                        cell.hyperlink = f"file:///{folder_path.replace(os.sep, '/')}"
                        cell.font = Font(color="0000EE", underline="single")
                        cell.value = "→ " + cell.value

            # 自动列宽
            for i, col in enumerate(ws.columns, 1):
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[get_column_letter(i)].width = min(max_len + 5, 60)

    # 自动打开Excel
    try:
        if os.name == 'nt':
            os.startfile(output_file)
        else:
            os.system(f'open "{output_file}"')
    except:
        pass

    messagebox.showinfo("大功告成！", f"月报已生成（带超链接，双击即可打开文件夹）\n\n{output_file}")

# ==================== 主界面 ====================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("签证自动化工具 v6.0")
    root.geometry("620x420")
    root.configure(bg="#f5f7fa")

    tk.Label(root, text="签证办理终极自动化工具", font=("微软雅黑", 22, "bold"), 
             bg="#f5f7fa", fg="#1565C0").pack(pady=35)
    
    tk.Label(root, text="完美兼容“金额前置”重命名格式", font=("微软雅黑", 11), 
             bg="#f5f7fa", fg="#2E7D32").pack(pady=8)

    tk.Button(root, text="1. 一键补全缺失的txt文件\n（永不覆盖已有内容）", 
              height=3, width=70, bg="#2E7D32", fg="white", font=("微软雅黑", 12, "bold"),
              command=smart_check_and_complete).pack(pady=20)

    tk.Button(root, text="2. 生成月报（带超链接，双击直达文件夹）\n支持任意重命名方式", 
              height=3, width=70, bg="#D81B60", fg="white", font=("微软雅黑", 12, "bold"),
              command=generate_report).pack(pady=15)

    tk.Label(root, text="v6.0 · 金额前置完全兼容 · 2025-12-03 · Excel双击即开文件夹", 
             font=("微软雅黑", 9), fg="#666", bg="#f5f7fa").pack(side="bottom", pady=25)

    root.mainloop()