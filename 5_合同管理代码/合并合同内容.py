import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox

# ================= 配置区域 =================
# 指向“5_合同管理”这一层
BASE_DIR = r'/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理'
# ===========================================

class UniversalMergeTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ContractMerge-Pro (跨状态合并版)")
        self.root.geometry("800x600")
        
        # 存储项目信息：{ "显示名称": "完整路径" }
        self.project_map = {}
        self.scan_all_projects()
        
        # UI 界面
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 选择框区域
        tk.Label(main_frame, text="1. 选择【主项目】 (保留此文件夹及所在状态位置)", font=('微软雅黑', 10, 'bold')).pack(anchor="w")
        self.main_cb = ttk.Combobox(main_frame, values=list(self.project_map.keys()), width=90)
        self.main_cb.pack(pady=5)

        tk.Label(main_frame, text="2. 选择【被合并项目】 (内容将并入上方项目)", font=('微软雅黑', 10, 'bold'), fg="red").pack(anchor="w", pady=(10, 0))
        self.sub_cb = ttk.Combobox(main_frame, values=list(self.project_map.keys()), width=90)
        self.sub_cb.pack(pady=5)

        # 按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="刷新项目列表", command=self.refresh_list).pack(side=tk.LEFT, padx=10)
        self.preview_btn = tk.Button(btn_frame, text="生成合并预览清单", command=self.preview_merge, bg="#2196F3", fg="white")
        self.preview_btn.pack(side=tk.LEFT, padx=10)

        # 日志预览
        tk.Label(main_frame, text="操作预览及执行日志:").pack(anchor="w")
        self.log_box = tk.Text(main_frame, height=18, width=100, bg="#f8f8f8", font=('Consolas', 9))
        self.log_box.pack(pady=5)

        self.run_btn = tk.Button(main_frame, text="确认并执行物理合并", command=self.run_merge, bg="#4CAF50", fg="white", state=tk.DISABLED, height=2, width=20)
        self.run_btn.pack(pady=10)

    def scan_all_projects(self):
        """扫描 BASE_DIR 下所有子文件夹中的项目文件夹"""
        self.project_map = {}
        if not os.path.exists(BASE_DIR):
            return

        # 遍历第一层：A_在建项目, D_已完工等
        for status_dir in os.listdir(BASE_DIR):
            status_path = os.path.join(BASE_DIR, status_dir)
            
            # 排除非文件夹和代码目录
            if os.path.isdir(status_path) and "代码" not in status_dir and not status_dir.startswith('.'):
                # 遍历第二层：项目文件夹
                for proj_dir in os.listdir(status_path):
                    proj_path = os.path.join(status_path, proj_dir)
                    # 识别符合编码规则的项目文件夹 (例如包含 THCC-)
                    if os.path.isdir(proj_path) and "THCC-" in proj_dir:
                        display_name = f"[{status_dir}] {proj_dir}"
                        self.project_map[display_name] = proj_path

    def refresh_list(self):
        self.scan_all_projects()
        self.main_cb['values'] = list(self.project_map.keys())
        self.sub_cb['values'] = list(self.project_map.keys())
        messagebox.showinfo("刷新", "已重新扫描目录结构")

    def get_prefix(self, folder_name):
        return folder_name.split('_')[0] if '_' in folder_name else None

    def preview_merge(self):
        main_key = self.main_cb.get()
        sub_key = self.sub_cb.get()
        
        if not main_key or not sub_key or main_key == sub_key:
            messagebox.showwarning("警告", "请选择两个不同的项目进行合并操作")
            return

        self.log_box.delete('1.0', tk.END)
        self.merge_tasks = [] # 存储 (源文件夹全路径, 目标文件夹全路径)
        
        main_path = self.project_map[main_key]
        sub_path = self.project_map[sub_key]

        self.log_box.insert(tk.END, f">>> 准备从以下路径合并内容 <<<\n")
        self.log_box.insert(tk.END, f"源路径: {sub_path}\n")
        self.log_box.insert(tk.END, f"目标路径: {main_path}\n\n")

        # 扫描被合并项目下的分类子文件夹 (A, E, F...)
        for folder in os.listdir(sub_path):
            if folder.startswith('.'): continue
            sub_prefix = self.get_prefix(folder)
            if not sub_prefix: continue

            # 在主项目中匹配同前缀的文件夹
            match_found = False
            for main_folder in os.listdir(main_path):
                if main_folder.startswith(sub_prefix + "_"):
                    src = os.path.join(sub_path, folder)
                    dst = os.path.join(main_path, main_folder)
                    items = os.listdir(src)
                    self.log_box.insert(tk.END, f"[分类对齐] {sub_prefix}: 发现 {len(items)} 个项 将并入 {main_folder}\n")
                    self.merge_tasks.append((src, dst))
                    match_found = True
                    break
            
            if not match_found:
                # 如果主项目没这个分类，直接把整个分类文件夹移过去
                src = os.path.join(sub_path, folder)
                dst = os.path.join(main_path, folder)
                self.log_box.insert(tk.END, f"[全量新增] {sub_prefix}: 主项目无此分类，将完整迁移文件夹\n")
                self.merge_tasks.append((src, dst))

        self.run_btn.config(state=tk.NORMAL)

    def run_merge(self):
        if not messagebox.askyesno("二次确认", "点击确定后，所有内容将被物理搬迁。被合并项目的空壳将被删除。是否继续？"):
            return

        try:
            for src, dst in self.merge_tasks:
                if not os.path.exists(dst):
                    os.makedirs(dst)
                
                for item in os.listdir(src):
                    s_file = os.path.join(src, item)
                    d_file = os.path.join(dst, item)
                    
                    # 避免同名文件直接覆盖，若冲突则自动加前缀
                    if os.path.exists(d_file):
                        d_file = os.path.join(dst, f"Merged_{item}")
                    
                    shutil.move(s_file, d_file)
            
            # 最后删除被合并项目的根目录
            sub_path = self.project_map[self.sub_cb.get()]
            shutil.rmtree(sub_path)
            
            messagebox.showinfo("成功", "合并已完成。")
            self.refresh_list()
            self.log_box.insert(tk.END, "\n[完成] 所有任务已成功执行。")
        except Exception as e:
            messagebox.showerror("致命错误", f"合并过程中发生异常: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalMergeTool(root)
    root.mainloop()