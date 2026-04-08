import os
import shutil
import hashlib
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, MULTIPLE

# ==============================================================
# 工具函数
# ==============================================================
def select_folder(prompt):
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title=prompt)
    root.destroy()
    return folder

def file_hash(filepath, block_size=65536):
    hasher = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(block_size), b''):
                hasher.update(chunk)
        return hasher.hexdigest()
    except:
        return None

def copy_with_skip(src_dir, dst_dir, log_list):
    os.makedirs(dst_dir, exist_ok=True)
    copied = False
    for item in os.listdir(src_dir):
        s_path = os.path.join(src_dir, item)
        d_path = os.path.join(dst_dir, item)
        try:
            if os.path.isdir(s_path):
                if copy_with_skip(s_path, d_path, log_list):
                    copied = True
            else:
                need_copy = True
                if os.path.exists(d_path):
                    if file_hash(s_path) == file_hash(d_path):
                        print(f"跳过: {s_path}")
                        need_copy = False
                    else:
                        print(f"覆盖: {s_path} → {d_path}")
                if need_copy:
                    shutil.copy2(s_path, d_path)
                    log_list.append((s_path, d_path))
                    print(f"复制: {s_path} → {d_path}")
                    copied = True
        except Exception as e:
            print(f"失败: {s_path} | {e}")
    return copied

def generate_log(data, output_root, log_type="copy"):
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(output_root, f"{log_type}_log_{ts}.txt")
    with open(log_path, 'w', encoding='utf-8') as f:
        if log_type == "copy":
            f.write(f"# 复制日志 - {datetime.datetime.now()}\n")
            f.write(f"# 共 {len(data)} 个文件\n\n")
            for i, (s, d) in enumerate(data, 1):
                f.write(f"{i}. {s}\n   → {d}\n")
        elif log_type == "match":
            f.write(f"# 匹配列表 - {datetime.datetime.now()}\n\n")
            for src_proj, src_path, dst_proj, dst_path in data:
                f.write(f"项目: {src_proj}\n")
                f.write(f"  源总包: {src_path}\n")
                f.write(f"  目标文件夹: {dst_path}\n\n")
    print(f"日志: {log_path}")
    return log_path

def save_input_tree(input_root, output_root):
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    tree_path = os.path.join(output_root, f"input_tree_{ts}.txt")
    try:
        with open(tree_path, 'w', encoding='utf-8') as f:
            f.write(f"输入层级结构 - {input_root}\n")
            f.write(f"生成时间: {datetime.datetime.now()}\n\n")
            _print_tree(input_root, "", f)
        print(f"层级保存: {tree_path}")
    except Exception as e:
        print(f"层级失败: {e}")

def _print_tree(path, prefix="", file=None):
    try:
        items = sorted([i for i in os.listdir(path) if not i.startswith('.')])
        for i, item in enumerate(items):
            conn = "├── " if i < len(items)-1 else "└── "
            line = f"{prefix}{conn}{item}"
            print(line)
            if file: file.write(line + "\n")
            subpath = os.path.join(path, item)
            if os.path.isdir(subpath):
                ext = "│   " if i < len(items)-1 else "    "
                _print_tree(subpath, prefix + ext, file)
    except PermissionError:
        pass

# ==============================================================
# 主程序
# ==============================================================
def main():
    # 1. 选择文件夹
    input_root = select_folder("【输入】选择包含总包合同的根目录")
    if not input_root: return
    output_root = select_folder("【输出】选择输出根目录")
    if not output_root: return

    # 2. 保存输入层级
    save_input_tree(input_root, output_root)

    # 3. 扫描输入端：找 总包合同
    print("\n扫描输入端...")
    source_contracts = []
    for cur_dir, dirs, _ in os.walk(input_root):
        for d in dirs:
            if d == "总包合同":
                contract_path = os.path.join(cur_dir, d)
                project_name = os.path.basename(cur_dir)
                source_contracts.append((project_name, contract_path))
                print(f"  发现: {project_name}")

    if not source_contracts:
        messagebox.showinfo("完成", "输入端未找到 总包合同 文件夹")
        return

    # 4. 扫描输出端：找同名项目 + 含“总包合同”字段的子文件夹
    print("\n扫描输出端...")
    matches = []
    for src_proj, src_path in source_contracts:
        found_target = False
        for out_dir, out_subdirs, _ in os.walk(output_root):
            out_proj = os.path.basename(out_dir)
            if out_proj != src_proj:
                continue
            # 在此项目下找含“总包合同”的文件夹
            target_subdir = None
            for sub in out_subdirs:
                if "总包合同" in sub:
                    target_subdir = os.path.join(out_dir, sub)
                    break
            if target_subdir:
                matches.append((src_proj, src_path, out_dir, target_subdir))
                print(f"  匹配: {src_proj} → {target_subdir}")
                found_target = True
                break
        if not found_target:
            print(f"  无匹配: {src_proj}")

    if not matches:
        messagebox.showinfo("完成", "未找到任何匹配的项目文件夹")
        return

    # 5. 弹窗确认所有匹配
    confirm_win = tk.Toplevel()
    confirm_win.title("确认复制任务")
    confirm_win.geometry("800x500")

    tk.Label(confirm_win, text="请勾选需要复制的项目（默认全选）", font=("Arial", 10)).pack(pady=5)

    listbox = Listbox(confirm_win, selectmode=MULTIPLE, font=("Consolas", 9))
    scrollbar = Scrollbar(confirm_win, orient="vertical")
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox.yview)
    listbox.pack(side="left", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    indices = []
    for i, (s_proj, s_path, d_proj, d_path) in enumerate(matches):
        text = f"{s_proj}\n  源: {s_path}\n  目标: {d_path}"
        listbox.insert(tk.END, text)
        indices.append(i)
    for i in range(len(matches)):
        listbox.select_set(i)  # 默认全选

    result = []

    def on_confirm():
        selected = listbox.curselection()
        for idx in selected:
            result.append(matches[idx])
        confirm_win.destroy()

    def on_cancel():
        confirm_win.destroy()

    btn_frame = tk.Frame(confirm_win)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="确认复制", command=on_confirm, bg="lightgreen").pack(side="left", padx=5)
    tk.Button(btn_frame, text="取消", command=on_cancel, bg="lightcoral").pack(side="left", padx=5)

    confirm_win.wait_window()

    if not result:
        messagebox.showinfo("取消", "操作已取消")
        return

    # 6. 执行复制
    log_list = []
    copied_count = 0
    for src_proj, src_path, _, dst_path in result:
        print(f"\n复制: {src_proj}")
        if copy_with_skip(src_path, dst_path, log_list):
            copied_count += 1

    # 7. 生成日志 + 完成
    match_log = generate_log(matches, output_root, "match")
    copy_log = generate_log(log_list, output_root, "copy")

    summary = (
        f"复制完成！\n\n"
        f"匹配项目: {len(matches)}\n"
        f"执行复制: {len(result)}\n"
        f"成功项目: {copied_count}\n"
        f"文件数: {len(log_list)}\n\n"
        f"日志:\n- 匹配: {match_log}\n- 复制: {copy_log}"
    )
    messagebox.showinfo("成功", summary)

if __name__ == "__main__":
    main()