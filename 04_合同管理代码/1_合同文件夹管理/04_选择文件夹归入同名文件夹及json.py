#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
把所选文件夹里直接存放的文件，各自移入同名文件夹，并写入 JSON。

合同名称 = 文件名（不含扩展名）。其余字段留空；若 JSON 已存在，只改合同名称，其它已填内容保留。
新增模板：在 TEMPLATES 里加一项即可，界面下拉框会自动出现。
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import threading
from collections import OrderedDict
from pathlib import Path
from urllib.parse import unquote, urlparse

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ModuleNotFoundError:
    tk = None
    filedialog = messagebox = ttk = None

DEFAULT_BROWSE = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑")

# 模板。name_field 会写成文件名（不含扩展名）。
TEMPLATES: dict[str, dict] = {
    "default": {
        "label": "默认（合同信息）",
        "filename": "合同信息.json",
        "name_field": "合同名称",
        "fields": [
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
            "备注",
        ],
    },
}

_ILLEGAL_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_HIDDEN_NAMES = {"Thumbs.db", "desktop.ini", "__MACOSX"}


def parse_folder_text(text: str) -> Path:
    raw = text.strip().strip('"').strip("'")
    if raw.lower().startswith("file://"):
        raw = unquote(urlparse(raw).path)
    return Path(raw).expanduser()


def is_hidden(name: str) -> bool:
    return name.startswith(".") or name in _HIDDEN_NAMES


def sanitize_folder_name(name: str) -> str:
    cleaned = _ILLEGAL_NAME.sub("_", name).strip().rstrip(".")
    return cleaned or "未命名"


def list_loose_files(folder: Path) -> list[Path]:
    files = [p for p in folder.iterdir() if p.is_file() and not is_hidden(p.name)]
    return sorted(files, key=lambda p: p.name.casefold())


def resolve_target_dir(parent: Path, folder_name: str) -> Path:
    candidate = parent / folder_name
    if candidate.is_dir() or not candidate.exists():
        return candidate
    index = 1
    while True:
        alternative = parent / f"{folder_name}_{index}"
        if alternative.is_dir() or not alternative.exists():
            return alternative
        index += 1


def unique_file(path: Path) -> Path:
    if not path.exists():
        return path
    index = 1
    while True:
        alternative = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not alternative.exists():
            return alternative
        index += 1


def build_payload(template: dict, contract_name: str, existing: dict | None) -> OrderedDict:
    data = dict(existing or {})
    name_field = template["name_field"]
    ordered: OrderedDict = OrderedDict()
    for key in template["fields"]:
        if key == name_field:
            ordered[key] = contract_name
        else:
            value = data.get(key, "")
            ordered[key] = "" if value is None else value
    for key, value in data.items():
        if key not in ordered:
            ordered[key] = value
    return ordered


def write_json(folder: Path, template: dict, contract_name: str) -> str:
    """写入 JSON。返回日志短句。"""
    json_path = folder / template["filename"]
    existing = None
    note = "新建"
    if json_path.exists():
        try:
            with json_path.open("r", encoding="utf-8") as handle:
                loaded = json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"已有 JSON 无法读取，未覆盖: {json_path.name}（{exc}）") from exc
        if not isinstance(loaded, dict):
            raise RuntimeError(f"已有 JSON 不是对象，未覆盖: {json_path.name}")
        existing = loaded
        previous = existing.get(template["name_field"], "")
        note = "更新" if previous != contract_name else "保持"
    payload = build_payload(template, contract_name, existing)
    with json_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return f"{note} {json_path.name}（{template['name_field']} = {contract_name}）"


def _ensure_target_dir(
    folder: Path,
    source: Path,
    folder_name: str,
    original_name: str,
    contract_name: str,
) -> tuple[Path, Path, bool, str]:
    """返回 (目标文件夹, 待移动文件, 是否新建, 日志)。无扩展名文件与文件夹同名时先挪开再建夹。"""
    desired = folder / folder_name
    if desired.exists() and desired.resolve() == source.resolve():
        temporary = unique_file(folder / f"{original_name}.moving")
        shutil.move(str(source), str(temporary))
        try:
            desired.mkdir(parents=False, exist_ok=False)
        except Exception:
            shutil.move(str(temporary), str(folder / original_name))
            raise
        return desired, temporary, True, f"新建文件夹: {desired.name}"

    if desired.is_dir():
        return desired, source, False, f"放入已有文件夹: {desired.name}"

    if desired.exists():
        target_dir = resolve_target_dir(folder, folder_name)
        created = False
        note = f"同名路径被文件占用，改用: {target_dir.name}"
        if not target_dir.exists():
            target_dir.mkdir(parents=False, exist_ok=False)
            created = True
        elif target_dir.is_dir():
            note = f"同名路径被文件占用，放入已有文件夹: {target_dir.name}"
        return target_dir, source, created, note

    desired.mkdir(parents=False, exist_ok=False)
    note = f"新建文件夹: {desired.name}"
    if folder_name != contract_name:
        note += f"（由 {contract_name} 替换非法字符）"
    return desired, source, True, note


def _rollback_failed_file(
    folder: Path,
    target_dir: Path,
    source: Path,
    original_name: str,
    created: bool,
    json_name: str,
) -> bool:
    """新建文件夹若还没有放进原文件，就删掉本次写出的 JSON 和空文件夹。临时文件放回原名。"""
    removed = False
    if created and target_dir.is_dir():
        json_path = target_dir / json_name
        contents = list(target_dir.iterdir())
        if json_path in contents and all(item == json_path for item in contents):
            json_path.unlink()
            contents = list(target_dir.iterdir())
        if not contents:
            target_dir.rmdir()
            removed = True
    original = folder / original_name
    if source.exists() and source != original and not original.exists():
        shutil.move(str(source), str(original))
    return removed


def process_folder(folder: Path, template_key: str) -> dict:
    if template_key not in TEMPLATES:
        raise KeyError(f"未知模板: {template_key}")
    if not folder.is_dir():
        raise NotADirectoryError(f"不是文件夹: {folder}")

    template = TEMPLATES[template_key]
    stats = {
        "folders_created": 0,
        "files_moved": 0,
        "json_written": 0,
        "failed": 0,
        "logs": [],
    }
    files = list_loose_files(folder)
    stats["logs"].append(f"待处理文件 {len(files)} 个")

    for source in files:
        original_name = source.name
        contract_name = source.stem.strip() or "未命名"
        folder_name = sanitize_folder_name(contract_name)
        target_dir = folder / folder_name
        created = False
        json_done = False
        try:
            target_dir, source, created, note = _ensure_target_dir(
                folder, source, folder_name, original_name, contract_name
            )
            if created:
                stats["folders_created"] += 1
            stats["logs"].append(note)

            json_note = write_json(target_dir, template, contract_name)
            json_done = True
            stats["json_written"] += 1
            stats["logs"].append(f"JSON {json_note}")

            destination = unique_file(target_dir / original_name)
            shutil.move(str(source), str(destination))
            stats["files_moved"] += 1
            moved = original_name if destination.name == original_name else f"{original_name} → {destination.name}"
            stats["logs"].append(f"移动: {moved} → {target_dir.name}/")
        except Exception as exc:
            stats["failed"] += 1
            stats["logs"].append(f"失败: {original_name}（{exc}）")
            removed_folder = _rollback_failed_file(
                folder, target_dir, source, original_name, created, template["filename"]
            )
            if removed_folder and created:
                stats["folders_created"] -= 1
                if json_done:
                    stats["json_written"] -= 1
    return stats


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("文件归入同名文件夹")
        self.root.geometry("760x560")
        self.root.minsize(640, 480)

        self.folder_path = tk.StringVar()
        self.template_label = tk.StringVar(value=TEMPLATES["default"]["label"])
        self._running = False
        self._label_to_key = {item["label"]: key for key, item in TEMPLATES.items()}

        self._build()

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 4}
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="选择或粘贴一个文件夹。其中直接存放的每个文件会移入同名文件夹，并写入 JSON。合同名称等于文件名（不含扩展名）。",
            wraplength=700,
        ).grid(row=0, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(main, text="文件夹").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(main, textvariable=self.folder_path).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(main, text="浏览…", command=self._browse).grid(row=1, column=2, sticky="w", **pad)

        ttk.Label(main, text="JSON 模板").grid(row=2, column=0, sticky="e", **pad)
        self.template_combo = ttk.Combobox(
            main,
            textvariable=self.template_label,
            values=list(self._label_to_key),
            state="readonly",
            width=28,
        )
        self.template_combo.grid(row=2, column=1, sticky="w", **pad)
        self.template_hint = ttk.Label(main, foreground="#555")
        self.template_hint.grid(row=3, column=1, columnspan=2, sticky="w", **pad)
        self.template_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_hint())
        self._refresh_hint()

        buttons = ttk.Frame(main)
        buttons.grid(row=4, column=0, columnspan=3, pady=8)
        self.run_button = ttk.Button(buttons, text="开始处理", command=self._run)
        self.run_button.pack(side="left", padx=8)
        ttk.Button(buttons, text="退出", command=self.root.destroy).pack(side="left", padx=8)

        ttk.Label(main, text="日志").grid(row=5, column=0, sticky="nw", **pad)
        log_frame = ttk.Frame(main)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", **pad)
        self.log = tk.Text(log_frame, height=16, wrap="word", font=("Menlo", 11))
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        main.columnconfigure(1, weight=1)
        main.rowconfigure(6, weight=1)
        self._append_log("可粘贴路径，或点「浏览…」选择文件夹。子文件夹里的文件不会动。")

    def _refresh_hint(self) -> None:
        template = TEMPLATES[self._selected_template_key()]
        self.template_hint.configure(
            text=f"将写入 {template['filename']}，{template['name_field']} = 文件名（不含扩展名）"
        )

    def _selected_template_key(self) -> str:
        return self._label_to_key.get(self.template_label.get(), "default")

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def _browse(self) -> None:
        current = self.folder_path.get().strip()
        initial = DEFAULT_BROWSE if DEFAULT_BROWSE.is_dir() else Path.home()
        if current:
            try:
                candidate = parse_folder_text(current)
            except Exception:
                candidate = None
            if candidate and candidate.is_dir():
                initial = candidate
            elif candidate and candidate.parent.is_dir():
                initial = candidate.parent
        selected = filedialog.askdirectory(title="选择要整理的文件夹", initialdir=str(initial))
        if selected:
            self.folder_path.set(selected)

    def _run(self) -> None:
        if self._running:
            return
        try:
            folder = parse_folder_text(self.folder_path.get())
        except Exception as exc:
            messagebox.showerror("路径无效", str(exc))
            return
        if not self.folder_path.get().strip():
            messagebox.showwarning("未选择文件夹", "请粘贴或选择一个文件夹。")
            return
        if not folder.is_dir():
            messagebox.showerror("路径无效", f"找不到这个文件夹:\n{folder}")
            return

        files = list_loose_files(folder)
        template = TEMPLATES[self._selected_template_key()]
        if not files:
            messagebox.showinfo("没有可处理的文件", "这个文件夹里没有直接存放的文件。子文件夹和隐藏文件会跳过。")
            return
        preview = "\n".join(path.name for path in files[:12])
        if len(files) > 12:
            preview += f"\n… 另有 {len(files) - 12} 个"
        agreed = messagebox.askokcancel(
            "确认处理",
            f"将处理 {len(files)} 个文件。\n"
            f"每个文件移入同名文件夹，并写入 {template['filename']}。\n"
            f"{template['name_field']} = 文件名（不含扩展名）。\n\n{preview}",
        )
        if not agreed:
            return

        self._running = True
        self.run_button.configure(state="disabled")
        self._append_log(f"\n文件夹: {folder}")
        self._append_log(f"模板: {template['label']}")
        threading.Thread(target=self._worker, args=(folder, self._selected_template_key()), daemon=True).start()

    def _worker(self, folder: Path, template_key: str) -> None:
        try:
            stats = process_folder(folder, template_key)
        except Exception as exc:
            message = str(exc)
            self.root.after(0, lambda message=message: self._finish_error(message))
            return
        self.root.after(0, lambda stats=stats: self._finish(stats))

    def _finish(self, stats: dict) -> None:
        for line in stats["logs"]:
            self._append_log(line)
        summary = (
            f"完成：新建文件夹 {stats['folders_created']}，"
            f"移动文件 {stats['files_moved']}，"
            f"写入 JSON {stats['json_written']}，"
            f"失败 {stats['failed']}"
        )
        self._append_log(summary)
        self._running = False
        self.run_button.configure(state="normal")
        if stats["failed"]:
            messagebox.showwarning("处理结束，有失败项", summary)
        else:
            messagebox.showinfo("处理完成", summary)

    def _finish_error(self, message: str) -> None:
        self._append_log(f"处理中断: {message}")
        self._running = False
        self.run_button.configure(state="normal")
        messagebox.showerror("处理失败", message)


def main() -> None:
    if tk is None:
        print(
            "当前 Python 没有图形界面组件 tkinter。\n"
            "请用系统自带的 Python 启动：\n"
            f"  /usr/bin/python3 \"{Path(__file__).resolve()}\"",
            file=sys.stderr,
        )
        raise SystemExit(1)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
