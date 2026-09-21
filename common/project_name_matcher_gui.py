# -*- coding: utf-8 -*-
"""
项目名称匹配 · 图形界面

选择 Excel、待匹配列、标准名称输出列、项目编码输出列；可选清理摘要前缀。
只改指定单元格的值，不改变原文件其它格式。

启动：
    python -m common.project_name_matcher
    python -m common.project_name_matcher gui
    python -m common.project_name_matcher_gui
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

_CODE_ROOT = Path(__file__).resolve().parents[1]
if str(_CODE_ROOT) not in sys.path:
    sys.path.insert(0, str(_CODE_ROOT))

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception as exc:  # pragma: no cover
    raise SystemExit(f"无法加载 tkinter，图形界面不可用: {exc}") from exc

from common.project_name_matcher import (
    fill_excel_project_columns,
    load_matcher,
    peek_excel_sheets_and_headers,
)

DEFAULT_EXCEL_DIR = Path("/Users/wyattsun/Desktop/财务数据")
NAME_CANDIDATES = ("项目名称", "工程名称", "原始项目名称")
NAME_OUT_CANDIDATES = ("项目标准名称", "标准项目名称", "标准名称")
CODE_OUT_CANDIDATES = ("项目编码", "工程编码", "编码")
SUMMARY_CANDIDATES = ("摘要",)
SUMMARY_OUT_CANDIDATES = ("摘要处理", "处理后摘要")


def _pick(columns, candidates) -> str:
    cols = [str(c).strip() for c in columns if str(c).strip()]
    for name in candidates:
        if name in cols:
            return name
    return cols[0] if cols else ""


class MatcherExcelApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("项目名称匹配 · 填写 Excel")
        self.root.geometry("860x720")
        self.root.minsize(760, 620)

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.header_row = tk.StringVar(value="1")
        self.match_col = tk.StringVar()
        self.name_out_col = tk.StringVar()
        self.code_out_col = tk.StringVar()
        self.summary_col = tk.StringVar()
        self.summary_out_col = tk.StringVar()
        self.process_summary = tk.BooleanVar(value=True)
        self.prompt_unmatched = tk.BooleanVar(value=True)
        self.only_empty = tk.BooleanVar(value=False)
        self.save_mode = tk.StringVar(value="overwrite")
        self.output_path = tk.StringVar()
        self.db_label = tk.StringVar()

        self._sheets: list[str] = []
        self._headers: dict[str, list[str]] = {}
        self._running = False
        self.matcher = load_matcher()
        self.db_label.set(
            f"匹配库: {self.matcher.db_path.name}　标准项目 {self.matcher.project_count} 个"
        )

        self._build()
        self._toggle_summary()
        self._toggle_output()

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 4}
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="选择 Excel，指定待匹配列与输出列。只改这些列的值，其它格式（列宽、样式、日期、筛选）保持不变。",
            wraplength=820,
        ).grid(row=0, column=0, columnspan=4, sticky="w", **pad)

        ttk.Label(main, text="Excel 文件").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(main, textvariable=self.excel_path).grid(
            row=1, column=1, columnspan=2, sticky="ew", **pad
        )
        ttk.Button(main, text="浏览…", command=self._browse_excel).grid(
            row=1, column=3, sticky="w", **pad
        )

        ttk.Label(main, text="工作表").grid(row=2, column=0, sticky="e", **pad)
        self.sheet_combo = ttk.Combobox(
            main, textvariable=self.sheet_name, state="readonly", width=28
        )
        self.sheet_combo.grid(row=2, column=1, sticky="w", **pad)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_sheet_headers())

        ttk.Label(main, text="表头行").grid(row=2, column=2, sticky="e", **pad)
        ttk.Spinbox(main, from_=1, to=20, textvariable=self.header_row, width=6).grid(
            row=2, column=3, sticky="w", **pad
        )

        ttk.Label(main, text="待匹配列").grid(row=3, column=0, sticky="e", **pad)
        self.match_combo = ttk.Combobox(main, textvariable=self.match_col, width=28)
        self.match_combo.grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(main, text="输出标准名称列").grid(row=4, column=0, sticky="e", **pad)
        self.name_out_combo = ttk.Combobox(main, textvariable=self.name_out_col, width=28)
        self.name_out_combo.grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(main, text="输出项目编码列").grid(row=5, column=0, sticky="e", **pad)
        self.code_out_combo = ttk.Combobox(main, textvariable=self.code_out_col, width=28)
        self.code_out_combo.grid(row=5, column=1, sticky="w", **pad)

        self.summary_chk = ttk.Checkbutton(
            main,
            text="处理摘要列：若摘要以项目名称开头则删去该前缀，删完后若以「-」开头再删掉",
            variable=self.process_summary,
            command=self._toggle_summary,
        )
        self.summary_chk.grid(row=6, column=0, columnspan=4, sticky="w", **pad)

        ttk.Label(main, text="摘要列").grid(row=7, column=0, sticky="e", **pad)
        self.summary_combo = ttk.Combobox(main, textvariable=self.summary_col, width=28)
        self.summary_combo.grid(row=7, column=1, sticky="w", **pad)

        ttk.Label(main, text="摘要输出列").grid(row=8, column=0, sticky="e", **pad)
        self.summary_out_combo = ttk.Combobox(
            main, textvariable=self.summary_out_col, width=28
        )
        self.summary_out_combo.grid(row=8, column=1, sticky="w", **pad)

        opts = ttk.Frame(main)
        opts.grid(row=9, column=0, columnspan=4, sticky="w", **pad)
        ttk.Checkbutton(
            opts,
            text="未匹配时弹窗补录（含此前保留原名的项）",
            variable=self.prompt_unmatched,
        ).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(
            opts, text="只填空白单元格（不覆盖已有值）", variable=self.only_empty
        ).pack(side="left")

        out_row = ttk.Frame(main)
        out_row.grid(row=10, column=0, columnspan=4, sticky="ew", **pad)
        ttk.Radiobutton(
            out_row,
            text="覆盖原文件",
            variable=self.save_mode,
            value="overwrite",
            command=self._toggle_output,
        ).pack(side="left")
        ttk.Radiobutton(
            out_row,
            text="另存为",
            variable=self.save_mode,
            value="saveas",
            command=self._toggle_output,
        ).pack(side="left", padx=(12, 8))
        self.out_entry = ttk.Entry(out_row, textvariable=self.output_path)
        self.out_entry.pack(side="left", fill="x", expand=True)
        self.out_btn = ttk.Button(out_row, text="选择…", command=self._browse_out)
        self.out_btn.pack(side="left", padx=(8, 0))

        ttk.Label(main, textvariable=self.db_label, foreground="#555").grid(
            row=11, column=0, columnspan=4, sticky="w", **pad
        )

        btns = ttk.Frame(main)
        btns.grid(row=12, column=0, columnspan=4, pady=8)
        self.run_btn = ttk.Button(btns, text="开始填写", command=self._run)
        self.run_btn.pack(side="left", padx=8)
        ttk.Button(btns, text="重新加载文件", command=self._reload_meta).pack(
            side="left", padx=8
        )
        ttk.Button(btns, text="退出", command=self.root.destroy).pack(side="left", padx=8)

        ttk.Label(main, text="日志").grid(row=13, column=0, sticky="nw", **pad)
        log_frame = ttk.Frame(main)
        log_frame.grid(row=14, column=0, columnspan=4, sticky="nsew", **pad)
        self.log = tk.Text(log_frame, height=14, wrap="word", font=("Menlo", 11))
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        main.columnconfigure(1, weight=1)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(14, weight=1)

        self._append_log(
            "请选择 Excel。下拉框可改列名；若输出列不存在，填写时会在表头右侧新增一列。"
        )

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def _toggle_summary(self) -> None:
        state = "normal" if self.process_summary.get() else "disabled"
        self.summary_combo.configure(state=state)
        self.summary_out_combo.configure(state=state)

    def _toggle_output(self) -> None:
        state = "disabled" if self.save_mode.get() == "overwrite" else "normal"
        self.out_entry.configure(state=state)
        self.out_btn.configure(state=state)

    def _browse_excel(self) -> None:
        initial = DEFAULT_EXCEL_DIR if DEFAULT_EXCEL_DIR.is_dir() else Path.home()
        current = self.excel_path.get().strip()
        if current:
            p = Path(current)
            if p.parent.is_dir():
                initial = p.parent
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择 Excel",
            initialdir=str(initial),
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.excel_path.set(path)
            self._load_meta()

    def _browse_out(self) -> None:
        src = self.excel_path.get().strip()
        initial = Path(src).parent if src else DEFAULT_EXCEL_DIR
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="另存为",
            initialdir=str(initial) if Path(initial).is_dir() else str(Path.home()),
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.output_path.set(path)
            self.save_mode.set("saveas")
            self._toggle_output()

    def _reload_meta(self) -> None:
        if not self.excel_path.get().strip():
            messagebox.showinfo("提示", "请先选择 Excel 文件", parent=self.root)
            return
        self._load_meta()

    def _load_meta(self) -> None:
        path = Path(self.excel_path.get().strip())
        if not path.is_file():
            messagebox.showerror("错误", f"文件不存在:\n{path}", parent=self.root)
            return
        try:
            header_row = int(self.header_row.get() or 1)
        except ValueError:
            header_row = 1
            self.header_row.set("1")
        try:
            sheets, headers = peek_excel_sheets_and_headers(path, header_row=header_row)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc), parent=self.root)
            return
        self._sheets = sheets
        self._headers = headers
        self.sheet_combo["values"] = sheets
        if sheets and self.sheet_name.get() not in sheets:
            self.sheet_name.set(sheets[0])
        self._apply_sheet_headers()
        self._append_log(f"已加载 {path.name}，工作表: {', '.join(sheets)}")

    def _apply_sheet_headers(self) -> None:
        sheet = self.sheet_name.get() or (self._sheets[0] if self._sheets else "")
        cols = [c for c in self._headers.get(sheet, []) if c]
        for combo in (
            self.match_combo,
            self.name_out_combo,
            self.code_out_combo,
            self.summary_combo,
            self.summary_out_combo,
        ):
            combo["values"] = cols
        if cols:
            self.match_col.set(_pick(cols, NAME_CANDIDATES))
            self.name_out_col.set(_pick(cols, NAME_OUT_CANDIDATES) or "项目标准名称")
            self.code_out_col.set(_pick(cols, CODE_OUT_CANDIDATES) or "项目编码")
            self.summary_col.set(_pick(cols, SUMMARY_CANDIDATES) or "摘要")
            self.summary_out_col.set(_pick(cols, SUMMARY_OUT_CANDIDATES) or "摘要处理")
            if "摘要" in cols and "摘要处理" in cols:
                self.process_summary.set(True)
            self._toggle_summary()

    def _run(self) -> None:
        if self._running:
            return
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showinfo("提示", "请先选择 Excel 文件", parent=self.root)
            return
        if not Path(path).is_file():
            messagebox.showerror("错误", f"文件不存在:\n{path}", parent=self.root)
            return
        match_col = self.match_col.get().strip()
        name_out = self.name_out_col.get().strip()
        code_out = self.code_out_col.get().strip()
        if not match_col or not name_out or not code_out:
            messagebox.showinfo("提示", "请填写待匹配列、标准名称列、项目编码列", parent=self.root)
            return
        try:
            header_row = int(self.header_row.get() or 1)
        except ValueError:
            messagebox.showerror("错误", "表头行必须是数字", parent=self.root)
            return

        output_path = None
        if self.save_mode.get() != "overwrite":
            output_path = self.output_path.get().strip()
            if not output_path:
                messagebox.showinfo("提示", "请选择另存路径，或改选覆盖原文件", parent=self.root)
                return

        self._running = True
        self.run_btn.configure(state="disabled")
        self._append_log("--------")
        self._append_log("开始填写…")
        try:
            stats = fill_excel_project_columns(
                path,
                sheet_name=self.sheet_name.get().strip() or None,
                match_col=match_col,
                name_out_col=name_out,
                code_out_col=code_out,
                summary_col=self.summary_col.get().strip() or None,
                summary_out_col=self.summary_out_col.get().strip() or None,
                process_summary=self.process_summary.get(),
                header_row=header_row,
                matcher=self.matcher,
                prompt=self.prompt_unmatched.get(),
                interactive=True,
                parent=self.root,
                only_empty=self.only_empty.get(),
                output_path=output_path,
                log=self._append_log,
            )
            self.matcher.save()
            self.db_label.set(
                f"匹配库: {self.matcher.db_path.name}　标准项目 {self.matcher.project_count} 个"
            )
            extra = ""
            if stats.get("unmatched_names"):
                extra = "\n未匹配: " + "、".join(stats["unmatched_names"][:15])
                if len(stats["unmatched_names"]) > 15:
                    extra += f" 等 {len(stats['unmatched_names'])} 个"
            messagebox.showinfo(
                "完成",
                f"已保存到:\n{stats['output_path']}\n\n"
                f"匹配 {stats['matched']} 行，未匹配 {stats['unmatched']} 行"
                + (
                    f"\n摘要去掉项目名前缀 {stats['summary_stripped']} 行"
                    if self.process_summary.get()
                    else ""
                )
                + extra,
                parent=self.root,
            )
        except Exception as exc:
            self._append_log(traceback.format_exc())
            messagebox.showerror("填写失败", str(exc), parent=self.root)
        finally:
            self._running = False
            self.run_btn.configure(state="normal")


def main(argv=None) -> int:
    _ = argv
    root = tk.Tk()
    MatcherExcelApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
