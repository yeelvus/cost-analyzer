# -*- coding: utf-8 -*-
"""摘要科目分类的命令行和窗口。

填科目只另存工作簿。写入记忆只改 learn_classify/data。
两者都不改月结目录。
"""

from __future__ import annotations

import argparse
import queue
import sys
import threading
import traceback
from pathlib import Path
from typing import Any, Callable, List, Optional


def _data_dir(explicit: Optional[str] = None) -> Path:
    if explicit:
        return Path(explicit).expanduser()
    from subject_fill import default_data_dir

    return default_data_dir()


def _cmd_fill(args: argparse.Namespace) -> int:
    from subject_fill import fill_workbook

    dest = fill_workbook(
        Path(args.input),
        Path(args.output) if args.output else None,
        sheet=args.sheet or None,
        header_row=args.header_row,
        summary_col=args.summary,
        project_col=args.project or "",
        subject_col=args.subject or "",
        vendor_col=args.vendor if args.vendor is not None else None,
        business_col=args.business,
        only_paid=bool(args.only_paid),
        inplace=bool(args.inplace),
        data_dir=_data_dir(args.data_dir),
    )
    print(f"✅ fill: {dest}")
    return 0


def _cmd_learn(args: argparse.Namespace) -> int:
    from ledger_book_classify import learn_from_ledger

    path = Path(args.input).expanduser()
    data_dir = _data_dir(args.data_dir)
    print(f"账簿: {path}")
    print(f"学习库: {data_dir}")
    if args.dry_run:
        print("模式: 仅预览，不写 memory / taxonomy")
    stats = learn_from_ledger(
        path,
        data_dir,
        sheet_name=args.sheet or None,
        auto_taxonomy=not bool(args.no_auto_taxonomy),
        dry_run=bool(args.dry_run),
        header_row=args.header_row,
        summary_col=args.summary,
        l1_col=args.l1,
        l2_col=args.l2,
        l3_col=args.l3,
        project_col=args.project,
        vendor_col=args.vendor,
    )
    print(f"✅ learn: {stats}")
    return 0


def _paths():
    from payment_hub.config import Paths

    return Paths()


def _cmd_pool(_args: argparse.Namespace) -> int:
    from payment_hub.learn_pool import export_learn_pool

    paths = _paths()
    stats = export_learn_pool(paths)
    print(f"✅ pool: {stats}")
    print(f"   打开: {paths.learn_pool_priority_xlsx}")
    print("   改完后: python 摘要科目分类.py apply")
    return 0


def _cmd_apply(args: argparse.Namespace) -> int:
    from payment_hub.learn_pool import apply_learn_pool

    stats = apply_learn_pool(
        _paths(),
        reclassify=False,
        auto_taxonomy=not bool(args.no_auto_taxonomy),
        learn_unconfirmed_edits=not bool(args.only_confirmed),
    )
    print(f"✅ apply: {stats}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="摘要科目分类",
        description="把摘要列对应成一级/二级/三级科目。不改月结文件。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "无参数打开窗口。\n"
            "填进月结结果请另运行: python 财务付款一体化.py run"
        ),
    )
    sub = parser.add_subparsers(dest="cmd")

    fill = sub.add_parser("fill", help="摘要列 → 一二三级科目，另存工作簿")
    fill.add_argument("-i", "--input", required=True, help="含摘要列的 Excel")
    fill.add_argument("-o", "--output", default=None, help="默认写到「原名_科目.xlsx」")
    fill.add_argument("--sheet", default="", help="工作表，默认「明细」或第一张")
    fill.add_argument("--header-row", type=int, default=1, help="表头行号，从 1 起")
    fill.add_argument("--summary", default="摘要", help="摘要列名")
    fill.add_argument("--subject", default="科目", help="科目列名，没有就留空")
    fill.add_argument("--vendor", default=None, help="供应商列名，默认自动识别")
    fill.add_argument("--project", default="项目名称", help="项目名称列，用于去掉摘要前缀")
    fill.add_argument("--business", default="业务摘要", help="已识别的业务摘要列")
    fill.add_argument("--only-paid", action="store_true", help="只分类实付金额非空的行")
    fill.add_argument("--inplace", action="store_true", help="覆盖原文件（默认不覆盖）")
    fill.add_argument("--data-dir", default=None, help="科目记忆目录，默认 learn_classify/data")
    fill.set_defaults(func=_cmd_fill)

    learn = sub.add_parser("learn", help="已填的一二三级科目 → 写入记忆，不改原文件")
    learn.add_argument("-i", "--input", required=True)
    learn.add_argument("--sheet", default="")
    learn.add_argument("--header-row", type=int, default=1)
    learn.add_argument("--summary", default=None, help="摘要列，默认自动")
    learn.add_argument("--l1", default=None, help="一级科目列")
    learn.add_argument("--l2", default=None, help="二级科目列")
    learn.add_argument("--l3", default=None, help="三级科目列，可空")
    learn.add_argument("--project", default=None)
    learn.add_argument("--vendor", default=None)
    learn.add_argument("--dry-run", action="store_true")
    learn.add_argument("--no-auto-taxonomy", action="store_true")
    learn.add_argument("--data-dir", default=None)
    learn.set_defaults(func=_cmd_learn)

    pool = sub.add_parser("pool", help="从月结明细导出学习池（不改 A2）")
    pool.set_defaults(func=_cmd_pool)

    apply = sub.add_parser("apply", help="应用学习池确认，只写记忆")
    apply.add_argument("--only-confirmed", action="store_true")
    apply.add_argument("--no-auto-taxonomy", action="store_true")
    apply.set_defaults(func=_cmd_apply)

    gui = sub.add_parser("gui", help="打开图形窗口")
    gui.set_defaults(func=lambda _args: launch_gui())
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "cmd", None):
        try:
            if gui_available():
                return launch_gui()
        except Exception as exc:
            print(f"图形窗口未能启动（{exc}）", file=sys.stderr)
        parser.print_help()
        return 0
    try:
        return int(args.func(args) or 0)
    except Exception as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 1


def gui_available() -> bool:
    try:
        import tkinter  # noqa: F401
        import tkinter.ttk  # noqa: F401

        return True
    except Exception:
        return False


def _tk():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    return tk, filedialog, messagebox, ttk


class _QueueWriter:
    def __init__(self, log_q: queue.Queue):
        self.q = log_q

    def write(self, text: str) -> int:
        if text:
            self.q.put(text)
        return len(text) if text else 0

    def flush(self) -> None:
        return None


class SubjectApp:
    def __init__(self, root: Any):
        tk, filedialog, messagebox, ttk = _tk()
        self.tk = tk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.ttk = ttk
        self.root = root
        self.root.title("摘要科目分类")
        self.root.geometry("780x640")
        self.root.minsize(680, 520)

        self._running = False
        self._log_q: queue.Queue = queue.Queue()
        self.path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.header_row = tk.StringVar(value="1")
        self.summary_col = tk.StringVar()
        self.subject_col = tk.StringVar()
        self.vendor_col = tk.StringVar()
        self.project_col = tk.StringVar()
        self.l1_col = tk.StringVar()
        self.l2_col = tk.StringVar()
        self.l3_col = tk.StringVar()
        self.dry_run = tk.BooleanVar(value=False)
        self.only_confirmed = tk.BooleanVar(value=False)
        self._more_open = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="就绪")
        self._sheets: list[str] = []
        self._headers: dict[str, list[str]] = {}

        self._build()
        self.root.after(120, self._drain_log)

    def _build(self) -> None:
        ttk = self.ttk
        tk = self.tk
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        ttk.Label(main, text="摘要对应一二三级科目", font=("", 16, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            main,
            text="只处理这一份表。写出的新文件不进月结目录，写入记忆也不改月结结果。",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))

        form = ttk.Frame(main)
        form.grid(row=2, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)
        ttk.Label(form, text="Excel").grid(row=0, column=0, sticky="e", padx=4, pady=3)
        ttk.Entry(form, textvariable=self.path).grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=4, pady=3
        )
        ttk.Button(form, text="浏览…", command=self._browse).grid(
            row=0, column=4, sticky="w", padx=4, pady=3
        )
        ttk.Label(form, text="工作表").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.sheet_combo = ttk.Combobox(
            form, textvariable=self.sheet_name, state="readonly", width=22
        )
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=3)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_headers())
        ttk.Label(form, text="表头行").grid(row=1, column=2, sticky="e", padx=4, pady=3)
        ttk.Spinbox(form, from_=1, to=20, textvariable=self.header_row, width=6).grid(
            row=1, column=3, sticky="w", padx=4, pady=3
        )
        ttk.Button(form, text="读取表头", command=self._reload_headers).grid(
            row=1, column=4, sticky="w", padx=4, pady=3
        )

        ttk.Label(form, text="摘要列").grid(row=2, column=0, sticky="e", padx=4, pady=3)
        self.summary_combo = ttk.Combobox(form, textvariable=self.summary_col, width=22)
        self.summary_combo.grid(row=2, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(form, text="科目列").grid(row=2, column=2, sticky="e", padx=4, pady=3)
        self.subject_combo = ttk.Combobox(form, textvariable=self.subject_col, width=22)
        self.subject_combo.grid(row=2, column=3, sticky="ew", padx=4, pady=3)

        ttk.Label(form, text="供应商").grid(row=3, column=0, sticky="e", padx=4, pady=3)
        self.vendor_combo = ttk.Combobox(form, textvariable=self.vendor_col, width=22)
        self.vendor_combo.grid(row=3, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(form, text="项目名称").grid(row=3, column=2, sticky="e", padx=4, pady=3)
        self.project_combo = ttk.Combobox(form, textvariable=self.project_col, width=22)
        self.project_combo.grid(row=3, column=3, sticky="ew", padx=4, pady=3)

        actions = ttk.Frame(main)
        actions.grid(row=3, column=0, sticky="ew", pady=(10, 4))
        actions.columnconfigure(1, weight=1)
        ttk.Button(actions, text="填入科目", command=self._fill, width=16).grid(
            row=0, column=0, sticky="w", padx=(0, 12)
        )
        ttk.Label(actions, text="另存为「原名_科目.xlsx」，原文件不动").grid(
            row=0, column=1, sticky="w"
        )

        learn = ttk.LabelFrame(main, text="用表里已经填好的科目校正对应关系", padding=8)
        learn.grid(row=4, column=0, sticky="ew", pady=(8, 4))
        learn.columnconfigure(1, weight=1)
        learn.columnconfigure(3, weight=1)
        ttk.Label(learn, text="一级科目").grid(row=0, column=0, sticky="e", padx=4, pady=3)
        self.l1_combo = ttk.Combobox(learn, textvariable=self.l1_col, width=22)
        self.l1_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(learn, text="二级科目").grid(row=0, column=2, sticky="e", padx=4, pady=3)
        self.l2_combo = ttk.Combobox(learn, textvariable=self.l2_col, width=22)
        self.l2_combo.grid(row=0, column=3, sticky="ew", padx=4, pady=3)
        ttk.Label(learn, text="三级科目").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.l3_combo = ttk.Combobox(learn, textvariable=self.l3_col, width=22)
        self.l3_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=3)
        ttk.Checkbutton(learn, text="仅预览，不写记忆", variable=self.dry_run).grid(
            row=1, column=2, columnspan=2, sticky="w", padx=4
        )
        ttk.Button(learn, text="写入记忆", command=self._learn, width=16).grid(
            row=2, column=0, sticky="w", padx=4, pady=(6, 2)
        )
        ttk.Label(learn, text="只更新科目记忆，不改这份 Excel，也不改月结文件").grid(
            row=2, column=1, columnspan=3, sticky="w", padx=4
        )

        ttk.Checkbutton(
            main, text="更多", variable=self._more_open, command=self._toggle_more
        ).grid(row=5, column=0, sticky="w", pady=(6, 2))

        self.more = ttk.LabelFrame(main, text="从月结明细整理学习池", padding=8)
        self.more.columnconfigure(1, weight=1)
        ttk.Button(self.more, text="导出学习池", command=self._pool, width=16).grid(
            row=0, column=0, sticky="w", padx=4, pady=3
        )
        ttk.Label(self.more, text="读月结明细，生成待确认的表").grid(
            row=0, column=1, sticky="w", padx=4
        )
        ttk.Button(self.more, text="应用确认", command=self._apply, width=16).grid(
            row=1, column=0, sticky="w", padx=4, pady=3
        )
        ttk.Label(self.more, text="把确认过的科目写入记忆，不回写月结文件").grid(
            row=1, column=1, sticky="w", padx=4
        )
        ttk.Checkbutton(
            self.more,
            text="只学习点了「确认」的行",
            variable=self.only_confirmed,
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=2)

        log_fr = ttk.LabelFrame(main, text="运行日志", padding=4)
        log_fr.grid(row=7, column=0, sticky="nsew", pady=(8, 4))
        log_fr.columnconfigure(0, weight=1)
        log_fr.rowconfigure(0, weight=1)
        main.rowconfigure(7, weight=1)
        self.log = tk.Text(log_fr, height=10, wrap="word", font=("Menlo", 11))
        scroll = ttk.Scrollbar(log_fr, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        bar = ttk.Frame(main)
        bar.grid(row=8, column=0, sticky="ew")
        ttk.Label(bar, textvariable=self.status).pack(side="left")
        ttk.Button(bar, text="清空日志", command=self._clear_log).pack(side="right")

    def _toggle_more(self) -> None:
        if self._more_open.get():
            self.more.grid(row=6, column=0, sticky="ew", pady=4)
        else:
            self.more.grid_remove()

    def _browse(self) -> None:
        current = Path(self.path.get().strip() or str(Path.home()))
        initial = current.parent if current.exists() else Path.home()
        picked = self.filedialog.askopenfilename(
            title="选择含摘要列的 Excel",
            initialdir=str(initial if initial.exists() else Path.home()),
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")],
        )
        if picked:
            self.path.set(picked)
            self._reload_headers()

    def _header_row_num(self) -> int:
        try:
            return int(str(self.header_row.get() or "1").strip() or "1")
        except ValueError:
            self.header_row.set("1")
            return 1

    def _reload_headers(self, silent: bool = False) -> None:
        path = self.path.get().strip()
        if not path or not Path(path).is_file():
            if not silent:
                self.messagebox.showwarning("读取表头", "请先选择有效的 Excel 文件。")
            return
        try:
            from ledger_book_classify import peek_excel_sheets_and_headers

            sheets, headers = peek_excel_sheets_and_headers(
                Path(path), header_row=self._header_row_num()
            )
        except Exception as exc:
            self.messagebox.showerror("读取表头失败", str(exc))
            return
        self._sheets = sheets
        self._headers = headers
        self.sheet_combo["values"] = sheets
        preferred = self.sheet_name.get().strip()
        if preferred not in sheets:
            if "明细" in sheets:
                preferred = "明细"
            elif "汇总" in sheets:
                preferred = "汇总"
            else:
                preferred = sheets[0] if sheets else ""
        self.sheet_name.set(preferred)
        self._apply_headers()

    def _apply_headers(self) -> None:
        cols = list(self._headers.get(self.sheet_name.get().strip(), []))
        optional = [""] + cols
        self.summary_combo["values"] = cols
        for combo in (self.subject_combo, self.vendor_combo, self.project_combo, self.l3_combo):
            combo["values"] = optional
        self.l1_combo["values"] = cols
        self.l2_combo["values"] = cols
        self.summary_col.set(self._pick(cols, ("摘要", "业务摘要", "摘要处理")))
        self.subject_col.set(self._pick(cols, ("科目", "费用科目"), allow_empty=True))
        self.vendor_col.set(
            self._pick(cols, ("供应商名称", "供应商", "收款单位", "收款人"), allow_empty=True)
        )
        self.project_col.set(
            self._pick(cols, ("项目名称", "标准项目名称", "工程名称"), allow_empty=True)
        )
        self.l1_col.set(self._pick(cols, ("一级科目", "正确一级科目", "正确一级", "一级")))
        self.l2_col.set(self._pick(cols, ("二级科目", "正确二级科目", "正确二级", "二级")))
        self.l3_col.set(
            self._pick(cols, ("三级科目", "正确三级科目", "正确三级", "三级"), allow_empty=True)
        )

    @staticmethod
    def _pick(columns: list[str], candidates: tuple[str, ...], *, allow_empty: bool = False) -> str:
        for name in candidates:
            if name in columns:
                return name
        if allow_empty:
            return ""
        return columns[0] if columns else ""

    def _require_file(self) -> Optional[Path]:
        path = self.path.get().strip()
        if not path or not Path(path).is_file():
            self.messagebox.showwarning("摘要科目分类", "请先选择有效的 Excel 文件。")
            return None
        return Path(path)

    def _fill(self) -> None:
        path = self._require_file()
        if path is None:
            return
        summary = self.summary_col.get().strip()
        if not summary:
            self.messagebox.showwarning("填入科目", "请指定摘要列。")
            return
        from subject_fill import default_output_path

        dest = default_output_path(path)
        sheet = self.sheet_name.get().strip()
        subject = self.subject_col.get().strip()
        vendor = self.vendor_col.get().strip()
        project = self.project_col.get().strip()

        def job() -> None:
            from subject_fill import fill_workbook

            fill_workbook(
                path,
                dest,
                sheet=sheet or None,
                header_row=self._header_row_num(),
                summary_col=summary,
                project_col=project,
                subject_col=subject,
                vendor_col=vendor,
                only_paid=False,
            )

        self._start(f"填入科目 → {dest.name}", job)

    def _learn(self) -> None:
        path = self._require_file()
        if path is None:
            return
        summary = self.summary_col.get().strip()
        l1 = self.l1_col.get().strip()
        l2 = self.l2_col.get().strip()
        if not summary or not l1 or not l2:
            self.messagebox.showwarning("写入记忆", "请指定摘要列、一级科目列、二级科目列。")
            return
        sheet = self.sheet_name.get().strip()
        l3 = self.l3_col.get().strip()
        project = self.project_col.get().strip()
        vendor = self.vendor_col.get().strip()
        dry = bool(self.dry_run.get())

        def job() -> None:
            from ledger_book_classify import learn_from_ledger

            learn_from_ledger(
                path,
                _data_dir(),
                sheet_name=sheet or None,
                dry_run=dry,
                header_row=self._header_row_num(),
                summary_col=summary,
                l1_col=l1,
                l2_col=l2,
                l3_col=l3,
                project_col=project or None,
                vendor_col=vendor or None,
            )

        label = "预览记忆" if dry else "写入记忆"
        self._start(label, job)

    def _pool(self) -> None:
        def job() -> None:
            _cmd_pool(argparse.Namespace())

        self._start("导出学习池", job)

    def _apply(self) -> None:
        only = bool(self.only_confirmed.get())

        def job() -> None:
            _cmd_apply(argparse.Namespace(only_confirmed=only, no_auto_taxonomy=False))

        self._start("应用确认", job)

    def _start(self, label: str, fn: Callable[[], None]) -> None:
        if self._running:
            self.messagebox.showinfo("进行中", "请等待当前任务结束。")
            return
        if not self.messagebox.askokcancel("确认执行", label):
            return
        self._running = True
        self.status.set("运行中: " + label)
        self._append_log("\n" + "=" * 56 + f"\n▶ {label}\n" + "=" * 56 + "\n")

        def job() -> None:
            writer = _QueueWriter(self._log_q)
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = writer  # type: ignore[assignment]
            code = 0
            try:
                fn()
            except Exception:
                writer.write(traceback.format_exc())
                code = 1
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                self.root.after(0, lambda c=code, name=label: self._job_done(c, name))

        threading.Thread(target=job, daemon=True).start()

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def _drain_log(self) -> None:
        try:
            while True:
                self._append_log(self._log_q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(120, self._drain_log)

    def _job_done(self, code: int, label: str) -> None:
        self._running = False
        ok = code == 0
        self.status.set(("完成" if ok else "失败") + f"（退出码 {code}）: " + label)
        self._append_log(f"\n{'✅' if ok else '⚠️'} 结束（退出码 {code}）\n")
        if ok:
            self.messagebox.showinfo("完成", label)
        else:
            self.messagebox.showerror("失败", f"{label}\n请查看下方日志。")


def launch_gui() -> int:
    tk, _filedialog, _messagebox, _ttk = _tk()
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.2)
    except tk.TclError:
        pass
    SubjectApp(root)
    root.mainloop()
    return 0
