# -*- coding: utf-8 -*-
"""财务付款一体化 · 图形窗口。

无参数运行 ``财务付款一体化.py`` 时进入本窗口。
命令行子命令与 ``--menu`` 仍走原来的 CLI。
"""

from __future__ import annotations

import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from typing import Any, Callable, List

from .config import Paths


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
    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, s: str) -> int:
        if s:
            self.q.put(s)
        return len(s) if s else 0

    def flush(self) -> None:
        return None


class PaymentHubApp:
    def __init__(self, root: Any):
        tk, filedialog, messagebox, ttk = _tk()
        self.tk = tk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.ttk = ttk
        self.root = root
        self.root.title("财务付款一体化")
        self.root.geometry("920x720")
        self.root.minsize(780, 600)

        self.paths = Paths()
        self._running = False
        self._log_q: queue.Queue = queue.Queue()
        self.prompt_project = tk.BooleanVar(value=False)
        self.ledger_dry_run = tk.BooleanVar(value=False)
        self.only_confirmed = tk.BooleanVar(value=False)
        self.ledger_path = tk.StringVar(value=str(self.paths.ledger_xlsx))
        self.status = tk.StringVar(value="就绪")

        self._build()
        self.root.after(120, self._drain_log)

    def _build(self) -> None:
        ttk = self.ttk
        tk = self.tk
        pad = {"padx": 8, "pady": 4}
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(5, weight=1)

        ttk.Label(
            main,
            text="日常：导出学习池 → 在 Excel 里改科目或点确认 → 应用确认。月结点「月结全流程」。",
            wraplength=880,
        ).grid(row=0, column=0, sticky="w", **pad)

        path_line = (
            f"学习池  {self.paths.learn_pool_root}\n"
            f"A_raw   {self.paths.raw_dir}\n"
            f"B_std   {self.paths.std_dir}    C_sum  {self.paths.sum_xlsx.parent}"
        )
        ttk.Label(main, text=path_line, foreground="#444", wraplength=880).grid(
            row=1, column=0, sticky="w", **pad
        )

        learn = ttk.LabelFrame(main, text="确认学习", padding=8)
        learn.grid(row=2, column=0, sticky="ew", **pad)
        self._btns(learn, [
            ("导出学习池", self._cmd("pool")),
            ("应用确认", self._apply),
            ("账簿写入记忆", self._learn_ledger),
            ("重刷分类", self._cmd("reclassify")),
            ("健康度", self._cmd("health")),
            ("待复核池", self._cmd("review")),
            ("导出分类总表", self._cmd("export-class")),
            ("应用分类总表", self._cmd("apply-class")),
            ("压缩记忆", self._cmd("compact-memory")),
        ], cols=3)

        month = ttk.LabelFrame(main, text="月结出账", padding=8)
        month.grid(row=3, column=0, sticky="ew", **pad)
        ttk.Button(month, text="月结全流程", command=self._run_month).grid(
            row=0, column=0, padx=4, pady=4, sticky="w"
        )
        ttk.Checkbutton(
            month,
            text="弹窗补录未匹配项目名（会卡住窗口直到补完）",
            variable=self.prompt_project,
        ).grid(row=0, column=1, columnspan=2, padx=8, sticky="w")

        maint = ttk.LabelFrame(main, text="维护 / 分步", padding=8)
        maint.grid(row=4, column=0, sticky="ew", **pad)
        self._btns(maint, [
            ("raw 标准化", self._cmd("raw-std", project=True)),
            ("分类科目", self._cmd("classify")),
            ("汇总 Excel", self._cmd("summary")),
            ("拆分 B_std", self._cmd("std", project=True)),
            ("现金流 HTML", self._cmd("cashflow")),
            ("不平明细", self._cmd("amount-mismatch")),
            ("导出科目体系", self._cmd("export-taxonomy")),
            ("导入科目体系", self._import_taxonomy),
            ("B_std 加下拉", self._cmd("dropdowns")),
            ("查看路径", self._cmd("show-paths")),
            ("打开学习池文件夹", self._open_learn_pool),
            ("打开汇总文件夹", self._open_sum),
        ], cols=4)

        opt = ttk.Frame(main)
        opt.grid(row=5, column=0, sticky="nsew", **pad)
        opt.columnconfigure(1, weight=1)
        opt.rowconfigure(2, weight=1)

        ttk.Label(opt, text="账簿 Excel").grid(row=0, column=0, sticky="e", padx=4)
        ttk.Entry(opt, textvariable=self.ledger_path).grid(
            row=0, column=1, sticky="ew", padx=4
        )
        ttk.Button(opt, text="浏览…", command=self._browse_ledger).grid(
            row=0, column=2, sticky="w", padx=4
        )
        ttk.Checkbutton(opt, text="账簿仅预览，不写记忆", variable=self.ledger_dry_run).grid(
            row=1, column=1, sticky="w", padx=4
        )
        ttk.Checkbutton(
            opt, text="apply 仅学习「确认」行", variable=self.only_confirmed
        ).grid(row=1, column=2, sticky="w", padx=4)

        log_fr = ttk.LabelFrame(opt, text="运行日志", padding=4)
        log_fr.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        log_fr.columnconfigure(0, weight=1)
        log_fr.rowconfigure(0, weight=1)
        self.log = tk.Text(log_fr, height=16, wrap="word", font=("Menlo", 11))
        scroll = ttk.Scrollbar(log_fr, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        bar = ttk.Frame(main)
        bar.grid(row=6, column=0, sticky="ew")
        ttk.Label(bar, textvariable=self.status).pack(side="left")
        ttk.Button(bar, text="清空日志", command=self._clear_log).pack(side="right")

    def _btns(self, parent, items: list, cols: int = 3) -> None:
        ttk = self.ttk
        for i, (label, cmd) in enumerate(items):
            ttk.Button(parent, text=label, command=cmd).grid(
                row=i // cols, column=i % cols, padx=4, pady=4, sticky="ew"
            )
            parent.columnconfigure(i % cols, weight=1)

    def _cmd(self, name: str, *, project: bool = False) -> Callable[[], None]:
        def _go() -> None:
            argv = [name]
            if project and self.prompt_project.get():
                argv.append("--prompt-project")
            self._run_argv(argv)
        return _go

    def _apply(self) -> None:
        argv = ["apply"]
        if self.only_confirmed.get():
            argv.append("--only-confirmed")
        self._run_argv(argv)

    def _run_month(self) -> None:
        argv = ["run"]
        if self.prompt_project.get():
            argv.append("--prompt-project")
        self._run_argv(argv)

    def _learn_ledger(self) -> None:
        path = self.ledger_path.get().strip()
        if not path:
            self.messagebox.showwarning("账簿", "请先选择账簿 Excel。")
            return
        argv = ["learn-ledger", "--ledger", path]
        if self.ledger_dry_run.get():
            argv.append("--dry-run")
        self._run_argv(argv)

    def _import_taxonomy(self) -> None:
        path = self.filedialog.askopenfilename(
            title="选择科目体系 Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")],
        )
        if not path:
            return
        self._run_argv(["import-taxonomy", "-i", path])

    def _browse_ledger(self) -> None:
        initial = Path(self.ledger_path.get() or str(self.paths.ledger_xlsx))
        path = self.filedialog.askopenfilename(
            title="选择账簿 Excel",
            initialdir=str(initial.parent if initial.parent.exists() else Path.home()),
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")],
        )
        if path:
            self.ledger_path.set(path)

    def _open_learn_pool(self) -> None:
        self._open_dir(self.paths.learn_pool_root)

    def _open_sum(self) -> None:
        self._open_dir(self.paths.sum_xlsx.parent)

    def _open_dir(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.Popen(["open", str(path)])
        except Exception as e:
            self.messagebox.showerror("打开文件夹", str(e))

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

    def _run_argv(self, argv: List[str]) -> None:
        if self._running:
            self.messagebox.showinfo("进行中", "请等待当前任务结束。")
            return
        if not self.messagebox.askokcancel("确认执行", "即将执行：\n  " + " ".join(argv)):
            return
        self._running = True
        self.status.set("运行中: " + " ".join(argv))
        self._append_log("\n" + "=" * 56 + f"\n▶ {' '.join(argv)}\n" + "=" * 56 + "\n")

        use_main_thread = "--prompt-project" in argv

        def job() -> None:
            from .cli import main as cli_main

            writer = _QueueWriter(self._log_q)
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = writer  # type: ignore[assignment]
            code = 1
            try:
                code = int(cli_main(argv) or 0)
            except SystemExit as e:
                code = int(e.code) if e.code not in (None, "") else 0
            except Exception:
                writer.write(traceback.format_exc())
                code = 1
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                self.root.after(0, lambda c=code, a=argv: self._job_done(c, a))

        if use_main_thread:
            job()
        else:
            threading.Thread(target=job, daemon=True).start()

    def _job_done(self, code: int, argv: List[str]) -> None:
        self._running = False
        ok = code == 0
        self.status.set(("完成" if ok else "失败") + f"（退出码 {code}）: " + " ".join(argv))
        self._append_log(f"\n{'✅' if ok else '⚠️'} 结束（退出码 {code}）\n")
        if ok:
            self.messagebox.showinfo("完成", "已完成：\n  " + " ".join(argv))
        else:
            self.messagebox.showerror("失败", f"退出码 {code}\n请查看下方日志。")


def launch_gui() -> int:
    tk, _filedialog, _messagebox, _ttk = _tk()
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.2)
    except tk.TclError:
        pass
    PaymentHubApp(root)
    root.mainloop()
    return 0
