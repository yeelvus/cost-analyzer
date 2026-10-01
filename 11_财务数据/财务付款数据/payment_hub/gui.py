# -*- coding: utf-8 -*-
"""财务付款月结窗口。

无参数运行 ``财务付款一体化.py`` 时进入本窗口。
科目对应关系在 ``摘要科目分类.py``，这里只跑流程，并按那份对应关系填入科目。
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
    from tkinter import messagebox, ttk

    return tk, messagebox, ttk


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
        tk, messagebox, ttk = _tk()
        self.tk = tk
        self.messagebox = messagebox
        self.ttk = ttk
        self.root = root
        self.root.title("财务付款")
        self.root.geometry("720x480")
        self.root.minsize(620, 400)

        self.paths = Paths()
        self._running = False
        self._log_q: queue.Queue = queue.Queue()
        self.prompt_project = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="就绪")
        self._more_open = tk.BooleanVar(value=False)

        self._build()
        self.root.after(120, self._drain_log)

    def _build(self) -> None:
        ttk = self.ttk
        tk = self.tk
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        ttk.Label(main, text="财务付款月结", font=("", 16, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            main,
            text="科目和摘要的对应关系在「摘要科目分类」里单独处理。这里只跑月结，并按那份对应关系填入科目。",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))

        steps = ttk.Frame(main)
        steps.grid(row=2, column=0, sticky="ew", pady=4)
        steps.columnconfigure(1, weight=1)
        ttk.Button(steps, text="月结全流程", command=self._run_month, width=16).grid(
            row=0, column=0, sticky="ew", padx=(0, 12), pady=4
        )
        ttk.Label(steps, text="原始表 → 摘要识别 → 填入科目 → 汇总和现金流").grid(
            row=0, column=1, sticky="w", pady=4
        )

        links = ttk.Frame(main)
        links.grid(row=3, column=0, sticky="w", pady=4)
        ttk.Label(links, text="打开").pack(side="left", padx=(0, 8))
        ttk.Button(links, text="汇总文件夹", command=self._open_sum).pack(side="left", padx=4)
        ttk.Checkbutton(
            links,
            text="月结时补录未匹配项目名",
            variable=self.prompt_project,
        ).pack(side="left", padx=(16, 0))

        ttk.Checkbutton(
            main,
            text="更多",
            variable=self._more_open,
            command=self._toggle_more,
        ).grid(row=4, column=0, sticky="w", pady=4)

        self.more = ttk.LabelFrame(main, text="不重跑整月时", padding=8)
        self.more.columnconfigure(1, weight=1)
        ttk.Button(self.more, text="只填入科目", command=self._cmd("reclassify"), width=16).grid(
            row=0, column=0, sticky="w", padx=4, pady=4
        )
        ttk.Label(self.more, text="按当前科目对应重写分类结果和项目明细，不改科目记忆").grid(
            row=0, column=1, sticky="w", padx=4
        )

        opt = ttk.Frame(main)
        opt.grid(row=6, column=0, sticky="nsew", pady=4)
        opt.columnconfigure(0, weight=1)
        opt.rowconfigure(0, weight=1)
        main.rowconfigure(6, weight=1)

        log_fr = ttk.LabelFrame(opt, text="运行日志", padding=4)
        log_fr.grid(row=0, column=0, sticky="nsew")
        log_fr.columnconfigure(0, weight=1)
        log_fr.rowconfigure(0, weight=1)
        self.log = tk.Text(log_fr, height=12, wrap="word", font=("Menlo", 11))
        scroll = ttk.Scrollbar(log_fr, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        bar = ttk.Frame(main)
        bar.grid(row=7, column=0, sticky="ew")
        ttk.Label(bar, textvariable=self.status).pack(side="left")
        ttk.Button(bar, text="清空日志", command=self._clear_log).pack(side="right")

    def _toggle_more(self) -> None:
        if self._more_open.get():
            self.more.grid(row=5, column=0, sticky="ew", pady=4)
        else:
            self.more.grid_remove()

    def _cmd(self, name: str) -> Callable[[], None]:
        def _go() -> None:
            self._run_argv([name])
        return _go

    def _run_month(self) -> None:
        argv = ["run"]
        if self.prompt_project.get():
            argv.append("--prompt-project")
        self._run_argv(argv)

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
    tk, _messagebox, _ttk = _tk()
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.2)
    except tk.TclError:
        pass
    PaymentHubApp(root)
    root.mainloop()
    return 0
