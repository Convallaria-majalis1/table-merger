# -*- coding: utf-8 -*-
"""
Excel / CSV 多表合并清洗器 —— 图形界面

给客户用的版本：双击就能打开，选文件、点开始、出结果。
技术栈：tkinter（Python 自带，不需要额外装界面库）+ pandas + openpyxl
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import core

APP_TITLE = "表格合并清洗器 v1.0"
SUPPORTED_TYPES = [("表格文件", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("所有文件", "*.*")]


def app_dir() -> Path:
    """程序所在目录：源码运行时是 main.py 所在目录，打包成 exe 后是 exe 所在目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.files: list[str] = []
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None

        root.title(APP_TITLE)
        root.geometry("760x620")
        root.minsize(680, 560)

        self._build_ui()
        self._poll_log()

    # ---------- 界面 ----------
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # 顶部说明
        header = ttk.Frame(self.root)
        header.pack(fill="x", **pad)
        ttk.Label(
            header,
            text="把多个 Excel / CSV 合并成一张干净的表",
            font=("Microsoft YaHei", 14, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="支持：多文件合并、一个文件多个 Sheet、自动清空格、去掉空行空列、按需去重",
            foreground="#666666",
        ).pack(anchor="w", pady=(4, 0))

        # 文件列表区
        file_frame = ttk.LabelFrame(self.root, text=" 1. 选择要合并的文件 ")
        file_frame.pack(fill="both", expand=False, **pad)

        list_wrap = ttk.Frame(file_frame)
        list_wrap.pack(fill="both", expand=True, padx=10, pady=8)

        self.listbox = tk.Listbox(list_wrap, height=6, selectmode=tk.EXTENDED, font=("Microsoft YaHei", 10))
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(btn_row, text="添加文件", command=self.add_files).pack(side="left")
        ttk.Button(btn_row, text="移除选中", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(btn_row, text="清空", command=self.clear_files).pack(side="left")
        self.count_label = ttk.Label(btn_row, text="已选 0 个文件", foreground="#666666")
        self.count_label.pack(side="right")

        # 选项区
        option_frame = ttk.LabelFrame(self.root, text=" 2. 处理选项 ")
        option_frame.pack(fill="x", **pad)

        self.var_all_sheets = tk.BooleanVar(value=False)
        self.var_add_source = tk.BooleanVar(value=True)
        self.var_dedupe = tk.BooleanVar(value=False)
        self.var_keep_empty = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            option_frame, text="读取每个文件的全部 Sheet（否则只读第一个）", variable=self.var_all_sheets
        ).grid(row=0, column=0, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(
            option_frame, text="加一列“来源文件”，方便追溯数据出处", variable=self.var_add_source
        ).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(
            option_frame, text="去掉整行重复的数据", variable=self.var_dedupe
        ).grid(row=2, column=0, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(
            option_frame, text="保留空行（默认删掉空行和空列）", variable=self.var_keep_empty
        ).grid(row=0, column=1, sticky="w", padx=10, pady=4)

        # 输出区
        out_frame = ttk.LabelFrame(self.root, text=" 3. 输出 ")
        out_frame.pack(fill="x", **pad)

        self.var_output = tk.StringVar(value="")
        out_entry = ttk.Entry(out_frame, textvariable=self.var_output)
        out_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        ttk.Button(out_frame, text="另存为…", command=self.choose_output).pack(side="left", padx=(0, 10))

        # 操作区
        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill="x", **pad)
        self.start_btn = ttk.Button(action_frame, text="开始合并", command=self.start)
        self.start_btn.pack(side="left")
        self.open_btn = ttk.Button(action_frame, text="打开结果文件夹", command=self.open_output_folder, state="disabled")
        self.open_btn.pack(side="left", padx=8)

        # 日志区
        log_frame = ttk.LabelFrame(self.root, text=" 处理日志 ")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log_frame, height=8, wrap="word", font=("Consolas", 9), state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=8)

    # ---------- 文件操作 ----------
    def add_files(self):
        selected = filedialog.askopenfilenames(title="选择要合并的表格", filetypes=SUPPORTED_TYPES)
        if not selected:
            return
        for path in selected:
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert(tk.END, Path(path).name)
        self._refresh_count()
        if not self.var_output.get() and self.files:
            # 默认存到程序所在目录，不落到示例数据文件夹里，免得找不到
            self.var_output.set(str(app_dir() / "合并结果.xlsx"))

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]
        self._refresh_count()

    def clear_files(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self._refresh_count()

    def _refresh_count(self):
        self.count_label.configure(text=f"已选 {len(self.files)} 个文件")

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="结果保存到", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")],
        )
        if path:
            self.var_output.set(path)

    # ---------- 执行 ----------
    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.files:
            messagebox.showwarning("还没选文件", "请先添加至少一个 Excel 或 CSV 文件。")
            return
        out_path = self.var_output.get().strip()
        if not out_path:
            messagebox.showwarning("还没设输出", "请先指定结果文件保存位置。")
            return

        self.start_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self._clear_log()
        self._log(f"开始处理，共 {len(self.files)} 个文件…")

        kwargs = dict(
            all_sheets=self.var_all_sheets.get(),
            add_source=self.var_add_source.get(),
            drop_duplicates=self.var_dedupe.get(),
            drop_empty_rows=not self.var_keep_empty.get(),
        )
        self.worker = threading.Thread(target=self._run_task, args=(list(self.files), out_path, kwargs), daemon=True)
        self.worker.start()

    def _run_task(self, files, out_path, kwargs):
        try:
            core.run(files, out_path, log=self._log, **kwargs)
            self.root.after(0, self._on_success, out_path)
        except Exception as e:  # 给用户看的错误信息要简洁
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, out_path):
        self.start_btn.configure(state="normal")
        self.open_btn.configure(state="normal")
        self.last_output = out_path
        messagebox.showinfo("完成", f"处理完成，结果已保存到：\n{out_path}")

    def _on_error(self, message):
        self.start_btn.configure(state="normal")
        self._log(f"处理失败：{message}")
        messagebox.showerror("处理失败", message)

    def open_output_folder(self):
        out_path = getattr(self, "last_output", "") or self.var_output.get()
        if not out_path or not Path(out_path).exists():
            messagebox.showinfo("提示", "还没有生成结果文件。")
            return
        folder = str(Path(out_path).parent)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showinfo("提示", f"打不开文件夹，请手动打开：\n{folder}\n({e})")

    # ---------- 日志 ----------
    def _log(self, message: str):
        """线程安全地往日志区写一行"""
        self.log_queue.put(message)

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _poll_log(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
