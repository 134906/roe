# app.py
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
from core import run_analysis

class RoeApp:
    def __init__(self, master):
        self.master = master
        master.title("ROE SCAN 分析")
        master.geometry("1000x750")

        # 设置网格权重，使日志区域可拉伸
        master.grid_rowconfigure(6, weight=1)
        master.grid_columnconfigure(1, weight=1)

        # ---- SPSS 文件选择 ----
        tk.Label(master, text="SPSS 数据文件 (.sav):").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.spss_entry = tk.Entry(master, width=60)
        self.spss_entry.grid(row=0, column=1, padx=5, sticky='ew')
        tk.Button(master, text="浏览", command=self.browse_spss).grid(row=0, column=2, padx=5)

        # ---- 定义文件选择 ----
        tk.Label(master, text="定义文件 (.xlsx):").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.def_entry = tk.Entry(master, width=60)
        self.def_entry.grid(row=1, column=1, padx=5, sticky='ew')
        tk.Button(master, text="浏览", command=self.browse_def).grid(row=1, column=2, padx=5)

        # ---- 输出路径（可选） ----
        tk.Label(master, text="输出路径 (可选):").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.out_entry = tk.Entry(master, width=60)
        self.out_entry.grid(row=2, column=1, padx=5, sticky='ew')
        self.out_entry.insert(0, "ROE_Results.xlsx")

        # ---- Workfile 模板文件（可选） ----
        tk.Label(master, text="Workfile 模板 (.xlsx):").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        self.wf_entry = tk.Entry(master, width=60)
        self.wf_entry.grid(row=3, column=1, padx=5, sticky='ew')
        tk.Button(master, text="浏览", command=self.browse_workfile).grid(row=3, column=2, padx=5)

        # ---- 新增：仅生成PPT复选框 ----
        self.ppt_only_var = tk.BooleanVar(value=False)
        self.ppt_only_cb = tk.Checkbutton(
            master,
            text="仅生成 PPT（使用已调整的 workfile）",
            variable=self.ppt_only_var
        )
        self.ppt_only_cb.grid(row=4, column=1, pady=2, sticky='w')

        # ---- 运行按钮 ----
        self.run_btn = tk.Button(master, text="运行分析", command=self.run, bg='lightblue', height=2, width=15)
        self.run_btn.grid(row=5, column=1, pady=15)

        # ---- 日志显示 ----
        self.log = scrolledtext.ScrolledText(master, height=30, width=100, wrap=tk.WORD)
        self.log.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

    def browse_spss(self):
        path = filedialog.askopenfilename(filetypes=[("SPSS files", "*.sav")])
        if path:
            self.spss_entry.delete(0, tk.END)
            self.spss_entry.insert(0, path)

    def browse_def(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.def_entry.delete(0, tk.END)
            self.def_entry.insert(0, path)

    def browse_workfile(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.wf_entry.delete(0, tk.END)
            self.wf_entry.insert(0, path)

    def log_message(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.master.update_idletasks()

    def run(self):
        spss = self.spss_entry.get()
        def_ = self.def_entry.get()
        out = self.out_entry.get().strip()
        wf = self.wf_entry.get().strip()
        ppt_only = self.ppt_only_var.get()

        if not out:
            out = "ROE_Results.xlsx"

        # 根据模式进行校验
        if ppt_only:
            if not wf or not os.path.exists(wf):
                messagebox.showerror("错误", "仅生成PPT模式需要选择有效的Workfile")
                return
        else:
            if not spss or not def_:
                messagebox.showerror("错误", "请选择SPSS数据文件和定义文件")
                return

        self.run_btn.config(state='disabled', text='运行中...')
        self.log.delete(1.0, tk.END)
        threading.Thread(target=self._run_thread, args=(spss, def_, out, wf, ppt_only), daemon=True).start()

    def _run_thread(self, spss, def_, out, wf='', ppt_only=False):
        # 重定向 print 输出到日志
        class PrintLogger:
            def __init__(self, callback):
                self.callback = callback
            def write(self, s):
                if s.strip():
                    self.callback(s.strip())
            def flush(self):
                pass

        sys.stdout = PrintLogger(self.log_message)
        try:
            workfile_path = wf if wf else None
            run_analysis(
                spss, def_, out,
                workfile_path=workfile_path,
                log_callback=self.log_message,
                ppt_only=ppt_only
            )
            if ppt_only:
                self.log_message("✅ PPT 生成成功！")
            else:
                self.log_message("✅ 所有分析任务成功完成！")
        except Exception as e:
            self.log_message(f"❌ 发生错误: {e}")
            import traceback
            self.log_message(traceback.format_exc())
        finally:
            self.master.after(0, lambda: self.run_btn.config(state='normal', text='运行分析'))
            sys.stdout = sys.__stdout__

if __name__ == "__main__":
    root = tk.Tk()
    app = RoeApp(root)
    root.mainloop()