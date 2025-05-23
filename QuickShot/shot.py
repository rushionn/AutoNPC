import os
import time
import threading
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, END
import tkinter as tk
import pythoncom
from win32com.shell import shell, shellcon
import subprocess
import sys

LAST_PATH_FILE = "last_path.txt"

def open_lnk_target(lnk_path):
    pythoncom.CoInitialize()
    shell_link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None,
        pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
    )
    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Load(lnk_path)
    target_path, _ = shell_link.GetPath(shell.SLGP_UNCPRIORITY)
    arguments = shell_link.GetArguments()
    return target_path, arguments

def open_files_in_folder(folder_path, interval=4, log_func=None):
    # 只開啟有勾選的檔案
    selected_files = []
    for var, entry in checkbox_vars_entries:
        if var.get():
            selected_files.append(entry.get())
    files = selected_files if selected_files else [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    files.sort()
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            if file.lower().endswith('.lnk'):
                target, args = open_lnk_target(file_path)
                if target and os.path.exists(target):
                    if log_func:
                        log_func(f"Opening shortcut target: {target} {args}")
                    # 隱藏 cmd 視窗
                    subprocess.Popen(
                        f'"{target}" {args}',
                        shell=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    if log_func:
                        log_func(f"無法解析捷徑或目標不存在: {file_path}")
            else:
                if log_func:
                    log_func(f"Opening: {file_path}")
                os.startfile(file_path)
        except Exception as e:
            if log_func:
                log_func(f"無法開啟: {file_path}，錯誤：{e}")
        time.sleep(interval)

def start_opening():
    folder = folder_var.get()
    try:
        interval = float(interval_var.get())
    except ValueError:
        log("請輸入正確的間隔秒數")
        return
    if not os.path.isdir(folder):
        log("請選擇正確的資料夾")
        return
    log(f"開始開啟 {folder} 內的檔案，每 {interval} 秒一個")
    threading.Thread(target=open_files_in_folder, args=(folder, interval, log), daemon=True).start()

def choose_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_var.set(folder)
        save_last_path(folder)  # 新增：儲存選擇的路徑
        show_files_in_folder(folder)
        update_checkboxes(folder)

def show_files_in_folder(folder):
    files_listbox.delete(0, END)
    if not os.path.isdir(folder):
        return
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    for file in files:
        files_listbox.insert(END, file)

def log(msg):
    log_text.configure(state="normal")
    log_text.insert("end", msg + "\n")
    log_text.see("end")
    log_text.configure(state="disabled")

def load_last_path():
    if os.path.exists(LAST_PATH_FILE):
        with open(LAST_PATH_FILE, "r", encoding="utf-8") as f:
            path = f.read().strip()
            if os.path.isdir(path):
                return path
    return os.path.dirname(os.path.abspath(__file__))

def save_last_path(path):
    with open(LAST_PATH_FILE, "w", encoding="utf-8") as f:
        f.write(path)

app = tb.Window(themename="cosmo")
app.title("QuickShot - 批次開啟資料夾檔案")

# 初始化資料夾路徑
folder_var = tb.StringVar(value=load_last_path())
interval_var = tb.StringVar(value="4")

frm = tb.Frame(app, padding=8)
frm.pack(fill="both", expand=True)

tb.Label(frm, text="資料夾:").grid(row=0, column=0, sticky="w")
tb.Entry(frm, textvariable=folder_var, width=40).grid(row=0, column=1, padx=2, columnspan=6)
tb.Button(frm, text="選擇...", command=choose_folder, bootstyle=SECONDARY).grid(row=0, column=7, columnspan=2, padx=2)

tb.Label(frm, text="間隔秒數:").grid(row=1, column=0, sticky="w")
tb.Entry(frm, textvariable=interval_var, width=6).grid(row=1, column=1, sticky="w")  # 改為6

# 全選勾選框
select_all_var = tk.BooleanVar(value=True)
def select_all_changed():
    for var, _ in checkbox_vars_entries:
        var.set(select_all_var.get())
tb.Checkbutton(frm, text="全選", variable=select_all_var, command=select_all_changed).grid(row=1, column=2, sticky="w", padx=(8,0))

# === 勾選框外層大框框 ===
checkbox_frame = tb.Frame(frm, borderwidth=1, relief="solid", padding=4)
checkbox_frame.grid(row=2, column=0, columnspan=8, sticky="w", pady=(4, 4))

checkbox_vars_entries = []
for i in range(10):
    var = tk.BooleanVar(value=False)
    entry = tb.Entry(checkbox_frame, width=22, state="readonly")  # 設為唯讀
    row = i % 5
    col = 0 if i < 5 else 3
    chk = tb.Checkbutton(checkbox_frame, variable=var, width=2)
    chk.grid(row=row, column=col, sticky="w", padx=1, pady=1)
    tb.Label(checkbox_frame, text=str(i+1), width=2).grid(row=row, column=col+1, sticky="w", padx=1)
    entry.grid(row=row, column=col+2, padx=1, pady=1)
    checkbox_vars_entries.append((var, entry))

# 檔案清單
tb.Label(frm, text="檔案清單:").grid(row=8, column=0, sticky="nw")
files_listbox = tk.Listbox(frm, height=8, width=60)
files_listbox.grid(row=8, column=1, columnspan=6, pady=2, sticky="w")

# 啟動按鈕移到檔案清單右邊
tb.Button(frm, text="啟動", command=start_opening, bootstyle=SUCCESS).grid(row=8, column=7, padx=(8,0), pady=2, sticky="w")

# 動態顯示（原日誌）
tb.Label(frm, text="動態顯示:").grid(row=9, column=0, sticky="nw")
log_text = tb.Text(frm, height=8, width=80, state="disabled")
log_text.grid(row=9, column=1, columnspan=7, pady=2, sticky="w")

def on_folder_var_change(*args):
    show_files_in_folder(folder_var.get())
    update_checkboxes(folder_var.get())

def update_checkboxes(folder):
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    for i in range(10):
        var, entry = checkbox_vars_entries[i]
        entry.config(state="normal")
        if i < len(files):
            entry.delete(0, END)
            entry.insert(0, files[i])
            var.set(True)
        else:
            entry.delete(0, END)
            var.set(False)
        entry.config(state="readonly")

# 初始化時顯示檔案清單與勾選框
show_files_in_folder(folder_var.get())
update_checkboxes(folder_var.get())
folder_var.trace_add("write", on_folder_var_change)

app.mainloop()