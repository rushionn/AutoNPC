import os
from pathlib import Path
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, scrolledtext

class CommonUITemplate:
    def __init__(self, root):
        self.root = root

        # 預設存檔路徑
        self.default_path = Path("Z:\\")
        self.info_display_width = 40  # 路徑框寬度

        # UI初始化
        self.create_widgets()

    def create_widgets(self):
        # 按鈕區域
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=5, anchor='w')

        button_width = 10

        self.btn_check = ttk.Button(button_frame, text="檢查", width=button_width, bootstyle="success")
        self.btn_check.pack(side=LEFT, padx=5)

        self.btn_copy = ttk.Button(button_frame, text="複製", width=button_width, bootstyle="warning")
        self.btn_copy.pack(side=LEFT, padx=5)

        self.btn_compress = ttk.Button(button_frame, text="壓縮", width=button_width, bootstyle="primary")
        self.btn_compress.pack(side=LEFT, padx=5)

        self.btn_toggle_language = ttk.Button(button_frame, text="日本語へ切換", bootstyle="danger")
        self.btn_toggle_language.pack(side=LEFT, padx=5)

        # 勾選框區域
        self.checkbox_frame = ttk.Frame(self.root)
        self.checkbox_frame.pack(pady=5, anchor='w')

        self.split_var = ttk.BooleanVar(value=True)
        self.chk_split = ttk.Checkbutton(self.checkbox_frame, variable=self.split_var, text="分割", bootstyle="success")
        self.chk_split.pack(side=LEFT, padx=5)

        # 普通勾選框（可依需求擴充）
        self.checkboxes = {}
        for folder_name in ["Desktop", "Documents", "Downloads", "Pictures", "Favorites"]:
            var = ttk.BooleanVar(value=True)
            self.checkboxes[folder_name] = var
            chk = ttk.Checkbutton(self.checkbox_frame, text=folder_name, variable=var, bootstyle="success")
            chk.pack(side=LEFT, padx=5)

        # 特殊勾選框
        special_frame = ttk.Frame(self.root)
        special_frame.pack(pady=5, anchor='w')
        self.special_checkboxes = {}
        for key in ["bookmarks", "signatures", "constructor"]:
            var = ttk.BooleanVar(value=True)
            self.special_checkboxes[key] = var
            chk = ttk.Checkbutton(special_frame, text=key, variable=var, bootstyle="info")
            chk.pack(side=LEFT, padx=5)

        # 信息框-顯示選擇資料夾路徑
        self.info_frame = ttk.Frame(self.root)
        self.info_frame.pack(pady=10, anchor='w')
        self.info_display = ttk.Label(self.info_frame, text=self.default_path, bootstyle="light", font=('微軟正黑體', 14), anchor='w', width=self.info_display_width)
        self.info_display.pack(side=LEFT, padx=10)
        self.btn_choose_location = ttk.Button(self.info_frame, text="選擇存檔位置", bootstyle="success")
        self.btn_choose_location.pack(side=LEFT, padx=5)

        # 日誌顯示框
        self.action_display = scrolledtext.ScrolledText(
            self.root,
            height=15,
            width=60,
            bg="#263238",
            fg="white",
            font=('微軟正黑體', 12)
        )
        self.action_display.pack(pady=5, padx=10)

        # 進度條框架
        self.progress_frame = ttk.Frame(self.root)
        self.progress_frame.pack(pady=5, anchor='w')
        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar", thickness=15)
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=700, mode='determinate', style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(side=LEFT, padx=10, fill=X)
        self.progress_label = ttk.Label(self.progress_frame, text="", bootstyle="light", font=('微軟正黑體', 10))
        self.progress_label.pack(side=RIGHT, padx=3)

        # 關於按鈕
        self.btn_about = ttk.Button(button_frame, text="詳細", bootstyle="danger")
        self.btn_about.pack(side=RIGHT, padx=5)

# 範例主程式
if __name__ == "__main__":
    root = ttk.Window(themename="superhero")
    ui = CommonUITemplate(root)
    root.mainloop()